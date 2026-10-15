#!/usr/bin/env python3
"""Read-only structural inspection of think-cell metadata in PPTX files.

Independent stdlib implementation. Does not execute Office, extract files, edit
presentations, or claim import compatibility. XML names are observations, not
a complete specification of think-cell's private persistence format.
"""
import argparse
import collections
import hashlib
import io
import json
import math
import pathlib
import posixpath
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

MAX_FILE = 128 * 1024 * 1024
MAX_PART = 32 * 1024 * 1024
MAX_TOTAL = 128 * 1024 * 1024
MAX_ENTRIES = 10000
MAX_NODES = 200000
MAX_DEPTH = 128
FREE = 0xFFFFFFFF
END = 0xFFFFFFFE
SPECIAL = 0xFFFFFFFA
CFB_MAGIC = bytes.fromhex('d0cf11e0a1b11ae1')
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
C = 'http://schemas.openxmlformats.org/drawingml/2006/chart'


class InspectionError(ValueError):
    pass


def digest(data):
    return hashlib.sha256(data).hexdigest()


def local(tag):
    return tag.rsplit('}', 1)[-1]


def parse_xml(data):
    """No DTD/entity declarations; bounded bytes, element count and depth."""
    size = len(data.encode('utf-8')) if isinstance(data, str) else len(data)
    if size > MAX_PART:
        raise InspectionError('XML part exceeds size limit')
    # Expat does not fetch external entities, and declarations are explicitly
    # refused, including UTF-16/32 byte encodings with interleaved NUL bytes.
    probe = data if isinstance(data, str) else data.replace(b'\x00', b'').decode('latin1')
    if '<!DOCTYPE' in probe.upper() or '<!ENTITY' in probe.upper():
        raise InspectionError('DTD and entity declarations are not supported')
    try:
        root = ET.fromstring(data)
    except (ET.ParseError, ValueError) as exc:
        raise InspectionError('Invalid XML: ' + str(exc)) from exc
    stack = [(root, 1)]
    count = 0
    while stack:
        node, depth = stack.pop()
        count += 1
        if count > MAX_NODES or depth > MAX_DEPTH:
            raise InspectionError('XML structure exceeds complexity limit')
        stack.extend((child, depth + 1) for child in node)
    return root


class SafeZip:
    def __init__(self, payload):
        if len(payload) > MAX_FILE:
            raise InspectionError('ZIP container exceeds size limit')
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(payload))
        except zipfile.BadZipFile as exc:
            raise InspectionError('Not a valid ZIP package') from exc
        infos = self.zip.infolist()
        if len(infos) > MAX_ENTRIES:
            self.zip.close()
            raise InspectionError('Too many ZIP entries')
        self.infos = {}
        total = 0
        for item in infos:
            name = item.filename
            if name in self.infos:
                raise InspectionError('Duplicate ZIP entry: ' + name)
            if name.startswith(('/', '\\')) or '\\' in name or '..' in name.split('/') or ':' in name:
                raise InspectionError('Unsafe ZIP part name: ' + name)
            if item.flag_bits & 1:
                raise InspectionError('Encrypted ZIP parts are unsupported')
            if item.file_size > MAX_PART:
                raise InspectionError('ZIP entry exceeds size limit: ' + name)
            total += item.file_size
            if total > MAX_TOTAL:
                raise InspectionError('ZIP expanded size exceeds limit')
            self.infos[name] = item
        self.read_total = 0

    def read(self, name):
        item = self.infos[name]
        try:
            with self.zip.open(item) as stream:
                data = stream.read(MAX_PART + 1)
        except (zipfile.BadZipFile, RuntimeError, NotImplementedError) as exc:
            raise InspectionError('Cannot read ZIP part: ' + name) from exc
        if len(data) != item.file_size or len(data) > MAX_PART:
            raise InspectionError('ZIP size mismatch: ' + name)
        self.read_total += len(data)
        if self.read_total > MAX_TOTAL:
            raise InspectionError('ZIP read budget exceeded')
        return data

    def close(self):
        self.zip.close()


class CompoundFile:
    """CFB stream reader with FAT, DIFAT and miniFAT support; no file writes."""
    def __init__(self, data):
        self.data = data
        if len(data) < 512 or len(data) > MAX_PART or data[:8] != CFB_MAGIC:
            raise InspectionError('Invalid or oversized CFB header')
        self.major = struct.unpack_from('<H', data, 26)[0]
        byteorder, shift, mini_shift = struct.unpack_from('<HHH', data, 28)
        if byteorder != 0xFFFE or (self.major, shift) not in ((3, 9), (4, 12)) or mini_shift != 6:
            raise InspectionError('Unsupported CFB version or sector size')
        self.sector_size = 1 << shift
        self.mini_size = 1 << mini_shift
        if len(data) % self.sector_size:
            raise InspectionError('CFB file is not sector-aligned')
        self.sector_count = len(data) // self.sector_size - 1
        fat_count = self.u32(44)
        self.cutoff = self.u32(56)
        if self.cutoff != 4096 or fat_count > self.sector_count:
            raise InspectionError('Invalid CFB allocation header')
        fat_sectors = [x for x in struct.unpack_from('<109I', data, 76) if x != FREE]
        difat, difat_count = self.u32(68), self.u32(72)
        if difat_count > self.sector_count:
            raise InspectionError('Invalid DIFAT count')
        seen = set()
        for _ in range(difat_count):
            if difat in seen:
                raise InspectionError('Cyclic DIFAT chain')
            seen.add(difat)
            values = self.words(self.sector(difat))
            fat_sectors.extend(x for x in values[:-1] if x != FREE)
            difat = values[-1]
        if difat_count and difat != END:
            raise InspectionError('Unterminated DIFAT chain')
        if len(fat_sectors) != fat_count or len(set(fat_sectors)) != len(fat_sectors):
            raise InspectionError('CFB FAT count mismatch or duplicate FAT sector')
        self.fat = []
        for sid in fat_sectors:
            self.fat.extend(self.words(self.sector(sid)))
        directory = self.chain(self.u32(48), self.fat, self.sector, self.sector_count)
        if not directory or len(directory) % 128:
            raise InspectionError('Invalid CFB directory')
        if len(directory) // 128 > MAX_ENTRIES:
            raise InspectionError('Too many CFB directory entries')
        self.entries = []
        for offset in range(0, len(directory), 128):
            part = directory[offset:offset + 128]
            n = struct.unpack_from('<H', part, 64)[0]
            typ = part[66]
            if typ not in (0, 1, 2, 5):
                raise InspectionError('Invalid CFB directory entry type')
            if typ and (n < 2 or n > 64 or n % 2 or part[n - 2:n] != b'\0\0'):
                raise InspectionError('Invalid CFB directory name')
            try:
                name = part[:max(0, n - 2)].decode('utf-16le')
            except UnicodeError as exc:
                raise InspectionError('Invalid CFB UTF-16 name') from exc
            size = struct.unpack_from('<Q', part, 120)[0]
            if self.major == 3:
                size &= 0xFFFFFFFF
            if size > MAX_PART:
                raise InspectionError('CFB stream exceeds size limit')
            self.entries.append({'name': name, 'type': typ, 'size': size,
                'left': struct.unpack_from('<I', part, 68)[0],
                'right': struct.unpack_from('<I', part, 72)[0],
                'child': struct.unpack_from('<I', part, 76)[0],
                'start': struct.unpack_from('<I', part, 116)[0]})
        root = self.entries[0]
        if root['type'] != 5:
            raise InspectionError('Missing CFB root storage')
        mini_count = self.u32(64)
        if mini_count > self.sector_count:
            raise InspectionError('Invalid miniFAT count')
        mini_data = self.chain(self.u32(60), self.fat, self.sector, self.sector_count,
                               mini_count * self.sector_size)
        self.mini_fat = self.words(mini_data)
        self.mini_stream = self.chain(root['start'], self.fat, self.sector, self.sector_count, root['size'])
        self.streams = {}
        todo = [(root['child'], '', 0)]
        visited = set()
        while todo:
            index, parent, depth = todo.pop()
            if index == FREE:
                continue
            if index >= len(self.entries) or index in visited or depth > MAX_DEPTH:
                raise InspectionError('Invalid or cyclic CFB directory tree')
            visited.add(index)
            entry = self.entries[index]
            if entry['type'] not in (1, 2):
                raise InspectionError('Unexpected CFB directory member')
            name = entry['name']
            if '/' in name or '\\' in name or name in ('.', '..'):
                raise InspectionError('Invalid CFB member name')
            path = parent + '/' + name if parent else name
            todo.extend([(entry['left'], parent, depth), (entry['right'], parent, depth)])
            if entry['type'] == 1:
                todo.append((entry['child'], path, depth + 1))
            else:
                if path.casefold() in {p.casefold() for p in self.streams}:
                    raise InspectionError('Duplicate CFB stream path')
                self.streams[path] = entry

    def u32(self, offset):
        return struct.unpack_from('<I', self.data, offset)[0]

    @staticmethod
    def words(data):
        return list(struct.unpack('<%dI' % (len(data) // 4), data))

    def sector(self, sid):
        if sid >= self.sector_count:
            raise InspectionError('CFB sector outside file')
        offset = (sid + 1) * self.sector_size
        return self.data[offset:offset + self.sector_size]

    @staticmethod
    def chain(start, table, read, maximum, size=None):
        if size == 0:
            if start not in (END, FREE, 0):
                raise InspectionError('Unexpected allocation for empty CFB stream')
            return b''
        sid = start
        seen = set()
        blocks = []
        total = 0
        while sid != END:
            if sid >= SPECIAL or sid >= len(table) or sid in seen or len(seen) >= maximum:
                raise InspectionError('Invalid, oversized or cyclic CFB chain')
            seen.add(sid)
            block = read(sid)
            if not block:
                raise InspectionError('Empty CFB sector')
            blocks.append(block)
            total += len(block)
            if total > MAX_PART:
                raise InspectionError('CFB chain exceeds size limit')
            if size is not None and total - len(block) >= size:
                raise InspectionError('Overallocated CFB stream chain')
            sid = table[sid]
        if size is not None and total < size:
            raise InspectionError('Truncated CFB stream chain')
        result = b''.join(blocks)
        return result if size is None else result[:size]

    def read(self, path):
        entry = self.streams[path]
        if entry['size'] < self.cutoff:
            def mini(sid):
                offset = sid * self.mini_size
                if offset + self.mini_size > len(self.mini_stream):
                    raise InspectionError('Mini-sector outside root mini-stream')
                return self.mini_stream[offset:offset + self.mini_size]
            return self.chain(entry['start'], self.mini_fat, mini, len(self.mini_fat), entry['size'])
        return self.chain(entry['start'], self.fat, self.sector, self.sector_count, entry['size'])


def var_value(element):
    if element is None:
        return None
    typ = element.get('type')
    result = {'type_code': typ}
    if typ == '1':
        literal = element.get('val')
        result['numeric_literal'] = literal
        try:
            value = float(literal)
            if not math.isfinite(value):
                raise ValueError()
            result['number'] = value
        except (TypeError, ValueError):
            result['warning'] = 'Invalid/non-finite numeric value retained as literal'
    elif typ == '5':
        result['string'] = element.text or ''
    elif typ == '6':
        dt = element.find('m_datetime')
        result['date_literal'] = dt.get('val') if dt is not None else None
        result['string_fallback'] = element.findtext('m_str')
    else:
        result['attributes'] = dict(element.attrib)
        result['text'] = element.text
        result['warning'] = 'Unmapped variable type; not interpreted'
    return result


def model_summary(root):
    ids = {}
    for element in root:
        key = element.get('id')
        if key is not None:
            if key in ids:
                raise InspectionError('Duplicate model object id: ' + key)
            ids[key] = element
    unresolved = sorted({e.get('idref') for e in root.iter()
                         if e.get('idref') not in (None, '0') and e.get('idref') not in ids})
    warnings = []

    def ref(element):
        return ids.get(element.get('idref')) if element is not None else None

    def links(element, path):
        collection = element.find(path)
        if collection is None:
            return []
        count = collection.get('length')
        if count is not None and count != str(len(collection)):
            warnings.append('Collection length mismatch: ' + element.tag + '/' + path)
        return [ref(e) for e in collection]

    def source(element, path):
        target = ref(element.find(path))
        if target is None:
            return None
        value = var_value(target.find('m_varval'))
        if value is None:
            return {'source_object_id':target.get('id'),'source_object_type':target.tag,
                    'value_status':'No persisted m_varval; derived or unknown source not evaluated'}
        value['source_object_id'] = target.get('id')
        value['source_object_type'] = target.tag
        return value

    def attr(element, path, key='val'):
        node = element.find(path)
        return node.get(key) if node is not None else None

    charts = []
    for element in root:
        if element.tag not in ('CSequenceChartSE', 'CScatterChartSE'):
            continue
        table = ref(element.find('m_dtable'))
        chart = {'object_type': element.tag, 'object_id': element.get('id'),
                 'automation_name': element.findtext('m_strName') or None,
                 'chart_type_code': attr(element, 'm_ect'),
                 'orientation_code': attr(element, 'm_eorient'),
                 'default_series_type_code': attr(element, 'm_estDefault'),
                 'layout_anchor_ids': [x.get('idref') for x in element.findall('m_agrdlnanchor/elem')],
                 'plot_gridline_bindings': {side: attr(element, 'm_rectgrdlnanchorPlot/' + side + '/m_grdlnBinding', 'idref') for side in ['left','top','right','bottom']}}
        geometry = ref(element.find('m_pptseqchart'))
        if geometry is None:
            geometry = element.find('m_pptscatchart')
        if geometry is not None:
            rect = geometry.find('m_rectPPTShape')
            chart['rendered_shape'] = {'token': geometry.findtext('m_bstrShapeName'),
                'model_rect_raw': dict(rect.attrib) if rect is not None else None,
                'unit_status': 'Raw private-model coordinates; conversion not assumed'}
        if table is None:
            chart['warning'] = 'Missing data-table reference'
            charts.append(chart)
            continue
        storage = attr(table, 'm_bstrRangeName')
        chart.update({'table_object_id': table.get('id'), 'table_type': table.tag,
            'table_automation_name': table.findtext('m_strName') or None,
            'external_storage_flag': attr(table, 'm_bExternalStorage'),
            'child_storage_name': storage,
            'child_package_stream': storage + '/Package' if storage else None})
        if element.tag == 'CSequenceChartSE':
            series = links(table, 'm_cscdser')
            chart['series'] = [{'object_id': s.get('id'), 'label': source(s, 'm_varsrc')} if s is not None else None for s in series]
            chart['categories'] = []
            for vector in links(table, 'ocol'):
                if vector is None:
                    chart['categories'].append({'warning':'Missing category reference'})
                    continue
                category = {'object_id': vector.get('id'),
                    'label': var_value(vector.find('m_varsrcCategory/m_varval')),
                    'datasheet_index': attr(vector, 'm_nIndexInDataSheet'),
                    'equals_scalar_ref': attr(vector, 'm_scdscEquals', 'idref'), 'values':[]}
                scalars = links(vector, 'ocol')
                if len(scalars) != len(series):
                    warnings.append('Scalar/series count mismatch: category ' + vector.get('id','?'))
                for scalar in scalars:
                    if scalar is None:
                        category['values'].append({'warning':'Missing scalar reference'})
                    else:
                        category['values'].append({'object_id': scalar.get('id'),
                            'absolute': source(scalar, 'm_varsrcAbsolute'),
                            'relative': source(scalar, 'm_varsrcRelative'),
                            'accumulated_absolute': source(scalar, 'm_varsrcAccuAbsolute'),
                            'connection_ref': attr(scalar, 'm_scdscconnect', 'idref')})
                chart['categories'].append(category)
            chart['vector_connection_refs'] = [x.get('idref') for x in table.findall('m_cscdveccon/elem')]
            chart['equals_status'] = 'Reference exposed only; total/e semantics unverified without waterfall samples'
        else:
            chart['series'] = [{'object_id':s.get('id'),'label':source(s,'m_varsrc')} if s is not None else None for s in links(table,'m_cscatdseries')]
            chart['points'] = []
            for vector in links(table, 'ocol'):
                if vector is None:
                    chart['points'].append({'warning':'Missing point reference'})
                else:
                    chart['points'].append({'object_id':vector.get('id'),
                        **{key:source(vector,path) for key,path in [('label','m_varsrcText'),('group','m_varsrcGroup'),('x','m_varsrcX'),('y','m_varsrcY'),('size','m_varsrcSize')]}})
            chart['group_status'] = 'Raw group cells retained; blank-cell inheritance not inferred'
        charts.append(chart)
    field_names = [e.text for e in root.findall('.//m_strName') if e.text]
    return {'root':root.tag, 'attributes':root.attrib,
        'version':root.find('version').attrib if root.find('version') is not None else None,
        'object_count':len(ids),'object_types':dict(collections.Counter(e.tag for e in ids.values())),
        'unresolved_idrefs':unresolved,'automation_names':field_names,'charts':charts,'warnings':warnings}


def relation_source(name):
    if name == '_rels/.rels':
        return ''
    return posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])


def inventory_nested_zip(payload):
    archive = SafeZip(payload)
    try:
        entries = [{'path':name,'bytes':info.file_size} for name,info in archive.infos.items()]
        return {'format':'XLSB' if 'xl/workbook.bin' in archive.infos else 'ZIP', 'entries':entries,
                'data_status':'Binary Excel worksheet records not interpreted'}
    finally:
        archive.close()


def inspect_bytes(payload, filename='<memory>'):
    archive = SafeZip(payload)
    result = {'schema_version':1,'input':{'name':filename,'bytes':len(payload),'sha256':digest(payload)},
        'mode':'read-only static inspection; not import, editing or round-trip compatibility',
        'entries':[],'namespaces':[],'relationships':[],'tags':[],'ole_objects':[],
        'compound_files':[],'native_charts':[],'text_targets':[],'shapes':[],
        'limitations':['Only explicit complete think-cellXML streams are parsed, without scanning binary data.',
            'Private enum codes retained; type semantics require validation.',
            'No waterfall/Gantt/Mekko reconstruction from unobserved schema.',
            'Embedded XLSB data and opaque OLE streams are not decoded.',
            'No Office execution, no documents changed, no extracted assets.']}
    namespaces = set()
    try:
        for name,info in archive.infos.items():
            data = archive.read(name)
            result['entries'].append({'path':name,'bytes':len(data),'sha256':digest(data)})
            if name.endswith(('.xml','.rels')):
                root = parse_xml(data)
                namespaces.update(s[1:].split('}')[0] for e in root.iter() for s in [e.tag]+list(e.attrib) if s.startswith('{'))
                if name.endswith('.rels'):
                    owner = relation_source(name)
                    for e in root:
                        target = e.get('Target','')
                        resolved = None if e.get('TargetMode') == 'External' else posixpath.normpath(posixpath.join(posixpath.dirname(owner),target))
                        result['relationships'].append({'part':name,'owner':owner,**e.attrib,'resolved_target':resolved})
                for e in root.iter():
                    if e.tag == '{'+P+'}tag':
                        value = e.get('val','')
                        item = {'part':name,'name':e.get('name'),'value_length':len(value),'value_sha256':digest(value.encode())}
                        if len(value) <= 100:
                            item['value'] = value
                        if e.get('name') == 'THINKCELLPRESENTATIONDONOTDELETE' and value.startswith('<?xml'):
                            item['presentation_metadata'] = model_summary(parse_xml(value))
                        result['tags'].append(item)
                    if e.tag == '{'+P+'}oleObj':
                        result['ole_objects'].append({'part':name,**e.attrib})
                    if e.tag == '{'+A+'}fld' and e.get('type','').startswith('thinkcell<?xml'):
                        value = e.get('type')[len('thinkcell'):]
                        result['text_targets'].append({'part':name,'field_id':e.get('id'),
                            'attribute':'a:fld/@type','prefix':'thinkcell','model':model_summary(parse_xml(value))})
                    if e.tag in ('{'+P+'}sp','{'+P+'}graphicFrame','{'+P+'}pic','{'+P+'}cxnSp','{'+P+'}grpSp'):
                        nv = next((child for child in e if local(child.tag).startswith('nv')),None)
                        prop = nv.find('{'+P+'}cNvPr') if nv is not None else None
                        if prop is not None:
                            xfrm = e.find('{'+P+'}xfrm')
                            if xfrm is None:
                                xfrm = e.find('{'+P+'}spPr/{'+A+'}xfrm')
                            result['shapes'].append({'part':name,'kind':local(e.tag),**prop.attrib,
                                'tag_rids':[tag.get('{'+R+'}id') for tag in nv.iter('{'+P+'}tags')],
                                'chart_rids':[chart.get('{'+R+'}id') for chart in e.iter('{'+C+'}chart')],
                                'transform_emu':{local(child.tag):child.attrib for child in xfrm} if xfrm is not None else None})
                if name.startswith('ppt/charts/') and name.endswith('.xml'):
                    result['native_charts'].append({'part':name,
                        'types':[local(e.tag) for e in root.iter() if local(e.tag).endswith('Chart')],
                        'grouping':[e.get('val') for e in root.iter('{'+C+'}grouping')],
                        'bar_direction':[e.get('val') for e in root.iter('{'+C+'}barDir')],
                        'series_count':len(list(root.iter('{'+C+'}ser')))})
            if name.startswith('ppt/embeddings/') and data[:8] == CFB_MAGIC:
                cfb = CompoundFile(data)
                container = {'part':name,'major_version':cfb.major,'sector_size':cfb.sector_size,'streams':[]}
                read_budget = 0
                for stream in cfb.streams:
                    content = cfb.read(stream)
                    read_budget += len(content)
                    if read_budget > MAX_TOTAL:
                        raise InspectionError('Compound stream budget exceeded')
                    item = {'path':stream,'bytes':len(content),'sha256':digest(content)}
                    if stream.split('/')[-1] == 'think-cellXML':
                        if not content.startswith(b'<?xml'):
                            raise InspectionError('Unexpected think-cellXML stream envelope')
                        item['envelope'] = {'xml_offset':0,'xml_bytes':len(content),'encoding_observed':'UTF-8' if b'encoding="UTF-8"' in content[:100] else 'unknown','signature':content[:51].decode('ascii',errors='replace')}
                        item['model'] = model_summary(parse_xml(content))
                    elif stream.split('/')[-1] == 'Package' and content.startswith(b'PK\x03\x04'):
                        item['nested_package'] = inventory_nested_zip(content)
                    else:
                        item['status'] = 'Opaque stream not decoded'
                    container['streams'].append(item)
                result['compound_files'].append(container)
        result['namespaces'] = sorted(namespaces)
        relation_map = {(r['owner'],r.get('Id')):r.get('resolved_target') for r in result['relationships']}
        tag_values = collections.defaultdict(set)
        for tag in result['tags']:
            if tag['name'] == 'THINKCELLSHAPEDONOTDELETE' and tag.get('value'):
                tag_values[tag['part']].add(tag['value'])
        native = {chart['part']:chart for chart in result['native_charts']}
        shape_by_token = collections.defaultdict(list)
        for shape in result['shapes']:
            for rid in shape['tag_rids']:
                for token in tag_values.get(relation_map.get((shape['part'],rid)),set()):
                    shape_by_token[token].append({'slide_part':shape['part'],'shape_id':shape.get('id'),
                        'shape_name':shape.get('name'),'transform_emu':shape['transform_emu'],
                        'native_charts':[native.get(relation_map.get((shape['part'],chart_rid))) for chart_rid in shape['chart_rids']]})
        for container in result['compound_files']:
            for stream in container['streams']:
                for chart in stream.get('model',{}).get('charts',[]):
                    token = chart.get('rendered_shape',{}).get('token')
                    chart['resolved_shape_links'] = shape_by_token.get(token,[])
        result['counts'] = {'parts':len(result['entries']),'native_charts':len(result['native_charts']),
            'thinkcell_chart_models':sum(len(s.get('model',{}).get('charts',[])) for c in result['compound_files'] for s in c['streams']),
            'text_targets':len(result['text_targets']), 'custom_xml_parts':sum(n.startswith('customXml/') and n.endswith('.xml') for n in archive.infos)}
        return result
    finally:
        archive.close()


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('presentation',type=pathlib.Path)
    parser.add_argument('--json',dest='output',type=pathlib.Path,help='Write inspection JSON here; original stays unchanged')
    args = parser.parse_args(argv)
    try:
        if args.output and args.output.resolve() == args.presentation.resolve():
            raise InspectionError('Output must differ from input presentation')
        if args.presentation.stat().st_size > MAX_FILE:
            raise InspectionError('Input exceeds size limit')
        result = inspect_bytes(args.presentation.read_bytes(),args.presentation.name)
        serialized = json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False)
        if args.output:
            args.output.write_text(serialized,encoding='utf-8')
            print('Read-only report: ' + str(args.output))
        else:
            print(serialized)
        return 0
    except (InspectionError,OSError,zipfile.BadZipFile) as exc:
        print('Inspection failed: ' + str(exc),file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
