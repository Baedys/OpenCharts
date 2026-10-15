"""Synthetic fixtures only: no proprietary presentations or payloads."""
import importlib.util
import io
import pathlib
import struct
import unittest
import warnings
import zipfile

MODULE = pathlib.Path(__file__).resolve().parents[1] / 'tools' / 'inspect_thinkcell.py'
spec = importlib.util.spec_from_file_location('inspect_thinkcell', MODULE)
reader = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reader)


def pptx(parts):
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts:
            archive.writestr(name, content)
    return output.getvalue()


def directory_entry(name, kind, start=reader.END, size=0, child=reader.FREE):
    data = bytearray(128)
    encoded = name.encode('utf-16le') + b'\0\0'
    data[:len(encoded)] = encoded
    struct.pack_into('<H', data, 64, len(encoded))
    data[66:68] = bytes([kind, 1])
    struct.pack_into('<III', data, 68, reader.FREE, reader.FREE, child)
    struct.pack_into('<IQ', data, 116, start, size)
    return bytes(data)


def cfb(payload, mini=False):
    """Original small CFB v3 fixture with one named stream."""
    header = bytearray(512)
    header[:8] = reader.CFB_MAGIC
    struct.pack_into('<HHHHH',header,24,0x3e,3,0xfffe,9,6)
    struct.pack_into('<IIIIIIIII',header,40,0,1,1,0,4096,3 if mini else reader.END,1 if mini else 0,reader.END,0)
    struct.pack_into('<109I',header,76,0,*([reader.FREE]*108))
    fat = [reader.FREE]*128
    fat[0] = 0xfffffffd
    fat[1] = reader.END
    if mini:
        if len(payload)>64:
            raise ValueError('Fixture mini stream max 64 bytes')
        fat[2] = fat[3] = reader.END
        root = directory_entry('Root Entry',5,start=2,size=64,child=1)
        stream = directory_entry('think-cellXML',2,start=0,size=len(payload))
        mini_fat = struct.pack('<128I',reader.END,*([reader.FREE]*127))
        blocks = [payload.ljust(512,b'\0'),mini_fat]
    else:
        content=payload.ljust(4096,b' ')
        if len(content)!=4096:
            raise ValueError('Fixture stream max 4096 bytes')
        for i in range(2,9):
            fat[i]=i+1
        fat[9]=reader.END
        root=directory_entry('Root Entry',5,child=1)
        stream=directory_entry('think-cellXML',2,start=2,size=4096)
        blocks=[content[i:i+512] for i in range(0,4096,512)]
    return bytes(header)+struct.pack('<128I',*fat)+(root+stream).ljust(512,b'\0')+b''.join(blocks)


MODEL = b'''<?xml version="1.0" encoding="UTF-8"?><root reqver="1"><version val="1"/>
<CSequenceChartSE id="1"><m_strName>Revenue</m_strName><m_ect val="0"/><m_dtable idref="2"/></CSequenceChartSE>
<CSequenceChartDataTable id="2"><m_bstrRangeName val="think-cellChild0"/><m_cscdser length="1"><elem idref="3"/></m_cscdser><ocol length="1"><elem idref="4"/></ocol></CSequenceChartDataTable>
<CSequenceChartDataSeries id="3"><m_varsrc idref="5"/></CSequenceChartDataSeries>
<CSequenceChartDataVector id="4"><m_varsrcCategory><m_varval type="5">Q1</m_varval></m_varsrcCategory><ocol length="1"><elem idref="6"/></ocol><m_scdscEquals idref="0"/></CSequenceChartDataVector>
<CVariableSource id="5"><m_varval type="5">Sales</m_varval></CVariableSource>
<CSequenceChartDataScalar id="6"><m_varsrcAbsolute idref="7"/><m_scdscconnect idref="0"/></CSequenceChartDataScalar>
<CVariableSource id="7"><m_varval type="1" val="1.25E+01"/></CVariableSource></root>'''


class AppearanceMetadataTests(unittest.TestCase):
    def test_group_transform_shape_flags_and_body_properties_retained(self):
        xml=('''<p:sld xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree><p:grpSp>
          <p:nvGrpSpPr><p:cNvPr id="2" name="Group"/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="60000"><a:off x="100" y="200"/><a:ext cx="300" cy="400"/><a:chOff x="0" y="0"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>
          <p:sp><p:nvSpPr><p:cNvPr id="3" name="Text"/></p:nvSpPr><p:spPr><a:xfrm flipH="1"><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr>
          <p:txBody><a:bodyPr anchor="ctr" lIns="0" tIns="0" rIns="0" bIns="0"/><a:p><a:r><a:t>Invented</a:t></a:r></a:p></p:txBody></p:sp>
          </p:grpSp></p:spTree></p:cSld></p:sld>'''%(reader.P,reader.A)).encode()
        result=reader.inspect_bytes(pptx([('ppt/slides/slide1.xml',xml)]))
        shape=next(s for s in result['shapes'] if s['id']=='3')
        self.assertEqual(shape['transform_attributes'],{'flipH':'1'})
        self.assertEqual(shape['parent_group_transforms'][0]['attributes'],{'rot':'60000'})
        self.assertEqual(shape['parent_group_transforms'][0]['coordinates']['off'],{'x':'100','y':'200'})
        self.assertEqual(shape['body_properties']['anchor'],'ctr');self.assertEqual(shape['body_properties']['lIns'],'0')

    def test_native_cache_labels_axis_and_gap(self):
        root=reader.parse_xml(('''<c:chartSpace xmlns:c="%s" xmlns:a="%s"><c:chart><c:plotArea>
          <c:barChart><c:barDir val="col"/><c:grouping val="stacked"/><c:gapWidth val="60"/><c:overlap val="100"/>
          <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Invented</c:v></c:tx>
            <c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></c:spPr>
            <c:val><c:numLit><c:pt idx="0"><c:v>7.25</c:v></c:pt></c:numLit></c:val>
            <c:dLbls><c:dLbl><c:idx val="0"/><c:dLblPos val="ctr"/><c:showVal val="1"/><c:numFmt formatCode="0"/>
              <c:txPr><a:p><a:pPr><a:defRPr sz="1100"><a:latin typeface="Synthetic"/></a:defRPr></a:pPr></a:p></c:txPr>
            </c:dLbl></c:dLbls></c:ser></c:barChart>
          <c:valAx><c:scaling><c:min val="0"/><c:max val="8"/></c:scaling><c:tickLblPos val="none"/>
            <c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:valAx>
        </c:plotArea></c:chart></c:chartSpace>'''%(reader.C,reader.A)).encode())
        result=reader.native_chart_metadata(root,'ppt/charts/chart2.xml')
        self.assertEqual(result['series'][0]['gap_width'],60)
        self.assertEqual(result['series'][0]['cache_values'],[{'index':'0','literal':'7.25'}])
        self.assertEqual(result['series'][0]['labels'][0]['font']['size_points'],11)
        self.assertEqual(result['axes'][0]['maximum'],8)
        self.assertTrue(result['axes'][0]['line_hidden']);self.assertFalse(result['has_legend'])

    def test_color_transforms_not_guessed(self):
        theme={'mapping':{'tx1':'dk1'},'colors':{'dk1':{'kind':'srgbClr','val':'121314','transforms':[]}}}
        self.assertEqual(reader.theme_color(theme,{'kind':'schemeClr','val':'tx1','transforms':[]}),'#121314')
        self.assertIsNone(reader.theme_color(theme,{'kind':'schemeClr','val':'tx1','transforms':[{'kind':'tint','val':'20000'}]}))

    def test_effective_run_font_overrides_list_size_preserves_theme_color(self):
        root=reader.parse_xml(('''<p:sp xmlns:p="%s" xmlns:a="%s"><p:txBody><a:lstStyle><a:lvl1pPr><a:defRPr sz="2600"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></a:lstStyle><a:p><a:pPr lvl="0"/><a:fld><a:rPr sz="1200"/><a:t>Synthetic</a:t></a:fld></a:p></p:txBody></p:sp>'''%(reader.P,reader.A)).encode())
        font=reader.shape_fonts(root)[0]
        self.assertEqual(font['size_points'],12);self.assertEqual(font['latin'],'+mn-lt');self.assertEqual(font['color']['val'],'tx1')

    def test_labels_follow_owned_model_refs(self):
        root=reader.parse_xml(MODEL)
        import xml.etree.ElementTree as ET
        series=root.find('CSequenceChartDataSeries')
        ET.SubElement(series,'m_scdserlabel',{'idref':'8'})
        label=ET.SubElement(root,'CSequenceChartDataSeriesLabel',{'id':'8'})
        box=ET.SubElement(label,'m_ppttb');ET.SubElement(box,'m_bstrShapeName').text='invented-token'
        ET.SubElement(box,'m_bPlaced',{'val':'1'})
        result=reader.model_summary(root)['charts'][0]['presentation_labels']
        self.assertEqual(len(result),1);self.assertEqual(result[0]['series_index'],0);self.assertEqual(result[0]['token'],'invented-token')


class CompoundTests(unittest.TestCase):
    def test_regular_stream(self):
        obj=reader.CompoundFile(cfb(MODEL))
        self.assertEqual(obj.read('think-cellXML').rstrip(),MODEL)

    def test_mini_stream(self):
        value=b'<?xml version="1.0"?><root/>'
        self.assertEqual(reader.CompoundFile(cfb(value,mini=True)).read('think-cellXML'),value)

    def test_cycle_rejected(self):
        damaged=bytearray(cfb(MODEL))
        struct.pack_into('<I',damaged,512+2*4,2)
        obj=reader.CompoundFile(bytes(damaged))
        with self.assertRaisesRegex(reader.InspectionError,'cyclic'):
            obj.read('think-cellXML')

    def test_outside_sector_rejected(self):
        damaged=bytearray(cfb(MODEL))
        struct.pack_into('<I',damaged,512+2*4,100)
        obj=reader.CompoundFile(bytes(damaged))
        with self.assertRaises(reader.InspectionError):
            obj.read('think-cellXML')

    def test_directory_cycle_rejected(self):
        damaged=bytearray(cfb(MODEL))
        struct.pack_into('<I',damaged,1024+128+72,1)
        with self.assertRaisesRegex(reader.InspectionError,'directory tree'):
            reader.CompoundFile(bytes(damaged))

    def test_truncation_rejected(self):
        with self.assertRaises(reader.InspectionError):
            reader.CompoundFile(cfb(MODEL)[:-1])

    def test_difat_count_rejected(self):
        damaged=bytearray(cfb(MODEL))
        struct.pack_into('<I',damaged,72,999999)
        with self.assertRaisesRegex(reader.InspectionError,'DIFAT'):
            reader.CompoundFile(bytes(damaged))

    def test_difat_extension(self):
        original=cfb(MODEL)
        header=bytearray(original[:512])
        struct.pack_into('<I',header,44,2)
        struct.pack_into('<II',header,68,2,1)
        fat=[reader.FREE]*128
        fat[0]=fat[3]=0xfffffffd
        fat[1]=reader.END
        fat[2]=0xfffffffc
        for i in range(4,11):
            fat[i]=i+1
        fat[11]=reader.END
        directory=bytearray(original[1024:1536])
        struct.pack_into('<I',directory,128+116,4)
        difat=struct.pack('<128I',3,*([reader.FREE]*126),reader.END)
        extra_fat=struct.pack('<128I',*([reader.FREE]*128))
        fixture=bytes(header)+struct.pack('<128I',*fat)+bytes(directory)+difat+extra_fat+original[1536:]
        self.assertEqual(reader.CompoundFile(fixture).read('think-cellXML').rstrip(),MODEL)


class ModelTests(unittest.TestCase):
    def test_reference_graph_values_and_storage(self):
        summary=reader.model_summary(reader.parse_xml(MODEL))
        chart=summary['charts'][0]
        self.assertEqual(summary['unresolved_idrefs'],[])
        self.assertEqual(chart['automation_name'],'Revenue')
        self.assertEqual(chart['child_package_stream'],'think-cellChild0/Package')
        self.assertEqual(chart['series'][0]['label']['string'],'Sales')
        self.assertEqual(chart['categories'][0]['label']['string'],'Q1')
        self.assertEqual(chart['categories'][0]['values'][0]['absolute']['number'],12.5)
        self.assertIsNone(chart['categories'][0]['values'][0]['relative'])

    def test_duplicate_ids_rejected(self):
        with self.assertRaisesRegex(reader.InspectionError,'Duplicate model'):
            reader.model_summary(reader.parse_xml('<root><A id="1"/><B id="1"/></root>'))

    def test_missing_reference_reported(self):
        summary=reader.model_summary(reader.parse_xml('<root><CSequenceChartSE id="1"><m_dtable idref="9"/></CSequenceChartSE></root>'))
        self.assertEqual(summary['unresolved_idrefs'],['9'])
        self.assertIn('warning',summary['charts'][0])

    def test_non_numeric_and_unknown_preserved(self):
        self.assertEqual(reader.var_value(reader.parse_xml('<m_varval type="5"/>'))['string'],'')
        self.assertIn('warning',reader.var_value(reader.parse_xml('<m_varval type="1" val="NaN"/>')))
        self.assertEqual(reader.var_value(reader.parse_xml('<m_varval type="42" val="x"/>'))['attributes']['val'],'x')


class PackageTests(unittest.TestCase):
    def test_full_inspection_synthetic_package(self):
        data=pptx([('ppt/embeddings/oleObject1.bin',cfb(MODEL)),
                    ('ppt/slides/slide1.xml','<p:sld xmlns:p="'+reader.P+'"/>')])
        result=reader.inspect_bytes(data)
        self.assertEqual(result['counts']['thinkcell_chart_models'],1)
        self.assertEqual(result['compound_files'][0]['streams'][0]['envelope']['xml_offset'],0)

    def test_named_field_prefix(self):
        root=reader.ET.Element('{'+reader.P+'}sld')
        field=reader.ET.SubElement(root,'{'+reader.A+'}fld')
        field.set('id','synthetic-id')
        field.set('type','thinkcell<?xml version="1.0" encoding="UTF-16"?><root><PersistentType><m_strName>Heading</m_strName></PersistentType></root>')
        result=reader.inspect_bytes(pptx([('ppt/slides/slide1.xml',reader.ET.tostring(root))]))
        self.assertEqual(result['text_targets'][0]['model']['automation_names'],['Heading'])

    def test_duplicate_zip_names_rejected(self):
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            data=pptx([('a.xml','<a/>'),('a.xml','<a/>')])
        with self.assertRaisesRegex(reader.InspectionError,'Duplicate ZIP'):
            reader.inspect_bytes(data)

    def test_zip_traversal_rejected(self):
        with self.assertRaisesRegex(reader.InspectionError,'Unsafe ZIP'):
            reader.inspect_bytes(pptx([('../escape','text')]))

    def test_dtd_rejected_utf8_and_utf16(self):
        xml='<!DOCTYPE root [<!ENTITY danger SYSTEM "file:///private">]><root>&danger;</root>'
        for encoding in ['utf-8','utf-16']:
            with self.assertRaisesRegex(reader.InspectionError,'DTD'):
                reader.parse_xml(xml.encode(encoding))

    def test_depth_limit(self):
        with self.assertRaisesRegex(reader.InspectionError,'complexity'):
            reader.parse_xml('<a>'*130+'</a>'*130)


if __name__=='__main__':
    unittest.main()
