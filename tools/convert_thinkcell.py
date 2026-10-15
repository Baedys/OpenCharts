#!/usr/bin/env python3
"""Convert observed sequence-chart data to an independent OpenSlideCharts JSON.

Input PPTX or inspect_thinkcell.py JSON stays unchanged. Explicit chart, kind and
layout selection required: private type codes are never guessed. This converts
supported data and a verified subset of native stacked-column appearance. It
does not preserve private editability or guarantee visual round-trip fidelity.
"""
import argparse
import json
import math
import pathlib
import sys

from inspect_thinkcell import InspectionError, MAX_FILE, inspect_bytes, theme_color


def frame_points(link):
    def neutral_flags(attributes):
        if attributes is None: return False
        try: rotation=int(attributes.get('rot','0'))
        except (ValueError,TypeError): return False
        return rotation%21600000==0 and all(attributes.get(k) in (None,'0','false') for k in ['flipH','flipV'])
    if not neutral_flags(link.get('transform_attributes',{})):
        raise InspectionError('Appearance: rotated/flipped shape transform unsupported')
    for group in link.get('parent_group_transforms',[]) or []:
        if not neutral_flags(group.get('attributes')):
            raise InspectionError('Appearance: nonneutral parent group transform unsupported')
        coords=group.get('coordinates') or {}
        try:
            same=all(int(coords[a][k])==int(coords[b][k]) for a,b,keys in
                [('off','chOff',['x','y']),('ext','chExt',['cx','cy'])] for k in keys)
            positive=all(int(coords['ext'][k])>0 for k in ['cx','cy'])
        except (KeyError,TypeError,ValueError): same=positive=False
        if not same or not positive:
            raise InspectionError('Appearance: parent group translation/scaling requires composition')
    try:
        t=link['transform_emu']
        frame={k:float(t[group][name])/12700 for k,group,name in
               [('X','off','x'),('Y','off','y'),('Width','ext','cx'),('Height','ext','cy')]}
    except (KeyError,TypeError,ValueError,OverflowError) as exc:
        raise InspectionError('Appearance: incomplete shape bounds') from exc
    if not all(math.isfinite(v) for v in frame.values()) or frame['Width']<=0 or frame['Height']<=0:
        raise InspectionError('Appearance: invalid shape bounds')
    return frame


def observed_appearance(chart, names, rows, kind, layout, warnings=None):
    """Strict initial appearance subset, checked against data and label geometry.

    Every inferred label placement must agree with its linked source shape.
    Unknown style is reported to the caller; no private enum guessed.
    """
    warnings=warnings if warnings is not None else []
    links=chart.get('resolved_shape_links',[])
    if not links: return None
    if len(links)!=1 or len(links[0].get('native_charts',[]))!=1:
        raise InspectionError('Appearance: ambiguous native chart relationship')
    native=links[0]['native_charts'][0]
    if not native or 'series' not in native: return None
    if kind!='Column' or layout!='Stacked' or native.get('types')!=['barChart'] or native.get('grouping')!=['stacked'] or native.get('bar_direction')!=['col']:
        raise InspectionError('Appearance: only observed stacked-column layout supported')
    if any(v<0 for row in rows for v in row['Values']):
        raise InspectionError('Appearance: signed stacked-label geometry not verified')
    if native['has_title'] or native['has_legend']:
        raise InspectionError('Appearance: native title/legend geometry not supported')
    ns=native['series']
    if len(ns)!=len(names): raise InspectionError('Appearance: native/model series count mismatch')
    if {s.get('order') for s in ns}!={str(i) for i in range(len(ns))}:
        raise InspectionError('Appearance: native series order invalid')
    ns=sorted(ns,key=lambda s:int(s['order']))
    frame=frame_points(links[0]);manual_offsets=[];value_offsets=[]
    order=[]; colors=[None]*len(names); formats=set(); fonts=set(); sizes=set()
    for series in ns:
        if series.get('point_overrides'):
            raise InspectionError('Appearance: per-point styles unsupported')
        cache=series['cache_values']
        if [x['index'] for x in cache]!=[str(i) for i in range(len(rows))]:
            raise InspectionError('Appearance: incomplete native cache indices')
        try: values=[float(x['literal']) for x in cache]
        except (ValueError,TypeError,OverflowError) as exc:
            raise InspectionError('Appearance: native numeric cache malformed') from exc
        matches=[i for i,name in enumerate(names) if (series['title'] is None or series['title']==name) and
                 all(math.isfinite(v) and math.isclose(v,row['Values'][i],rel_tol=1e-12,abs_tol=1e-12) for v,row in zip(values,rows))]
        if len(matches)!=1 or matches[0] in order:
            raise InspectionError('Appearance: native/model series identity ambiguous')
        si=matches[0];order.append(si);colors[si]=series.get('resolved_color')
        if not colors[si]: raise InspectionError('Appearance: series color cannot be resolved exactly')
        labels=series['labels']
        if len(labels)!=len(rows) or {l['index'] for l in labels}!={str(i) for i in range(len(rows))}:
            raise InspectionError('Appearance: partial value labels unsupported')
        for label in labels:
            font=label.get('font') or {}
            if any(font.get(k) not in (None,'0','false') for k in ['bold','italic']):
                raise InspectionError('Appearance: bold/italic value font unsupported')
            if label['show_value']!='1' or label['position']!='ctr' or font.get('resolved_color')!='#FFFFFF':
                raise InspectionError('Appearance: only centered white value labels supported')
            formats.add(label['format']);fonts.add(font.get('resolved_name'));sizes.add(font.get('size_points'))
            manual=label.get('manual_layout') or {}
            if manual:
                if set(manual)-{'x','y','xMode','yMode'} or any(manual.get(k) not in (None,'factor') for k in ['xMode','yMode']):
                    raise InspectionError('Appearance: unsupported value-label manual layout mode/size')
                try: x=float(manual.get('x',0));y=float(manual.get('y',0))
                except (ValueError,TypeError,OverflowError) as exc: raise InspectionError('Appearance: invalid value-label manual offset') from exc
                if not math.isfinite(x) or not math.isfinite(y): raise InspectionError('Appearance: nonfinite value-label manual offset')
                if x or y:
                    # MS-OE376 ST_LayoutMode: explicit factor is displacement
                    # from default position, relative to chart width/height.
                    # Missing modes matched explicit factor pixel-for-pixel in
                    # Office 16.0.20430.20146 on the checked stacked-column
                    # corpus. Keep this empirical assumption visible.
                    if (x and manual.get('xMode') is None) or (y and manual.get('yMode') is None):
                        manual_offsets.append((x,y))
                    value_offsets.append({'SeriesIndex':si,'CategoryIndex':int(label['index']),
                                          'X':x*frame['Width'],'Y':y*frame['Height']})
    if len(formats)!=1 or next(iter(formats)) not in ('#,##0;"-"#,##0','#,##0','0') or len(fonts)!=1 or None in fonts or len(sizes)!=1 or None in sizes:
        raise InspectionError('Appearance: nonuniform or unsupported number/font formatting')
    if any(not isinstance(s,(int,float)) or not math.isfinite(s) or not 2<=s<=96 for s in sizes):
        raise InspectionError('Appearance: value font size outside renderer limits')
    axes=[a for a in native['axes'] if a['kind']=='valAx']
    if len(axes)!=1: raise InspectionError('Appearance: multiple/missing value axes')
    axis=axes[0]
    if axis['orientation']!='minMax' or axis['minimum']!=0 or axis['maximum'] is None or axis['maximum']<=0:
        raise InspectionError('Appearance: unsupported native scale')
    if axis['tick_labels']!='none' or not axis['line_hidden'] or (axis['major_gridlines_present'] and not axis['major_gridlines_hidden']):
        raise InspectionError('Appearance: visible value-axis geometry unsupported')
    manual=native.get('manual_plot_layout') or {}
    if manual.get('layoutTarget')!='inner' or manual.get('xMode')!='edge' or manual.get('yMode')!='edge':
        raise InspectionError('Appearance: unsupported manual plot coordinate mode')
    try: fractions={k:float(manual[source]) for k,source in [('X','x'),('Y','y'),('Width','w'),('Height','h')]}
    except (ValueError,TypeError,KeyError,OverflowError) as exc: raise InspectionError('Appearance: missing manual plot bounds') from exc
    if not all(math.isfinite(v) for v in fractions.values()) or fractions['Width']<=0 or fractions['Height']<=0 or fractions['X']<0 or fractions['Y']<0 or fractions['X']+fractions['Width']>1.000001 or fractions['Y']+fractions['Height']>1.000001:
        raise InspectionError('Appearance: invalid manual plot bounds')
    plot={'X':frame['X']+fractions['X']*frame['Width'],'Y':frame['Y']+fractions['Y']*frame['Height'],
          'Width':fractions['Width']*frame['Width'],'Height':fractions['Height']*frame['Height']}
    gaps={s['gap_width'] for s in ns};overlaps={s['overlap'] for s in ns}
    if len(gaps)!=1 or None in gaps or not 0<=next(iter(gaps))<=500 or len(overlaps)!=1 or overlaps!={100.0}:
        raise InspectionError('Appearance: inconsistent gaps/overlap')
    step=plot['Width']/len(rows);bar_width=step/(1+next(iter(gaps))/100)
    linked=[]; text_colors=set();role_fonts={role:set() for role in ['series','category','total']}; role_indices={role:set() for role in role_fonts}
    role_gaps={role:[] for role in role_fonts};role_heights={role:[] for role in role_fonts}
    baseline=plot['Y']+plot['Height'];scale=plot['Height']/axis['maximum']
    for label in chart.get('presentation_labels',[]):
        if label['placed']!='1': continue
        role=label['role'];targets=label.get('resolved_shape_links',[])
        if role not in role_indices or len(targets)!=1 or targets[0]['slide_part']!=links[0]['slide_part']:
            raise InspectionError('Appearance: unresolved or ambiguous linked text label')
        bounds=frame_points(targets[0]);text=targets[0]['text']
        try: size=float(label['font_size'])
        except (TypeError,ValueError,OverflowError) as exc: raise InspectionError('Appearance: invalid linked label font') from exc
        if not math.isfinite(size) or not 2<=size<=96: raise InspectionError('Appearance: invalid linked label font')
        for f in targets[0].get('text_fonts',[]):
            if any(f.get(k) not in (None,'0','false') for k in ['bold','italic']):
                raise InspectionError('Appearance: bold/italic linked text font unsupported')
            if f.get('size_points') is not None and f['size_points']!=size:
                raise InspectionError('Appearance: model/native label font mismatch')
            color=theme_color(native.get('theme',{}),f.get('color'))
            if color is None: raise InspectionError('Appearance: linked text color unresolved')
            text_colors.add(color)
            font_name=f.get('latin')
            font_name={'+mn-lt':native.get('theme',{}).get('fonts',{}).get('minorFont'),
                       '+mj-lt':native.get('theme',{}).get('fonts',{}).get('majorFont')}.get(font_name,font_name)
            if font_name not in fonts: raise InspectionError('Appearance: linked text font differs from value labels')
        role_fonts[role].add(size)
        ci=label.get('category_index');si=label.get('series_index')
        role_index=si if role=='series' else ci
        if not isinstance(role_index,int) or isinstance(role_index,bool) or role_index in role_indices[role]:
            raise InspectionError('Appearance: duplicated/invalid '+role+' label index')
        role_indices[role].add(role_index);role_heights[role].append(bounds['Height'])
        center_x=bounds['X']+bounds['Width']/2;center_y=bounds['Y']+bounds['Height']/2
        if role=='series':
            if not isinstance(si,int) or not 0<=si<len(names) or text!=names[si]: raise InspectionError('Appearance: series label mismatch')
            before=sum(rows[0]['Values'][i] for i in order[:order.index(si)])
            expected=baseline-(before+rows[0]['Values'][si]/2)*scale
            if abs(center_y-expected)>1 or bounds['X']+bounds['Width']>=plot['X']+plot['Width']/len(rows)/2:
                raise InspectionError('Appearance: series label is not at first-stack segment center')
            role_gaps[role].append(plot['X']+(step-bar_width)/2-bounds['X']-bounds['Width'])
        else:
            if not isinstance(ci,int) or not 0<=ci<len(rows): raise InspectionError('Appearance: category label index invalid')
            expected_x=plot['X']+(ci+.5)*plot['Width']/len(rows)
            if abs(center_x-expected_x)>1: raise InspectionError('Appearance: linked label not centered on category')
            if role=='category':
                if text!=rows[ci]['Label'] or not 0<=bounds['Y']-baseline<=size:
                    raise InspectionError('Appearance: unsupported category label content/position')
                role_gaps[role].append(bounds['Y']-baseline)
            else:
                total=sum(rows[ci]['Values']);top=baseline-total*scale
                if text!=format(total,'.0f') or not 0<=top-(bounds['Y']+bounds['Height'])<=size:
                    raise InspectionError('Appearance: total label content/position not verified')
                role_gaps[role].append(top-bounds['Y']-bounds['Height'])
        linked.append(bounds)
    if role_indices!={'series':set(range(len(names))),'category':set(range(len(rows))),'total':set(range(len(rows)))} or any(len(v)!=1 for v in role_fonts.values()):
        raise InspectionError('Appearance: incomplete or nonuniform associated labels')
    metrics={}
    for role,prefix in [('series','Series'),('category','Category'),('total','Total')]:
        for suffix,values in [('Gap',role_gaps[role]),('Height',role_heights[role])]:
            if min(values)<0 or max(values)-min(values)>.1:
                raise InspectionError('Appearance: nonuniform or negative '+role+' label '+suffix.lower())
            metrics[prefix+'Label'+suffix]=sum(values)/len(values)
    if len(text_colors)!=1: raise InspectionError('Appearance: linked text colors are not uniform')
    category_axes=[a for a in native['axes'] if a['kind']=='catAx']
    if len(category_axes)!=1 or category_axes[0]['line_hidden'] or not category_axes[0].get('resolved_line_color') or category_axes[0].get('line_width_points') is None or not 0<category_axes[0]['line_width_points']<=10:
        raise InspectionError('Appearance: category baseline style unresolved')
    all_bounds=[frame]+linked
    left=min(b['X'] for b in all_bounds);top=min(b['Y'] for b in all_bounds)
    right=max(b['X']+b['Width'] for b in all_bounds);bottom=max(b['Y']+b['Height'] for b in all_bounds)
    union={'X':left,'Y':top,'Width':right-left,'Height':bottom-top}
    if manual_offsets:
        warnings.append('Appearance subset: absent data-label layout modes interpreted as factor, as observed for the checked stacked-column profile in Office 16.0.20430.20146; behavior outside that profile is not established')
    return {**metrics,'ValueLabelOffsets':value_offsets,'SeriesColors':colors,'DisplayOrder':order,'ShowTitle':False,'ShowLegend':False,
        'ShowValueAxis':False,'ShowGridlines':False,'ValueFormat':next(iter(formats)),
        'ValueDigits':0,'FontName':next(iter(fonts)),'ValueFontSize':next(iter(sizes)),
        'CategoryFontSize':next(iter(role_fonts['category'])),'SeriesFontSize':next(iter(role_fonts['series'])),
        'TotalFontSize':next(iter(role_fonts['total'])),'TotalFormat':next(iter(formats)),
        'ShowTotals':True,'SeriesLabelsAtFirstStack':True,'GapRatio':next(iter(gaps))/100,
        'TextColor':next(iter(text_colors)),'ValueTextColor':'#FFFFFF',
        'AxisColor':category_axes[0]['resolved_line_color'],'AxisLineWidth':category_axes[0]['line_width_points'],
        'OverlapPercent':100,'Minimum':axis['minimum'],'Maximum':axis['maximum'],
        'SourceFrame':union,'PlotRect':{'X':(plot['X']-left)/union['Width'],'Y':(plot['Y']-top)/union['Height'],
          'Width':plot['Width']/union['Width'],'Height':plot['Height']/union['Height']}}


def chart_records(report):
    if not isinstance(report, dict) or report.get('schema_version') != 1:
        raise InspectionError('Expected inspection JSON schema_version 1')
    records = []
    for container in report.get('compound_files', []):
        for stream in container.get('streams', []):
            model = stream.get('model', {})
            for chart in model.get('charts', []):
                records.append({'index':len(records)+1,'part':container.get('part'),
                    'stream':stream.get('path'),'model':model,'chart':chart})
    return records


def finite_number(value, location):
    if not isinstance(value, dict) or value.get('type_code') != '1':
        raise InspectionError(location + ': expected an observed numeric source (type 1); no coercion performed')
    if value.get('warning'):
        raise InspectionError(location + ': ' + value['warning'])
    literal = value.get('numeric_literal')
    number = value.get('number')
    if isinstance(number, bool) or not isinstance(number, (int,float)) or not math.isfinite(number):
        raise InspectionError(location + ': numeric value is missing or non-finite')
    try:
        parsed = float(literal)
    except (ValueError,TypeError,OverflowError) as exc:
        raise InspectionError(location + ': invalid numeric literal') from exc
    if not math.isfinite(parsed) or parsed != number:
        raise InspectionError(location + ': numeric literal/value mismatch')
    return number


def label_value(value, location):
    if not isinstance(value, dict):
        raise InspectionError(location + ': missing label source')
    typ = value.get('type_code')
    if typ == '5' and isinstance(value.get('string'),str):
        return value['string']
    if typ == '1':
        number = finite_number(value,location)
        return format(number,'.15g') if float(format(number,'.15g')) == number else repr(number)
    if typ == '6' and isinstance(value.get('date_literal'),str) and value['date_literal']:
        return value['date_literal']
    raise InspectionError(location + ': unsupported label type; no value discarded')


def convert_record(record, kind, layout, title=None):
    if kind not in ('Column','Bar','Line'):
        raise InspectionError('Kind must be Column, Bar or Line')
    if layout not in ('Clustered','Stacked','Percent'):
        raise InspectionError('Layout must be Clustered, Stacked or Percent')
    if kind == 'Line' and layout != 'Clustered':
        raise InspectionError('Line supports only Clustered layout; stacked/percent lines are not implemented')
    model, chart = record['model'], record['chart']
    if model.get('unresolved_idrefs'):
        raise InspectionError('Model contains unresolved object references: ' + ', '.join(model['unresolved_idrefs']))
    if model.get('warnings'):
        raise InspectionError('Model structural warnings: ' + '; '.join(model['warnings']))
    if chart.get('warning'):
        raise InspectionError('Chart is incomplete: ' + chart['warning'])
    if chart.get('object_type') != 'CSequenceChartSE':
        raise InspectionError('Only CSequenceChartSE data supported; scatter/bubble and other model classes are not converted')
    if chart.get('vector_connection_refs'):
        raise InspectionError('Chart has vector connections; unsupported geometry would be lost')
    series = chart.get('series')
    categories = chart.get('categories')
    if not isinstance(series,list) or not 1 <= len(series) <= 8:
        raise InspectionError('Expected between 1 and 8 series')
    if not isinstance(categories,list) or not 1 <= len(categories) <= 80:
        raise InspectionError('Expected between 1 and 80 categories')
    names=[]
    for index,entry in enumerate(series,1):
        if not isinstance(entry,dict):
            raise InspectionError('Missing series reference: '+str(index))
        names.append(label_value(entry.get('label'),'Series '+str(index)))
    rows=[]
    for index,category in enumerate(categories,1):
        location='Category '+str(index)
        if not isinstance(category,dict) or category.get('warning'):
            raise InspectionError(location+': missing or invalid category reference')
        if category.get('equals_scalar_ref') not in (None,'0'):
            raise InspectionError(location+': equals/total reference unsupported; no waterfall semantics inferred')
        values=category.get('values')
        if not isinstance(values,list) or len(values)!=len(names):
            raise InspectionError(location+': scalar count differs from series count')
        numeric=[]
        for col,value in enumerate(values,1):
            cell=location+', series '+str(col)
            if not isinstance(value,dict) or value.get('warning'):
                raise InspectionError(cell+': missing or invalid scalar reference')
            if value.get('connection_ref') not in (None,'0'):
                raise InspectionError(cell+': scalar connection unsupported')
            numeric.append(finite_number(value.get('absolute'),cell))
        if layout == 'Percent' and any(value < 0 for value in numeric):
            raise InspectionError(location+': Percent layout requires nonnegative values')
        if layout == 'Stacked':
            positive, negative = 0.0, 0.0
            for value in numeric:
                if value < 0:
                    negative += value
                else:
                    positive += value
                if not math.isfinite(positive) or not math.isfinite(negative):
                    raise InspectionError(location+': stacked positive/negative cumulative value overflow')
        rows.append({'Label':label_value(category.get('label'),location+' label'),'Values':numeric})
    result={'Kind':kind,'Title':title if title is not None else (chart.get('automation_name') or 'Graphique importé'),
        'LayoutMode':layout,'Series':names,'Rows':rows}
    if not isinstance(result['Title'],str):
        raise InspectionError('Title must be text')
    try:
        appearance_warnings=[]
        appearance=observed_appearance(chart,names,rows,kind,layout,appearance_warnings)
        if appearance: result['Appearance']=appearance
        if appearance_warnings: result['ImportWarnings']=appearance_warnings
    except InspectionError as exc:
        result['ImportWarnings']=[str(exc)+'; data converted, source appearance not applied']
    return result


def load_report(path):
    if path.stat().st_size>MAX_FILE:
        raise InspectionError('Input exceeds size limit')
    data=path.read_bytes()
    if data.startswith(b'PK'):
        return inspect_bytes(data,path.name)
    try:
        return json.loads(data.decode('utf-8-sig'),parse_constant=lambda value: (_ for _ in ()).throw(InspectionError('Non-finite JSON number: '+value)))
    except (ValueError,UnicodeError) as exc:
        raise InspectionError('Expected PPTX or inspection JSON: '+str(exc)) from exc


def main(argv=None):
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input',type=pathlib.Path)
    parser.add_argument('--list',action='store_true',help='List one-based chart indices without converting')
    parser.add_argument('--chart',type=int,help='One-based chart index from --list')
    parser.add_argument('--kind',choices=['Column','Bar','Line'])
    parser.add_argument('--layout',choices=['Clustered','Stacked','Percent'])
    parser.add_argument('--title')
    parser.add_argument('--json',dest='output',type=pathlib.Path,help='Destination OpenSlideCharts JSON (default: stdout)')
    args=parser.parse_args(argv)
    if not args.list and (args.chart is None or args.kind is None or args.layout is None):
        parser.error('Conversion requires --chart, --kind and --layout; use --list first')
    try:
        if args.output and args.output.resolve()==args.input.resolve():
            raise InspectionError('Output must differ from input')
        records=chart_records(load_report(args.input))
        if args.list:
            result=[{'index':r['index'],'part':r['part'],'stream':r['stream'],
                'name':r['chart'].get('automation_name'),'object_type':r['chart'].get('object_type'),
                'type_code_observed':r['chart'].get('chart_type_code'),
                'series_count':len(r['chart'].get('series',[])),
                'category_count':len(r['chart'].get('categories',[])),
                'native_shape_links':r['chart'].get('resolved_shape_links',[])} for r in records]
        else:
            if not 1<=args.chart<=len(records):
                raise InspectionError('Chart index outside 1..'+str(len(records)))
            result=convert_record(records[args.chart-1],args.kind,args.layout,args.title)
        serialized=json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False)
        if args.output:
            args.output.write_text(serialized,encoding='utf-8')
            print('Written: '+str(args.output))
        else:
            print(serialized)
        if not args.list:
            print('Original unchanged. Verified appearance subset transferred when available; private editability, Excel links and unsupported annotations are not transferred.',file=sys.stderr)
            for warning in result.get('ImportWarnings',[]): print('Warning: '+warning,file=sys.stderr)
        return 0
    except (InspectionError,OSError,KeyError,TypeError,AttributeError) as exc:
        print('Conversion failed: '+str(exc),file=sys.stderr)
        return 2


if __name__=='__main__':
    sys.exit(main())
