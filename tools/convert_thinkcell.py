#!/usr/bin/env python3
"""Convert observed sequence-chart data to an independent OpenSlideCharts JSON.

Input PPTX or inspect_thinkcell.py JSON stays unchanged. Explicit chart, kind and
layout selection required: private type codes are never guessed. This converts
supported data only; it does not preserve think-cell layout, style or editability.
"""
import argparse
import json
import math
import pathlib
import sys

from inspect_thinkcell import InspectionError, MAX_FILE, inspect_bytes


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
            print('Data conversion only: original unchanged. Layout/style/annotations/Excel links are not transferred; chosen kind/layout are explicit overrides.',file=sys.stderr)
        return 0
    except (InspectionError,OSError,KeyError,TypeError,AttributeError) as exc:
        print('Conversion failed: '+str(exc),file=sys.stderr)
        return 2


if __name__=='__main__':
    sys.exit(main())
