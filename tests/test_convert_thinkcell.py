"""Conversion tests use invented metadata, never proprietary samples."""
import copy
import importlib.util
import pathlib
import sys
import unittest

TOOLS=pathlib.Path(__file__).resolve().parents[1]/'tools'
sys.path.insert(0,str(TOOLS))
import convert_thinkcell as convert
from inspect_thinkcell import InspectionError


def numeric(value):
    return {'type_code':'1','numeric_literal':str(value),'number':value}


def record():
    return {'model':{'unresolved_idrefs':[],'warnings':[]},'chart':{
        'object_type':'CSequenceChartSE','automation_name':'Synthetic',
        'series':[{'label':{'type_code':'5','string':'North'}},{'label':{'type_code':'5','string':'South'}}],
        'vector_connection_refs':[],
        'categories':[{'label':{'type_code':'5','string':'Q1'},'equals_scalar_ref':'0',
            'values':[{'absolute':numeric(10),'connection_ref':'0'},{'absolute':numeric(-2),'connection_ref':'0'}]},
            {'label':{'type_code':'1','number':2025,'numeric_literal':'2025'},'equals_scalar_ref':'0',
            'values':[{'absolute':numeric(12),'connection_ref':'0'},{'absolute':numeric(3.5),'connection_ref':'0'}]}]}}


class ConversionTests(unittest.TestCase):
    def test_explicit_kind_layout_and_values(self):
        result=convert.convert_record(record(),'Bar','Stacked')
        self.assertEqual(result['Kind'],'Bar')
        self.assertEqual(result['LayoutMode'],'Stacked')
        self.assertEqual(result['Series'],['North','South'])
        self.assertEqual(result['Rows'],[{'Label':'Q1','Values':[10,-2]},{'Label':'2025','Values':[12,3.5]}])

    def test_input_unchanged(self):
        source=record(); before=copy.deepcopy(source)
        convert.convert_record(source,'Line','Clustered')
        self.assertEqual(source,before)

    def test_explicit_settings_required(self):
        with self.assertRaises(InspectionError):
            convert.convert_record(record(),None,'Clustered')

    def test_scatter_rejected(self):
        source=record(); source['chart']['object_type']='CScatterChartSE'
        with self.assertRaisesRegex(InspectionError,'scatter'):
            convert.convert_record(source,'Line','Clustered')

    def test_equals_rejected(self):
        source=record(); source['chart']['categories'][0]['equals_scalar_ref']='7'
        with self.assertRaisesRegex(InspectionError,'equals/total'):
            convert.convert_record(source,'Column','Stacked')

    def test_connections_rejected(self):
        source=record(); source['chart']['categories'][0]['values'][0]['connection_ref']='77'
        with self.assertRaisesRegex(InspectionError,'connection'):
            convert.convert_record(source,'Column','Stacked')

    def test_missing_reference_rejected(self):
        source=record(); source['model']['unresolved_idrefs']=['42']
        with self.assertRaisesRegex(InspectionError,'unresolved'):
            convert.convert_record(source,'Column','Stacked')

    def test_empty_cell_not_silently_zeroed(self):
        source=record(); source['chart']['categories'][0]['values'][0]['absolute']={'type_code':'5','string':''}
        with self.assertRaisesRegex(InspectionError,'numeric source'):
            convert.convert_record(source,'Column','Stacked')

    def test_null_source_not_silently_zeroed(self):
        source=record(); source['chart']['categories'][0]['values'][0]['absolute']=None
        with self.assertRaises(InspectionError):
            convert.convert_record(source,'Column','Stacked')

    def test_non_finite_rejected(self):
        source=record(); source['chart']['categories'][0]['values'][0]['absolute']=numeric(float('inf'))
        with self.assertRaisesRegex(InspectionError,'non-finite'):
            convert.convert_record(source,'Column','Stacked')

    def test_literal_mismatch_rejected(self):
        source=record(); source['chart']['categories'][0]['values'][0]['absolute']['number']=99
        with self.assertRaisesRegex(InspectionError,'mismatch'):
            convert.convert_record(source,'Column','Stacked')

    def test_shape_limit(self):
        source=record(); source['chart']['categories']*=41
        with self.assertRaisesRegex(InspectionError,'80 categories'):
            convert.convert_record(source,'Column','Stacked')

    def test_scalar_count_rejected(self):
        source=record(); source['chart']['categories'][0]['values'].pop()
        with self.assertRaisesRegex(InspectionError,'count'):
            convert.convert_record(source,'Column','Stacked')

    def test_chart_index_list(self):
        source=record()
        report={'schema_version':1,'compound_files':[{'part':'a.bin','streams':[{'path':'think-cellXML','model':dict(source['model'],charts=[source['chart']])}]}]}
        self.assertEqual(convert.chart_records(report)[0]['index'],1)

    def test_line_rejects_stack_and_percent(self):
        for layout in ['Stacked','Percent']:
            with self.assertRaisesRegex(InspectionError,'Line supports only Clustered'):
                convert.convert_record(record(),'Line',layout)

    def test_percent_negative_rejected(self):
        with self.assertRaisesRegex(InspectionError,'nonnegative'):
            convert.convert_record(record(),'Column','Percent')

    def test_stacked_overflow_rejected_in_both_directions(self):
        for sign in [1,-1]:
            source=record()
            for cell in source['chart']['categories'][0]['values']:
                cell['absolute']=numeric(sign*1e308)
            with self.assertRaisesRegex(InspectionError,'cumulative value overflow'):
                convert.convert_record(source,'Column','Stacked')

    def test_percent_extreme_values_accepted(self):
        source=record()
        for cell in source['chart']['categories'][0]['values']:
            cell['absolute']=numeric(1e308)
        result=convert.convert_record(source,'Column','Percent')
        self.assertEqual(result['Rows'][0]['Values'],[1e308,1e308])

    def test_zero_percent_category_accepted(self):
        source=record()
        for cell in source['chart']['categories'][0]['values']:
            cell['absolute']=numeric(0)
        self.assertEqual(convert.convert_record(source,'Bar','Percent')['Rows'][0]['Values'],[0,0])


if __name__=='__main__':
    unittest.main()
