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


def appearance_record():
    source=record();chart=source['chart']
    for category,values in zip(chart['categories'],[[4,6],[8,2]]):
        for cell,value in zip(category['values'],values): cell['absolute']=numeric(value)
    theme={'colors':{},'mapping':{},'fonts':{'minorFont':'Example Sans'},'override_present':False}
    white={'size_points':10,'resolved_name':'Example Sans','resolved_color':'#FFFFFF'}
    native={'types':['barChart'],'grouping':['stacked'],'bar_direction':['col'],'has_title':False,'has_legend':False,
      'theme':theme,'series':[],
      'manual_plot_layout':{'layoutTarget':'inner','xMode':'edge','yMode':'edge','x':'.1','y':'.1','w':'.8','h':'.8'},
      'axes':[{'kind':'valAx','orientation':'minMax','minimum':0,'maximum':10,'tick_labels':'none',
               'line_hidden':True,'major_gridlines_present':False,'major_gridlines_hidden':False},
              {'kind':'catAx','line_hidden':False,'resolved_line_color':'#101010','line_width_points':.5}]}
    for index,(name,values,color) in enumerate([('South',[6,2],'#112233'),('North',[4,8],'#445566')]):
        native['series'].append({'order':str(index),'title':name,'cache_values':[{'index':str(i),'literal':str(v)} for i,v in enumerate(values)],
            'resolved_color':color,'gap_width':50,'overlap':100,'point_overrides':0,
            'labels':[{'index':str(i),'show_value':'1','position':'ctr','format':'0','font':copy.deepcopy(white)} for i in range(2)]})
    def link(x,y,w,h,text=''):
        return {'slide_part':'ppt/slides/slide1.xml','text':text,
           'text_fonts':[{'size_points':10,'latin':'+mn-lt','color':{'kind':'srgbClr','val':'101010','transforms':[]}}],
           'transform_emu':{'off':{'x':str(x*12700),'y':str(y*12700)},'ext':{'cx':str(w*12700),'cy':str(h*12700)}}}
    frame=link(100,50,200,100);frame['native_charts']=[native];chart['resolved_shape_links']=[frame]
    labels=[]
    for si,(name,y) in enumerate([('North',71),('South',111)]):
        labels.append({'role':'series','placed':'1','series_index':si,'category_index':None,'font_size':'10',
                       'resolved_shape_links':[link(80,y,30,10,name)]})
    for ci,(text,x) in enumerate([('Q1',160),('2025',240)]):
        for role,y,t in [('category',142,text),('total',48,'10')]:
            labels.append({'role':role,'placed':'1','series_index':None,'category_index':ci,'font_size':'10',
                          'resolved_shape_links':[link(x-10,y,20,10,t)]})
    chart['presentation_labels']=labels
    return source


class AppearanceTests(unittest.TestCase):
    def test_verified_union_geometry_and_order(self):
        result=convert.convert_record(appearance_record(),'Column','Stacked')
        self.assertNotIn('ImportWarnings',result)
        a=result['Appearance']
        self.assertEqual(a['DisplayOrder'],[1,0]);self.assertEqual(a['SeriesColors'],['#445566','#112233'])
        self.assertEqual(a['SourceFrame'],{'X':80,'Y':48,'Width':220,'Height':104})
        self.assertAlmostEqual(a['PlotRect']['X'],40/220)
        self.assertEqual(a['AxisLineWidth'],.5);self.assertEqual(a['GapRatio'],.5)
        self.assertEqual(a['TextColor'],'#101010');self.assertEqual(a['Maximum'],10)

    def test_xml_series_order_is_not_display_order(self):
        source=appearance_record();source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'].reverse()
        self.assertEqual(convert.convert_record(source,'Column','Stacked')['Appearance']['DisplayOrder'],[1,0])

    def test_uncertain_appearance_is_explicit_not_silent(self):
        source=appearance_record();source['chart']['presentation_labels'][0]['resolved_shape_links'][0]['transform_emu']['off']['y']='0'
        result=convert.convert_record(source,'Column','Stacked')
        self.assertNotIn('Appearance',result);self.assertIn('segment center',result['ImportWarnings'][0])
        self.assertEqual(result['Rows'][0]['Values'],[4,6])

    def test_unknown_color_rejected(self):
        source=appearance_record();source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['resolved_color']=None
        self.assertIn('color',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])

    def test_point_override_not_silently_lost(self):
        source=appearance_record();source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['point_overrides']=1
        self.assertIn('per-point',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])

    def test_changed_layout_drops_only_appearance_with_warning(self):
        result=convert.convert_record(appearance_record(),'Bar','Stacked')
        self.assertNotIn('Appearance',result);self.assertIn('stacked-column',result['ImportWarnings'][0])

    def test_input_appearance_unchanged(self):
        source=appearance_record();before=copy.deepcopy(source)
        convert.convert_record(source,'Column','Stacked');self.assertEqual(source,before)

    def test_bold_or_italic_is_not_silently_lost(self):
        for flag in ['bold','italic']:
            source=appearance_record()
            source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['labels'][0]['font'][flag]='1'
            self.assertIn('bold/italic',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])
            source=appearance_record();source['chart']['presentation_labels'][0]['resolved_shape_links'][0]['text_fonts'][0][flag]='true'
            self.assertIn('bold/italic',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])

    def test_explicit_factor_offsets_preserved_in_points(self):
        source=appearance_record()
        source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['labels'][0]['manual_layout']={'x':'.3','y':'-.4','xMode':'factor','yMode':'factor'}
        result=convert.convert_record(source,'Column','Stacked')
        self.assertEqual(result['Appearance']['ValueLabelOffsets'],[{'SeriesIndex':1,'CategoryIndex':0,'X':60,'Y':-40}])

    def test_absent_manual_mode_uses_observed_factor_with_warning(self):
        source=appearance_record()
        source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['labels'][0]['manual_layout']={'x':'.3','y':'.4'}
        result=convert.convert_record(source,'Column','Stacked')
        self.assertEqual(result['Appearance']['ValueLabelOffsets'],[{'SeriesIndex':1,'CategoryIndex':0,'X':60,'Y':40}])
        self.assertIn('observed',result['ImportWarnings'][0])

    def test_unsupported_explicit_manual_mode_not_assumed_factor(self):
        source=appearance_record()
        source['chart']['resolved_shape_links'][0]['native_charts'][0]['series'][0]['labels'][0]['manual_layout']={'x':'.3','y':'.4','xMode':'edge'}
        result=convert.convert_record(source,'Column','Stacked')
        self.assertNotIn('Appearance',result);self.assertIn('manual layout mode',result['ImportWarnings'][0])

    def test_duplicate_role_indices_rejected_even_with_equal_content(self):
        for role in ['series','category','total']:
            source=appearance_record();labels=source['chart']['presentation_labels'];indices=[i for i,l in enumerate(labels) if l['role']==role]
            labels[indices[1]]=copy.deepcopy(labels[indices[0]])
            result=convert.convert_record(source,'Column','Stacked')
            self.assertNotIn('Appearance',result);self.assertIn('duplicated',result['ImportWarnings'][0])

    def test_transformed_shapes_and_groups_rejected(self):
        for attrs in [{'rot':'60000'},{'flipH':'1'},{'flipV':'true'}]:
            source=appearance_record();source['chart']['resolved_shape_links'][0]['transform_attributes']=attrs
            self.assertIn('transform',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])
        source=appearance_record();source['chart']['presentation_labels'][0]['resolved_shape_links'][0]['parent_group_transforms']=[
            {'attributes':{},'coordinates':{'off':{'x':'10','y':'0'},'chOff':{'x':'0','y':'0'},'ext':{'cx':'100','cy':'100'},'chExt':{'cx':'100','cy':'100'}}}]
        self.assertIn('composition',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])

    def test_neutral_group_accepted(self):
        source=appearance_record();source['chart']['resolved_shape_links'][0]['parent_group_transforms']=[
            {'attributes':{'rot':'0','flipH':'false'},'coordinates':{'off':{'x':'10','y':'20'},'chOff':{'x':'10','y':'20'},'ext':{'cx':'100','cy':'200'},'chExt':{'cx':'100','cy':'200'}}}]
        self.assertIn('Appearance',convert.convert_record(source,'Column','Stacked'))

    def test_label_metrics_measured_not_fixed_defaults(self):
        result=convert.convert_record(appearance_record(),'Column','Stacked')['Appearance']
        self.assertAlmostEqual(result['SeriesLabelGap'],23.3333333333)
        self.assertEqual(result['CategoryLabelGap'],2);self.assertEqual(result['TotalLabelGap'],2)
        for role in ['Series','Category','Total']: self.assertEqual(result[role+'LabelHeight'],10)

    def test_nonuniform_label_heights_warn(self):
        source=appearance_record();target=source['chart']['presentation_labels'][0]['resolved_shape_links'][0]
        target['transform_emu']['ext']['cy']=str(10.5*12700)
        self.assertIn('nonuniform',convert.convert_record(source,'Column','Stacked')['ImportWarnings'][0])


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
