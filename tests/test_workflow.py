import copy
import json
from pathlib import Path
import tempfile
import unittest
import zipfile

from openpyxl import Workbook,load_workbook
from config_builder.workflow import KEYS,HEADERS,prepare_config,apply_config,save,validate_rows
from main import main


class ConfigWorkflowTests(unittest.TestCase):
    def setUp(self):
        scratch=Path(__file__).resolve().parents[1]/'.test_tmp';scratch.mkdir(exist_ok=True)
        self.temp=tempfile.TemporaryDirectory(dir=scratch);self.addCleanup(self.temp.cleanup)
        self.root=Path(self.temp.name);self.source=self.root/'form.xlsx'
        wb=Workbook();s=wb.active;s.title='Form';s['A1']='Project';s.merge_cells('B1:D1');s['B1']='Existing project'
        s['A3']='Welder';s['F3']='WPS';s['A4']='W-001';s['F4']='WPS-01';s['A8']='=1+2'
        wb.create_sheet('Condition')['Q2']='process';wb.save(self.source);wb.close()
        self.ready=prepare_config(self.source,self.root/'jobs',['Form','Condition'])
        self.request_path=Path(self.ready['request']);self.request=json.loads(self.request_path.read_text())
        row={k:'' for k in KEYS};row.update(label='Project',cells='A1',group_name='Header',sequence=1,db='B1',data_type='Text',sheet='Form')
        self.plan=dict(source_hash=self.request['source_hash'],rows=[row],review_notes=['Required/default/expression unconfirmed.'])

    def apply(self):
        save(self.request_path.with_name('config-plan.json'),self.plan)
        return apply_config(self.request_path,self.root/'output')

    def test_end_to_end_preserves_original_and_other_package_assets(self):
        before=self.source.read_bytes();result=self.apply();self.assertEqual(self.source.read_bytes(),before)
        with zipfile.ZipFile(self.source) as a,zipfile.ZipFile(result['output']) as b:
            allowed={'xl/workbook.xml','xl/_rels/workbook.xml.rels','[Content_Types].xml'}
            for name in a.namelist():
                if name not in allowed:self.assertEqual(a.read(name),b.read(name),name)
        wb=load_workbook(result['output']);self.addCleanup(wb.close)
        self.assertEqual([wb['Config'].cell(1,c).value for c in range(1,15)],HEADERS)
        self.assertEqual(wb['Config']['D2'].value,1)
        self.assertEqual(wb['Form']['A8'].value,'=1+2')
        self.assertEqual(wb['Form']['B1'].value,'Existing project')
        self.assertEqual(wb['Config'].freeze_panes,'A2')
        self.assertEqual(result['api_calls'],0)

    def test_existing_config_replaced_only_in_output(self):
        wb=load_workbook(self.source);wb.create_sheet('Config')['A1']='Outdated mapping';wb.save(self.source);wb.close()
        self.ready=prepare_config(self.source,self.root/'jobs',['Form','Condition']);self.request_path=Path(self.ready['request'])
        self.plan['source_hash']=json.loads(self.request_path.read_text())['source_hash']
        result=self.apply()
        with zipfile.ZipFile(self.source) as a,zipfile.ZipFile(result['output']) as b:
            changed=[n for n in a.namelist() if a.read(n)!=b.read(n)]
        self.assertEqual(changed,['xl/worksheets/sheet3.xml'])

    def test_changed_workbook_and_existing_output_rejected(self):
        self.apply()
        with self.assertRaisesRegex(ValueError,'existing files'):
            self.apply()
        wb=load_workbook(self.source);wb.active['B1']='Changed';wb.save(self.source);wb.close()
        with self.assertRaisesRegex(ValueError,'changed'):
            self.apply()

    def test_merged_child_duplicate_target_and_label_overlap_rejected(self):
        self.plan['rows'][0]['db']='C1'
        with self.assertRaisesRegex(ValueError,'merged anchor'):validate_rows(self.plan,self.request,self.source)
        self.plan['rows'][0]['db']='B1';duplicate=copy.deepcopy(self.plan['rows'][0]);duplicate['sequence']=2
        self.plan['rows'].append(duplicate)
        with self.assertRaisesRegex(ValueError,'Duplicate DB'):validate_rows(self.plan,self.request,self.source)
        self.plan['rows'].pop();self.plan['rows'][0]['db']='A1'
        with self.assertRaisesRegex(ValueError,'overlaps'):validate_rows(self.plan,self.request,self.source)

    def test_tabular_schema_and_dropdown_evidence(self):
        row=self.plan['rows'][0];row.update(label='Welder, WPS',cells='',db='A4',group_name='DataTable',data_type='tabular',column_list='A, F')
        validate_rows(self.plan,self.request,self.source)
        row['column_list']='A, A'
        with self.assertRaisesRegex(ValueError,'ColumnList'):validate_rows(self.plan,self.request,self.source)
        row.update(label='Process',db='Q2',sheet='Condition',data_type='Dropdown',column_list='')
        with self.assertRaisesRegex(ValueError,'explicit options'):validate_rows(self.plan,self.request,self.source)
        row['option_value']='PROCESS, NON PROCESS';validate_rows(self.plan,self.request,self.source)

    def test_rejects_executable_metadata_and_unknown_sheet(self):
        self.plan['rows'][0]['default_value']='=HYPERLINK("https://example.invalid")'
        with self.assertRaisesRegex(ValueError,'executable'):validate_rows(self.plan,self.request,self.source)
        self.plan['rows'][0]['default_value']='';self.plan['rows'][0]['sheet']='Missing'
        with self.assertRaisesRegex(ValueError,'selected'):validate_rows(self.plan,self.request,self.source)

    def test_condition_target_can_reference_form_label(self):
        row=self.plan['rows'][0]
        row.update(label='Process',cells='A1',sheet='Condition',db='Q2',data_type='Dropdown',option_value='PROCESS, NON PROCESS')
        validate_rows(self.plan,self.request,self.source)
        row['cells']='Z99'
        with self.assertRaisesRegex(ValueError,'label sheet'):validate_rows(self.plan,self.request,self.source)

    def test_refuses_config_formula_dependencies(self):
        wb=load_workbook(self.source);wb.create_sheet('Config')['A1']='Value';wb['Form']['F8']='=Config!A1';wb.save(self.source);wb.close()
        with self.assertRaisesRegex(ValueError,'dependencies'):prepare_config(self.source,self.root/'jobs',['Form'])

    def test_prepare_keeps_plan_and_cli_needs_no_formatter_config(self):
        self.apply();before=self.request_path.with_name('config-plan.json').read_bytes()
        prepare_config(self.source,self.root/'jobs',['Form','Condition'])
        self.assertEqual(before,self.request_path.with_name('config-plan.json').read_bytes())
        self.assertEqual(main(['prepare',str(self.source),'--sheets','Form','--jobs',str(self.root/'cli-jobs')]),0)


if __name__=='__main__':unittest.main()
