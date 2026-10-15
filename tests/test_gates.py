"""Synthetic evidence-gate checks; no rendering, provider or network calls."""
import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('toolkit', ROOT/'router/toolkit.py')
toolkit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(toolkit)

class GateTests(unittest.TestCase):
    def test_valid_plan(self):
        self.assertEqual(toolkit.validate_plan(ROOT/'examples/page-plan.json')['status'], 'passed')

    def test_reject_incomplete_plan(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as scratch:
            p=Path(scratch)/'bad.json'
            p.write_text(json.dumps({'target_slide_count':3,'slides':[]}),encoding='utf-8')
            self.assertEqual(toolkit.validate_plan(p)['status'],'failed')

    def test_reject_missing_reference(self):
        args=SimpleNamespace(prompt='Reference example',reference=[str(ROOT/'absent.pptx')],kind='formal',renderer=None,academic_ready=False,slides=1,project=None)
        with self.assertRaises(ValueError):toolkit.route(args)

    def test_hash_and_visual_review_gate(self):
        with tempfile.TemporaryDirectory(dir=Path(__file__).parent) as scratch:
            p=Path(scratch)/'synthetic.bin';p.write_bytes(b'Synthetic hash input, not a rendered deck.')
            qa=Path(scratch)/'qa';qa.mkdir()
            toolkit.write(qa/'qa-report.json',{'source_sha256':'stale','status':'passed','render_status':'passed','slide_count':1})
            result=toolkit.verify(SimpleNamespace(pptx=str(p),qa=str(qa)))
            self.assertEqual(result['status'],'failed')
            self.assertIn('QA is stale',result['errors'])
            self.assertIn('Current visual review missing or needs repair',result['errors'])
            digest=hashlib.sha256(p.read_bytes()).hexdigest()
            toolkit.write(qa/'qa-report.json',{'source_sha256':digest,'status':'passed','render_status':'passed','slide_count':1})
            toolkit.write(qa/'visual-review.json',{'source_sha256':digest,'status':'passed','slides_reviewed':[1]})
            self.assertEqual(toolkit.verify(SimpleNamespace(pptx=str(p),qa=str(qa)))['status'],'passed')
            # Synthetic receipt exercises gate logic; it is not evidence of real rendering.

if __name__=='__main__':unittest.main()
