
# Test Fixtures (not included)
Place these when CV1→Template1 is perfect:
- cv1.docx
- template1.docx
- cv1_in_template1_golden.docx
- cv1_sections.json
Then remove @pytest.mark.skip in tests/test_pipeline_cv1.py and run `pytest -q`.
