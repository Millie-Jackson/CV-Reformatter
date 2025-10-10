import json, os
from docx import Document
from cv_reformatter.pipeline import reformat_cv_cv1_to_template1

def _make_minimal_template(path: str):
    doc = Document()
    # a blank doc acts as the template; meta & writers will build the content
    doc.save(path)

def _sample_fields():
    return {
        "summary": "Results-driven leader with international experience.",
        "skills": ["Leadership", "Financial Modelling"],
        # PD via canonical key; 'Other Headings' remap is also supported elsewhere
        "professional_development": [
            "LeanSigma Champion Training, TBM",
            "EVA Training, Stern Stewart"
        ],
        "education": [
            {"year": "2021", "institution": "Uni Somewhere", "title": "MBA", "award": "Distinction"}
        ],
        "employment_history": [
            {
                "title": "Head of Ops",
                "company": "Franchise Brands plc",
                "location": "London",
                "start": "Jan 2021",
                "end": "Present",
                "company_blurb": "Multi-brand franchisor.",
                "bullets": ["Led X", "Saved £3m"]
            }
        ],
        "additional_information": ["Full UK Driving Licence"]
    }

def test_pipeline_e2e_builds_doc_with_correct_section_order(tmp_path):
    template = tmp_path / "template.docx"
    _make_minimal_template(str(template))

    data_path = tmp_path / "fields.json"
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(_sample_fields(), f)

    out_path = tmp_path / "out.docx"

    # run pipeline with our standard profiles
    reformat_cv_cv1_to_template1(
        input_docx=str(template),
        template_docx=str(template),
        out_path=str(out_path),
        data_json=str(data_path),
        meta_profile="template1",
        no_legacy=True,
        section_profile_name="template1_sections.json",
    )

    # read back and assert ordering + content
    doc = Document(str(out_path))
    texts = [p.text for p in doc.paragraphs]
    joined = "\n".join(texts)

    # order: Personal Profile -> Key Skills -> Professional Development -> Education -> Employment -> Additional Info
    idx = {k: joined.find(k) for k in [
        "PERSONAL PROFILE", "KEY SKILLS", "PROFESSIONAL DEVELOPMENT",
        "EDUCATION", "EMPLOYMENT HISTORY", "ADDITIONAL INFORMATION"
    ]}
    # All present and in order
    assert all(v >= 0 for v in idx.values())
    assert idx["PERSONAL PROFILE"] < idx["KEY SKILLS"] < idx["PROFESSIONAL DEVELOPMENT"] < idx["EDUCATION"] < idx["EMPLOYMENT HISTORY"] < idx["ADDITIONAL INFORMATION"]

    # Spot-check content from each writer made it in
    assert "Results-driven leader with international experience." in joined
    assert "Leadership" in joined and "Financial Modelling" in joined
    assert "LeanSigma Champion Training, TBM" in joined
    assert "2021" in joined and "Uni Somewhere" in joined and "MBA" in joined
    assert "Head of Ops" in joined and "Franchise Brands plc" in joined and "Saved £3m" in joined
    assert "Full UK Driving Licence" in joined
