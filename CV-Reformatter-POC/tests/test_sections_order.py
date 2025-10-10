import os
from cv_reformatter.pipeline import _order_sections, _load_section_profile

def test_prof_dev_follows_key_skills_and_dedupes():
    poc_root = os.path.dirname(os.path.dirname(__file__))
    prof = _load_section_profile(poc_root)

    blocks = [
        {"title": "Personal Profile", "body": "aaa"},
        {"title": "Key Skills", "body": "bbb"},
        {"title": "Other Headings", "body": "LeanSix, EVA, MBA electives"},
        {"title": "Education", "body": "ccc"},
        {"title": "Other Headings", "body": "duplicate should be dropped"},
        {"title": "Employment History", "body": "ddd"}
    ]

    ordered = _order_sections(blocks, prof)
    titles = [b["title"] for b in ordered]
    assert titles[:4] == [
        "PERSONAL PROFILE",
        "KEY SKILLS",
        "PROFESSIONAL DEVELOPMENT",
        "EDUCATION",
    ]
    assert titles[4] == "EMPLOYMENT HISTORY"
    assert titles.count("PROFESSIONAL DEVELOPMENT") == 1
