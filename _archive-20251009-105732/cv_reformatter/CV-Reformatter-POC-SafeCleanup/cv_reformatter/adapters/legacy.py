
def run_legacy(input_docx: str, template_docx: str):
    try:
        from . import legacy_impl  # provided in this folder
    except Exception:
        return None
    if hasattr(legacy_impl, "reformat_cv"):
        return legacy_impl.reformat_cv(input_docx, template_docx)  # type: ignore
    return None
