
from zipfile import ZipFile
from hashlib import md5

def docx_binary_fingerprint(path: str) -> str:
    members = []
    with ZipFile(path, 'r') as z:
        for name in ["word/document.xml", "word/styles.xml", "word/numbering.xml"]:
            try:
                members.append(z.read(name))
            except KeyError:
                pass
    h = md5()
    for blob in members:
        h.update(blob)
    return h.hexdigest()
