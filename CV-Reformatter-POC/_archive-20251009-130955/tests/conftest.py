
import pytest, os
FIX = os.path.join(os.path.dirname(__file__), "fixtures")
def require_fixture(name: str) -> str:
    path = os.path.join(FIX, name)
    if not os.path.exists(path):
        pytest.skip(f"Missing fixture: {name}. Put it in tests/fixtures.")
    return path
