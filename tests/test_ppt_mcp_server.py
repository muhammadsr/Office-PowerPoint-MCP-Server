import pytest
import ppt_mcp_server as server

@pytest.fixture(autouse=True)
def reset_session():
    # before each test, clear out any existing session
    server._session = None

def test_create_presentation_returns_one_slide_and_id():
    resp = server.create_presentation()
    assert resp["slide_count"] == 1, "New session should always start with exactly one slide"
    assert "presentation_id" in resp and isinstance(resp["presentation_id"], str)

def test_create_presentation_generates_new_id_each_time():
    first = server.create_presentation()
    second = server.create_presentation()
    assert first["presentation_id"] != second["presentation_id"], "IDs must differ for fresh sessions"
