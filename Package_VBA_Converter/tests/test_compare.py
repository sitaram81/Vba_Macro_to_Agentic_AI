# Simple unit test to validate snapshot comparator
from executor import AgentExecutor
from storage import ProjectStore
import tempfile, os

def test_compare_basic():
    tmp = tempfile.mkdtemp()
    store = ProjectStore(tmp)
    # create fake project
    pid = store.create_project_from_file(__file__)  # just copy this file as placeholder
    # write fake metadata
    store.save_metadata(pid, {"macro_to_run": None})
    # instantiate executor
    config = {"app": {}, "storage": {"projects_dir": tmp}}
    exec = AgentExecutor(config, store)
    # call compare directly
    excel_snap = {"Sheet1": [["1","2"],["3","4"]]}
    python_snap = {"module_a.py": {"Sheet1": [["1","2"],["3","4"]]}}
    comp = exec._compare_snapshots(excel_snap, python_snap)
    assert comp and comp[0]["match"]