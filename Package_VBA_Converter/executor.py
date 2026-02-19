import importlib.util, os, json, sys, traceback
from vba_extractor import VBAExtractor
import win32com.client
import pandas as pd

class AgentExecutor:
    def __init__(self, config, store):
        self.config = config
        self.store = store

    def run_project(self, project_id, progress_callback=None):
        project_path = self.store.project_path(project_id)
        meta = self.store.load_metadata(project_id)
        converted_dir = os.path.join(project_path, "converted")
        # 1) Run original Excel macros and capture outputs (user must specify which macro to run)
        # For demo, we run a top-level Sub named Auto_Open or a user-selected macro
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(os.path.join(project_path, "original.xlsm"))
        # find macros
        macros = excel.Application.Run  # placeholder; we will list macros via VBProject
        # For safety, we will not auto-run arbitrary macros. Instead, we will call a user-specified macro name from metadata.
        macro_to_run = meta.get("macro_to_run")
        results = {"excel": None, "python": None, "comparison": None}
        if macro_to_run:
            try:
                if progress_callback: progress_callback(f"Running Excel macro {macro_to_run}")
                excel.Application.Run(f"'{wb.Name}'!{macro_to_run}")
                # After run, capture workbook state snapshot (e.g., sheet values)
                snapshot = self._snapshot_workbook(wb)
                results["excel"] = snapshot
            except Exception as e:
                results["excel_error"] = str(e)
        else:
            results["excel_note"] = "No macro_to_run specified in metadata; skipping Excel run."

        # 2) Run converted Python functions in order
        py_snapshot = {}
        for fname in os.listdir(converted_dir):
            if not fname.endswith(".py"): continue
            path = os.path.join(converted_dir, fname)
            try:
                if progress_callback: progress_callback(f"Executing {fname}")
                spec = importlib.util.spec_from_file_location(fname[:-3], path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                # convention: converted module should expose a function named run or main
                if hasattr(mod, "run"):
                    out = mod.run(project_path=project_path)
                elif hasattr(mod, "main"):
                    out = mod.main(project_path=project_path)
                else:
                    out = None
                py_snapshot[fname] = out
            except Exception as e:
                py_snapshot[fname] = {"error": str(e), "trace": traceback.format_exc()}
        results["python"] = py_snapshot

        # 3) Compare snapshots (simple equality for sheet values)
        if results.get("excel") and results.get("python"):
            comp = self._compare_snapshots(results["excel"], results["python"])
            results["comparison"] = comp
        wb.Close(False)
        excel.Quit()
        return results

    def _snapshot_workbook(self, wb):
        data = {}
        for sh in wb.Worksheets:
            used = sh.UsedRange
            rows = used.Rows.Count
            cols = used.Columns.Count
            arr = []
            for r in range(1, rows+1):
                row = []
                for c in range(1, cols+1):
                    row.append(str(used.Cells(r,c).Value))
                arr.append(row)
            data[sh.Name] = arr
        return data

    def _compare_snapshots(self, excel_snap, python_snap):
        # Very simple comparator: check if any converted module returned a snapshot matching any sheet
        matches = []
        for mod, out in python_snap.items():
            if isinstance(out, dict):
                # try to match by sheet name keys
                for sheet, arr in excel_snap.items():
                    if out.get(sheet) == arr:
                        matches.append({"module": mod, "sheet": sheet, "match": True})
        return matches