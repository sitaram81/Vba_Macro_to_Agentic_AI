import win32com.client
import os, json, re
from collections import defaultdict

class VBAExtractor:
    def __init__(self, config):
        self.config = config

    def extract_all(self, project_path):
        """
        Opens the workbook via COM and extracts:
        - VBComponents (modules, classes, userforms)
        - Sheet code modules
        - ThisWorkbook code
        - Named formulas
        - Cell formulas for each sheet (optionally limited)
        Returns a dict with all content.
        """
        wb_path = os.path.join(project_path, "original.xlsm")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(wb_path, ReadOnly=True)
        vba_project = wb.VBProject
        components = {}
        for comp in vba_project.VBComponents:
            try:
                code = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            except Exception:
                code = ""
            components[comp.Name] = {
                "type": comp.Type,
                "code": code
            }
        # sheet formulas
        sheets = {}
        for sh in wb.Worksheets:
            # collect named ranges and formulas (limited to used range)
            used = sh.UsedRange
            formulas = {}
            if used is not None:
                rows = used.Rows.Count
                cols = used.Columns.Count
                for r in range(1, rows+1):
                    for c in range(1, cols+1):
                        cell = used.Cells(r, c)
                        f = cell.Formula
                        if f:
                            formulas[f"{r},{c}"] = f
            sheets[sh.Name] = {"formulas": formulas}
        # named formulas
        names = {}
        for nm in wb.Names:
            try:
                names[nm.Name] = nm.RefersTo
            except Exception:
                names[nm.Name] = None
        metadata = {
            "components": components,
            "sheets": sheets,
            "names": names,
            "path": wb_path
        }
        wb.Close(False)
        excel.Quit()
        return metadata