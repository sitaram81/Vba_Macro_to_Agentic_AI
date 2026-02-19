import os, json, time, re
from openai import OpenAI
from dotenv import load_dotenv

PROMPT_TEMPLATE = """
You are an expert translator of Excel VBA macros into Python code using pandas and openpyxl/pywin32 where necessary.
Task: Convert the following VBA procedure into a Python function that reproduces the same logic and side effects.
Provide:
1) Python function code only (no explanation) with a clear signature.
2) A short confidence score between 0 and 1 and a one-line reason.
Input metadata:
- Procedure name: {proc_name}
- Component: {component}
- Full VBA code:
{vba_code}
Constraints:
- Use pandas for table-like data, openpyxl for cell-level formulas if needed, and win32com for Excel-specific operations that cannot be reproduced otherwise.
- If the VBA uses Excel-specific objects (Range, Cells, Worksheets), prefer to implement using win32com to call Excel directly, but also provide a pure-Python fallback if feasible.
- Keep the function idempotent and testable.
Return a JSON object with keys: code, confidence, reason.
"""

PROMPT_TEMPLATE2="""
You are an execution agent. Given a project folder with converted Python modules and metadata, perform the following steps:
1. Load metadata.json and identify macro_to_run.
2. If macro_to_run exists, run the original Excel macro via COM and capture a workbook snapshot.
3. Execute converted Python modules in topological order from metadata or flow.json.
4. For each module, capture its output snapshot.
5. Compare Excel snapshot and Python snapshots and produce a JSON report with per-procedure confidence and match status.
6. If any procedure has confidence < 0.6 or mismatch, mark it for manual review and include the code snippet and reason.
Return the JSON report only.
"""

class VBAConverter:
    def __init__(self, config):
        self.config = config
        provider = config["llm"]["provider"]
        if provider == "openai":
            load_dotenv()
            api_key = os.environ.get(config["llm"]["api_key_env"])
            self.client = OpenAI(api_key=api_key)
            self.model = config["llm"].get("model", "gpt-4o-mini")
        else:
            raise NotImplementedError("Only openai provider implemented in this template")

    def convert_proc(self, proc_name, component, vba_code):
        prompt = PROMPT_TEMPLATE.format(proc_name=proc_name, component=component, vba_code=vba_code)
        # call LLM
        resp = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role":"user","content":prompt}],
            max_tokens= self.config["app"].get("max_conversion_tokens", 3000),
            temperature=0.0
        )
        text = resp.choices[0].message.content.strip()
        # Expect JSON at the end; try to parse
        import json
        m = re.search(r'\{[\s\S]*\}\s*$', text)
        if m:
            try:
                j = json.loads(m.group(0))
                return j
            except Exception:
                # fallback: wrap entire text as code with low confidence
                return {"code": text, "confidence": 0.2, "reason": "LLM output not JSON-parsable; returned raw text"}
        else:
            return {"code": text, "confidence": 0.2, "reason": "LLM output missing JSON; returned raw text"}

    def convert_project(self, project_id, vba_data, flow, project_path, progress_callback=None):
        components = vba_data["components"]
        converted_dir = os.path.join(project_path, "converted")
        os.makedirs(converted_dir, exist_ok=True)
        summary = {"procedures": []}
        for comp_name, comp in components.items():
            code = comp.get("code","")
            # find procs
            from flow_builder import FlowBuilder
            fb = FlowBuilder()
            procs = fb.find_procs(code)
            for p in procs:
                if progress_callback:
                    progress_callback(f"Converting {p} in {comp_name} ...")
                # extract proc body
                body_re = re.compile(r'(Sub|Function)\s+' + re.escape(p) + r'[\s\S]*?(?=(Sub|Function|$))', re.IGNORECASE)
                m = body_re.search(code)
                body = m.group(0) if m else code
                res = self.convert_proc(p, comp_name, body)
                fname = f"module_{comp_name}_{p}.py"
                with open(os.path.join(converted_dir, fname), "w", encoding="utf-8") as f:
                    f.write("# Converted from VBA\n")
                    f.write(res["code"])
                summary["procedures"].append({
                    "name": p,
                    "component": comp_name,
                    "file": fname,
                    "confidence": res.get("confidence", 0),
                    "reason": res.get("reason", "")
                })
        return {"summary": summary}