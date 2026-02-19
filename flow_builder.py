import re
import networkx as nx

class FlowBuilder:
    def __init__(self):
        self.call_re = re.compile(r'\bCall\s+([A-Za-z0-9_]+)|\b([A-Za-z0-9_]+)\s*\(' , re.IGNORECASE)

    def find_procs(self, code):
        procs = []
        for m in re.finditer(r'^\s*(Public|Private)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)', code, re.IGNORECASE | re.MULTILINE):
            procs.append(m.group(3))
        return procs

    def find_calls(self, code):
        calls = set()
        for m in self.call_re.finditer(code):
            name = m.group(1) or m.group(2)
            if name:
                calls.add(name)
        return list(calls)

    def build_flow(self, vba_data):
        G = nx.DiGraph()
        components = vba_data["components"]
        for comp_name, comp in components.items():
            code = comp["code"] or ""
            procs = self.find_procs(code)
            for p in procs:
                G.add_node(p, component=comp_name)
            for p in procs:
                # find calls inside proc body (simple heuristic)
                body_re = re.compile(r'(Sub|Function)\s+' + re.escape(p) + r'[\s\S]*?(?=(Sub|Function|$))', re.IGNORECASE)
                m = body_re.search(code)
                if m:
                    body = m.group(0)
                    calls = self.find_calls(body)
                    for c in calls:
                        if c != p:
                            G.add_edge(p, c)
        # return adjacency and topological order if possible
        try:
            order = list(nx.topological_sort(G))
        except Exception:
            order = list(G.nodes())
        return {"graph": G, "order": order}