import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from vba_extractor import VBAExtractor
from converter import VBAConverter
from executor import AgentExecutor
from storage import ProjectStore
import threading, os

class AppGUI:
    def __init__(self, config):
        self.config = config
        self.root = tk.Tk()
        self.root.title("VBA to Python Migration Agent")
        self.store = ProjectStore(config["storage"]["projects_dir"])
        self._build_ui()

    def _build_ui(self):
        frm = ttk.Frame(self.root, padding=12)
        frm.grid(sticky="nsew")
        ttk.Button(frm, text="Select Excel Macro File", command=self.select_file).grid(row=0, column=0, sticky="w")
        self.file_label = ttk.Label(frm, text="No file selected")
        self.file_label.grid(row=0, column=1, sticky="w")
        ttk.Button(frm, text="Open Project", command=self.open_project).grid(row=1, column=0, sticky="w")
        self.project_combo = ttk.Combobox(frm, values=self.store.list_projects())
        self.project_combo.grid(row=1, column=1, sticky="w")
        ttk.Button(frm, text="Convert Selected Project", command=self.convert_project).grid(row=2, column=0, sticky="w")
        ttk.Button(frm, text="Run Agentic Workflow", command=self.run_workflow).grid(row=2, column=1, sticky="w")
        self.log = tk.Text(frm, height=20, width=100)
        self.log.grid(row=3, column=0, columnspan=2)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def log_write(self, text):
        self.log.insert("end", text + "\n")
        self.log.see("end")

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsm *.xls *.xla")])
        if not path:
            return
        self.file_label.config(text=path)
        self.log_write(f"Selected {path}")
        # create project
        project_id = self.store.create_project_from_file(path)
        self.project_combo["values"] = self.store.list_projects()
        self.project_combo.set(project_id)
        self.log_write(f"Project created: {project_id}")

    def open_project(self):
        pid = self.project_combo.get()
        if not pid:
            messagebox.showinfo("Select project", "Choose a project first")
            return
        meta = self.store.load_metadata(pid)
        self.log_write(f"Project metadata: {meta}")

    def convert_project(self):
        pid = self.project_combo.get()
        if not pid:
            messagebox.showinfo("Select project", "Choose a project first")
            return
        threading.Thread(target=self._convert_thread, args=(pid,), daemon=True).start()

    def _convert_thread(self, pid):
        self.log_write("Starting extraction...")
        extractor = VBAExtractor(self.config)
        project_path = self.store.project_path(pid)
        vba_data = extractor.extract_all(project_path)
        self.log_write("Extraction complete. Building flow...")
        from flow_builder import FlowBuilder
        fb = FlowBuilder()
        flow = fb.build_flow(vba_data)
        self.log_write("Flow built. Converting via LLM...")
        conv = VBAConverter(self.config)
        conv_results = conv.convert_project(pid, vba_data, flow, project_path, progress_callback=self.log_write)
        self.log_write("Conversion finished.")
        self.store.save_metadata(pid, {"converted": True, "conversion_summary": conv_results["summary"]})
        self.log_write("Saved conversion metadata.")

    def run_workflow(self):
        pid = self.project_combo.get()
        if not pid:
            messagebox.showinfo("Select project", "Choose a project first")
            return
        threading.Thread(target=self._run_thread, args=(pid,), daemon=True).start()

    def _run_thread(self, pid):
        self.log_write("Starting agentic execution...")
        executor = AgentExecutor(self.config, self.store)
        result = executor.run_project(pid, progress_callback=self.log_write)
        self.log_write("Execution finished.")
        self.log_write(str(result))

    def on_close(self):
        self.root.destroy()

    def run(self):
        self.root.mainloop()