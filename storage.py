import os, shutil, uuid, json

class ProjectStore:
    def __init__(self, base_dir):
        self.base_dir = base_dir
        os.makedirs(self.base_dir, exist_ok=True)

    def list_projects(self):
        return [d for d in os.listdir(self.base_dir) if os.path.isdir(os.path.join(self.base_dir, d))]

    def create_project_from_file(self, file_path):
        pid = str(uuid.uuid4())[:8]
        proj_dir = os.path.join(self.base_dir, pid)
        os.makedirs(proj_dir, exist_ok=True)
        dest = os.path.join(proj_dir, "original.xlsm")
        shutil.copyfile(file_path, dest)
        meta = {"id": pid, "original_filename": os.path.basename(file_path)}
        with open(os.path.join(proj_dir, "metadata.json"), "w", encoding="utf-8") as f:
            json.dump(meta, f, indent=2)
        os.makedirs(os.path.join(proj_dir, "converted"), exist_ok=True)
        os.makedirs(os.path.join(proj_dir, "logs"), exist_ok=True)
        return pid

    def project_path(self, pid):
        return os.path.join(self.base_dir, pid)

    def load_metadata(self, pid):
        p = os.path.join(self.project_path(pid), "metadata.json")
        if not os.path.exists(p):
            return {}
        import json
        with open(p, "r", encoding="utf-8") as f:
            return json.load(f)

    def save_metadata(self, pid, meta):
        p = os.path.join(self.project_path(pid), "metadata.json")
        existing = self.load_metadata(pid)
        existing.update(meta)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)

def ensure_projects_dir(path):
    os.makedirs(path, exist_ok=True)