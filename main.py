from gui import AppGUI
from storage import ensure_projects_dir
import yaml, os, sys

def load_config(path="config_example.yaml"):
    with open(path, "r") as f:
        return yaml.safe_load(f)

def main():
    config = load_config()
    ensure_projects_dir(config["storage"]["projects_dir"])
    app = AppGUI(config)
    app.run()

if __name__ == "__main__":
    main()