"""
config.py
Persistent user configuration for the מיכפל automation tools.

Settings (data folder, output folder, ...) are stored as JSON and survive
between runs. The only place default paths live is DEFAULTS below; everything
else reads through load()/get() so nothing downstream hardcodes a path.

The config file lives inside the project folder (next to this module) so it
travels with the app on both Linux (dev) and Windows (deployment):
    <project>/config.json
(override with the MICPAL_CONFIG environment variable if needed.)
"""

import json
import os

# The seed values used the first time the app runs (or for any key the saved
# config is missing). Users change them in the UI; the changes persist. These
# are defaults, not hardcoded usage -- code reads load()/get(), never these.
#
# Defaults differ by OS: the מיכפל data lives on a network share mounted at
# /mnt/Z/Msk8 on the Linux dev box, but reached as the Z: drive on the Windows
# machines the app deploys to.
if os.name == "nt":
    DEFAULTS = {
        "data_dir": r"Z:\Msk8",
        "out_dir": r"Z:\Msk8\company_templates",
    }
else:
    DEFAULTS = {
        "data_dir": "/mnt/Z/Msk8",
        "out_dir": "/mnt/Z/Msk8/company_templates",
    }


def config_path():
    """Absolute path to the JSON config file (honors MICPAL_CONFIG override)."""
    override = os.environ.get("MICPAL_CONFIG")
    if override:
        return override
    project_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(project_dir, "config.json")


def load():
    """
    Return the current config: DEFAULTS overlaid with whatever is saved on
    disk. Unknown/missing keys fall back to their default. Never raises on a
    missing or corrupt file -- it just falls back to defaults.
    """
    cfg = dict(DEFAULTS)
    path = config_path()
    try:
        with open(path, "r", encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            cfg.update({k: v for k, v in saved.items() if k in DEFAULTS})
    except (FileNotFoundError, ValueError, OSError):
        pass
    return cfg


def save(cfg):
    """
    Persist the given config dict (only known keys are written). Creates the
    parent directory if needed. Returns the path written.
    """
    to_write = {k: cfg[k] for k in DEFAULTS if k in cfg}
    path = config_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(to_write, f, indent=2, ensure_ascii=False)
    return path


def get(key):
    """Convenience: a single value from the current (loaded) config."""
    return load()[key]
