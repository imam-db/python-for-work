"""Load utility scripts from dashed folders as importable modules."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

REPO_ROOT = Path(__file__).resolve().parents[1]


def load_module(folder: str, filename: str) -> ModuleType:
    """Import a script from ``folder/filename`` under a unique module name.

    Each utility folder makes its own ``common`` imports work by adding the
    repository root to ``sys.path``. The folder itself is added as well so
    scripts can use sibling modules if they ever need to.
    """
    module_path = REPO_ROOT / folder / filename
    if not module_path.exists():
        raise FileNotFoundError(f"Utility script not found: {module_path}")

    alias = f"pfw_{folder.replace('-', '_')}_{module_path.stem}"
    cached = sys.modules.get(alias)
    if cached is not None:
        return cached

    for directory in (REPO_ROOT, REPO_ROOT / folder):
        path = str(directory)
        if path not in sys.path:
            sys.path.insert(0, path)

    spec = importlib.util.spec_from_file_location(alias, module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot load module from {module_path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[alias] = module
    spec.loader.exec_module(module)
    return module
