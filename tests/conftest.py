import importlib.util
import os
import sys
import types

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@pytest.fixture(scope="session")
def erg():
    # easygui needs a display; the functions under test don't use it, so stub it out
    sys.modules.setdefault("easygui", types.ModuleType("easygui"))
    spec = importlib.util.spec_from_file_location("expense_rg", os.path.join(ROOT, "Expense-RG.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module
