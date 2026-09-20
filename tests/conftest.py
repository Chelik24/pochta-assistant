"""Чтобы тесты видели fetch_mail.py в корне проекта."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
