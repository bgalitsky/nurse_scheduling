"""Запуск окна medeid: python run_app.py (или кнопка Run в PyCharm на этом файле)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from medeid.app import main  # noqa: E402

if __name__ == "__main__":
    main()
