"""PyInstaller용 GUI 실행 진입점 (콘솔 창을 띄우지 않음)."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / 'src'))
from excel_sheet_merger.gui import run_app


if __name__ == '__main__':
    run_app()
