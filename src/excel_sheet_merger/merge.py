"""Microsoft Excel COM으로 시트를 복사해 임시 저장 후 최종 이름으로 확정."""

from __future__ import annotations

import os
import uuid
from contextlib import contextmanager
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any, Iterator

from .classify import MergeGroup


class MergeStatus(str, Enum):
    CREATED = "생성 완료"
    EXISTS = "기존 파일 존재로 건너뜀"


@dataclass(frozen=True)
class MergeResult:
    path: Path
    status: MergeStatus


def merge_group(group: MergeGroup, output_dir: Path, excel: Any) -> MergeResult:
    """한 그룹만 처리한다. 성공할 때만 최종 경로가 생성된다.

    excel은 Excel.Application COM 객체 또는 같은 인터페이스의 테스트 대역.
    """
    if len(group.inputs) < 2:
        raise ValueError("병합 대상이 2개 미만인 그룹입니다")
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    final_path = output_dir / group.output_name
    if final_path.exists():
        return MergeResult(final_path, MergeStatus.EXISTS)
    tmp_path = output_dir / f"_병합중_{uuid.uuid4().hex}.xlsx"
    target = None
    try:
        target = excel.Workbooks.Add()
        # Excel 설정에 따라 새 통합 문서의 기본 시트가 1~3개일 수 있다.
        empty_sheets = [target.Sheets(i) for i in range(1, int(target.Sheets.Count) + 1)]
        for item in group.inputs:
            source = None
            try:
                source = excel.Workbooks.Open(
                    Filename=str(item.path.resolve()),
                    UpdateLinks=0,
                    ReadOnly=True,
                    IgnoreReadOnlyRecommended=True,
                    AddToMru=False,
                )
                if int(source.Sheets.Count) != 1:
                    raise ValueError(f"원본 시트 수가 1개가 아닙니다: {item.file_name}")
                source.Sheets(1).Copy(After=target.Sheets(int(target.Sheets.Count)))
                target.Sheets(int(target.Sheets.Count)).Name = item.sheet_name
            finally:
                if source is not None:
                    source.Close(SaveChanges=False)
        for initial in empty_sheets:
            initial.Delete()
        target.SaveAs(Filename=str(tmp_path.resolve()), FileFormat=51)  # xlOpenXMLWorkbook (.xlsx)
        target.Close(SaveChanges=False)
        target = None
        if final_path.exists():
            return MergeResult(final_path, MergeStatus.EXISTS)
        # Windows의 os.rename은 이미 있는 목적지 파일을 덮어쓰지 않는다.
        # 임시 파일과 목적지는 같은 폴더라 동일 볼륨이다.
        try:
            os.rename(tmp_path, final_path)
        except FileExistsError:
            return MergeResult(final_path, MergeStatus.EXISTS)
        return MergeResult(final_path, MergeStatus.CREATED)
    finally:
        if target is not None:
            try:
                target.Close(SaveChanges=False)
            except Exception:
                pass
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


@contextmanager
def excel_session() -> Iterator[Any]:
    """별도 Excel 프로세스를 생성, 매크로/링크를 비활성화하고 종료한다.

    Tkinter 작업 스레드에서 호출하므로 COM STA를 스레드 내부에서 초기화한다.
    """
    if os.name != "nt":
        raise RuntimeError("Excel 자동병합은 Microsoft Excel이 설치된 Windows에서만 실행할 수 있습니다")
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise RuntimeError("Windows Excel 구성요소(pywin32)가 없습니다. EXE를 다시 빌드해야 합니다") from exc

    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.DispatchEx("Excel.Application")
        except Exception as exc:
            raise RuntimeError("Microsoft Excel 데스크톱 프로그램을 시작할 수 없습니다") from exc
        # 매크로 실행 차단이 실패하면 안전을 위해 아무 파일도 열지 않는다.
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        app.Visible = False
        app.DisplayAlerts = False
        app.ScreenUpdating = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        yield app
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
