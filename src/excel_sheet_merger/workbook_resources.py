"""Own openpyxl 3.1.5 serialization resources, including on handled failures.

Normal worksheet serialization mirrors openpyxl's MIT-licensed ExcelWriter;
only resource lifetime changes. Never enumerate or sweep global TEMP files.
"""
from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl.worksheet._writer import WorksheetWriter
from openpyxl.writer.excel import ExcelWriter


def _dispose_writer(writer) -> None:
    """Best effort; a cleanup refusal must not hide a conversion/I/O error."""
    if writer is None:
        return
    try:
        writer.close()
    except Exception:
        # A failed XML generator may already be closed or only partly started.
        pass
    try:
        writer.cleanup()
    except (OSError, ValueError):
        # Successful serialization may already have removed it. OS-denied
        # removal remains registered in openpyxl for its exit-time cleanup.
        pass


def close_workbook_resources(book) -> None:
    """Finalize only this task's streamed worksheets; safe after save or error."""
    if book is None:
        return
    if book.write_only:
        for sheet in book.worksheets:
            # Capture BEFORE close: do not create a new writer for an empty sheet.
            writer = getattr(sheet, '_writer', None)
            if writer is None:
                continue
            try:
                if not sheet.closed:
                    sheet.close()
            except Exception:
                pass
            _dispose_writer(writer)
    try:
        book.close()
    except Exception:
        # Finalization is not allowed to replace the original job outcome.
        pass


class _OwnedExcelWriter(ExcelWriter):
    def write_worksheet(self, ws):
        # Keep the exact serializer/manifest/drawing behavior of pinned 3.1.5.
        writer = getattr(ws, '_writer', None) if self.workbook.write_only else None
        try:
            ws._drawing = SpreadsheetDrawing()
            ws._drawing.charts = ws._charts
            ws._drawing.images = ws._images
            if self.workbook.write_only:
                if not ws.closed:
                    ws.close()
                writer = ws._writer
            else:
                writer = WorksheetWriter(ws)
                writer.write()
            ws._rels = writer._rels
            self._archive.write(writer.out, ws.path[1:])
            self.manifest.append(ws)
        finally:
            _dispose_writer(writer)


def save_workbook_safely(book, path: Path) -> None:
    """Serialize unchanged workbook content and close the ZIP on every exit."""
    if book.read_only:
        raise TypeError('Workbook is read-only')
    if book.write_only and not book.worksheets:
        book.create_sheet()
    archive = ZipFile(path, 'w', ZIP_DEFLATED, allowZip64=True)
    try:
        # Match openpyxl.writer.excel.save_workbook's metadata update.
        book.properties.modified = datetime.now(timezone.utc).replace(tzinfo=None)
        _OwnedExcelWriter(book, archive).write_data()
    except BaseException:
        try:
            archive.close()
        except Exception:
            pass
        raise
    else:
        # Closing writes the ZIP directory. A failure here must fail the merge,
        # not publish an incomplete workbook as a successful result.
        archive.close()
