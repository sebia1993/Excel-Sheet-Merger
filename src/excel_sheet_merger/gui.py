"""Windows 11 Korean CSV/Excel → XLSX GUI; theme-only redesign, unchanged engine."""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .classify import MergeGroup, ScanReport, scan_folder
from .merge import MergeStatus, merge_group
from .ui_theme import COLORS, FONT, configure_theme


def summarize_scan(report: ScanReport) -> str:
    files = sum(len(group.inputs) for group in report.groups)
    return (f"검색 파일 {report.checked}개  |  병합 그룹 {len(report.groups)}개  |  "
            f"병합 대상 {files}개  |  단독 파일 {len(report.singles)}개  |  제외 {len(report.rejected)}개")


def summary_metrics(report: ScanReport) -> tuple[str, str, str, str]:
    """Values displayed on the four summary cards (includes rejected inputs)."""
    return (str(report.checked), str(len(report.groups)),
            str(len(report.singles)), str(len(report.rejected)))


class MergeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Excel Sheet Merger  |  CSV·Excel → XLSX")
        self.geometry("1100x760")
        self.minsize(900, 640)
        self.configure(bg=COLORS["background"])
        self.folder = tk.StringVar(value="")
        self.summary = tk.StringVar(value="폴더를 선택하면 병합 결과와 시간별 시트를 미리 확인할 수 있습니다.")
        self.state_text = tk.StringVar(value="원본 CSV·Excel 파일이 있는 폴더를 선택하세요.")
        self.metrics = {key: tk.StringVar(value="0") for key in ("checked", "groups", "singles", "rejected")}
        self.report: ScanReport | None = None
        self.group_map: dict[str, MergeGroup] = {}
        self.events: queue.Queue[tuple] = queue.Queue()
        self.busy = False
        self.log_expanded = False
        self._style()
        self._layout()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._poll_job = self.after(120, self._consume_events)

    def _style(self) -> None:
        self.theme_style = configure_theme(self)

    def _card(self, parent: tk.Misc, padding: int = 15) -> tuple[tk.Frame, ttk.Frame]:
        """Bordered content surface, independent of platform-specific corner APIs."""
        frame = tk.Frame(parent, bg=COLORS["border"], padx=1, pady=1)
        body = ttk.Frame(frame, padding=padding, style="Office.Card.TFrame")
        body.pack(fill="both", expand=True)
        return frame, body

    def _layout(self) -> None:
        # The footer and log reserve their space; the preview expands into the rest.
        root = ttk.Frame(self, padding=(20, 12, 20, 12), style="Office.Page.TFrame")
        root.pack(fill="both", expand=True)
        self._header(root)
        self._folder_card(root)
        self._metric_cards(root)
        self._footer(root)
        self._log_card(root)
        self._preview_card(root)
        self._set_enabled(False)
        self._write_log("안내: 원본 CSV·Excel을 변경하지 않고 시간별 시트가 있는 XLSX로 병합합니다. 기존 결과는 건너뜁니다.")

    def _header(self, root: ttk.Frame) -> None:
        header = ttk.Frame(root, style="Office.Page.TFrame")
        header.pack(fill="x", pady=(0, 10))
        icon = tk.Canvas(header, width=40, height=40, bg=COLORS["primary"], highlightthickness=0)
        icon.pack(side="left", padx=(0, 12))
        # Original four-column document glyph drawn without third-party assets.
        for x in (12, 20, 28):
            icon.create_line(x, 9, x, 32, fill="#FFFFFF", width=1)
        for y in (9, 17, 25, 32):
            icon.create_line(12, y, 28, y, fill="#FFFFFF", width=1)
        label_group = ttk.Frame(header, style="Office.Page.TFrame")
        label_group.pack(side="left", fill="x", expand=True)
        ttk.Label(label_group, text="Excel Sheet Merger", style="Office.Title.TLabel").pack(anchor="w")
        ttk.Label(label_group, text="CSV·Excel 파일을 시간별 시트로 자동 병합 · Excel 설치 불필요", style="Office.Subtitle.TLabel").pack(anchor="w")
        # Label pill has its own background independent of page theme.
        pill = ttk.Frame(header, style="Office.Subtle.TFrame")
        pill.pack(side="right")
        ttk.Label(pill, text="CSV·Excel  →  XLSX", style="Office.Pill.TLabel").pack()

    def _folder_card(self, root: ttk.Frame) -> None:
        frame, card = self._card(root, 10)
        frame.pack(fill="x", pady=(0, 7))
        heading = ttk.Frame(card, style="Office.Card.TFrame")
        heading.pack(fill="x", pady=(0, 5))
        ttk.Label(heading, text="01  원본 폴더", style="Office.Heading.TLabel").pack(side="left")
        ttk.Label(heading, text="원본 파일은 변경하지 않습니다", style="Office.Hint.TLabel").pack(side="right")
        file_row = ttk.Frame(card, style="Office.Card.TFrame")
        file_row.pack(fill="x")
        file_row.columnconfigure(0, weight=1)
        self.folder_entry = ttk.Entry(file_row, textvariable=self.folder, style="Office.Path.TEntry", state="readonly")
        self.folder_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.pick_button = ttk.Button(file_row, text="폴더 선택", style="Office.Secondary.TButton", command=self._pick_folder)
        self.pick_button.grid(row=0, column=1, padx=(0, 8))
        self.scan_button = ttk.Button(file_row, text="파일 분석", style="Office.Secondary.TButton", command=self._scan)
        self.scan_button.grid(row=0, column=2)

    def _metric_cards(self, root: ttk.Frame) -> None:
        strip = ttk.Frame(root, style="Office.Page.TFrame")
        strip.pack(fill="x", pady=(0, 7))
        captions = (("checked", "검색 파일"), ("groups", "병합 그룹"),
                    ("singles", "단독 파일"), ("rejected", "제외 파일"))
        for i, (key, caption) in enumerate(captions):
            strip.columnconfigure(i, weight=1, uniform="metrics")
            frame, card = self._card(strip, 7)
            frame.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 5, 0 if i == 3 else 5))
            ttk.Label(card, textvariable=self.metrics[key], style="Office.Metric.TLabel").pack(anchor="w")
            ttk.Label(card, text=caption, style="Office.MetricHint.TLabel").pack(anchor="w")

    def _preview_card(self, root: ttk.Frame) -> None:
        frame, card = self._card(root, 9)
        frame.pack(side="top", fill="both", expand=True)
        header = ttk.Frame(card, style="Office.Card.TFrame")
        header.pack(fill="x", pady=(0, 6))
        ttk.Label(header, text="02  병합 대상 미리보기", style="Office.Heading.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.summary, style="Office.Hint.TLabel", wraplength=580).pack(side="right")
        table_container = ttk.Frame(card, style="Office.Card.TFrame")
        table_container.pack(fill="both", expand=True)
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_container, style="Office.Treeview", columns=("sheet", "source"), show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="병합 결과 파일 / 대상")
        self.tree.heading("sheet", text="시트 이름 / 개수")
        self.tree.heading("source", text="원본 CSV·Excel / 제외 사유")
        self.tree.column("#0", width=355, minwidth=230)
        self.tree.column("sheet", width=150, minwidth=105, anchor="center", stretch=False)
        self.tree.column("source", width=430, minwidth=270)
        self.tree.tag_configure("group", foreground=COLORS["primary"], background=COLORS["surface_subtle"], font=(FONT, 9, "bold"))
        self.tree.tag_configure("single", foreground=COLORS["warning"])
        self.tree.tag_configure("rejected", foreground=COLORS["error"])
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)
        ttk.Label(card, text="하위 시트를 선택하면 해당 그룹 전체가 병합됩니다.", style="Office.Hint.TLabel").pack(anchor="w", pady=(5, 0))

    def _log_card(self, root: ttk.Frame) -> None:
        frame, card = self._card(root, 4)
        frame.pack(side="bottom", fill="x", pady=(6, 6))
        self.log_toggle = ttk.Button(card, text="작업 기록 보기  ▾", style="Office.Link.TButton", command=self._toggle_log)
        self.log_toggle.pack(anchor="w")
        self.log_panel = ttk.Frame(card, style="Office.Card.TFrame")
        self.log = tk.Text(self.log_panel, height=4, wrap="word", font=(FONT, 9), borderwidth=0,
                           background=COLORS["surface"], foreground=COLORS["text"], padx=7, pady=5,
                           selectbackground=COLORS["surface_subtle"], selectforeground=COLORS["text"])
        log_scroll = ttk.Scrollbar(self.log_panel, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=log_scroll.set)
        self.log.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log.configure(state="disabled")

    def _footer(self, root: ttk.Frame) -> None:
        frame, card = self._card(root, 8)
        frame.pack(side="bottom", fill="x")
        progress_line = ttk.Frame(card, style="Office.Card.TFrame")
        progress_line.pack(fill="x", pady=(0, 6))
        ttk.Label(progress_line, text="03  병합 실행", style="Office.Heading.TLabel").pack(side="left")
        ttk.Label(progress_line, textvariable=self.state_text, style="Office.Status.TLabel").pack(side="right")
        self.progress = ttk.Progressbar(card, style="Office.Horizontal.TProgressbar", mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x", pady=(0, 7))
        action_row = ttk.Frame(card, style="Office.Card.TFrame")
        action_row.pack(fill="x")
        self.all_button = ttk.Button(action_row, text="전체 그룹 병합", style="Office.Primary.TButton", command=lambda: self._start_merge(False))
        self.all_button.pack(side="right")
        self.selected_button = ttk.Button(action_row, text="선택 그룹 병합", style="Office.Secondary.TButton", command=lambda: self._start_merge(True))
        self.selected_button.pack(side="right", padx=(0, 8))
        self.open_button = ttk.Button(action_row, text="결과 폴더 열기", style="Office.Secondary.TButton", command=self._open_results)
        self.open_button.pack(side="left")

    def _toggle_log(self) -> None:
        if self.log_expanded:
            self.log_panel.pack_forget()
            self.log_toggle.configure(text="작업 기록 보기  ▾")
            self.log_expanded = False
        else:
            self.log_panel.pack(fill="x", pady=(3, 0))
            self.log_toggle.configure(text="작업 기록 접기  ▴")
            self.log_expanded = True

    def _show_log(self) -> None:
        if not self.log_expanded:
            self._toggle_log()

    def _set_enabled(self, can_merge: bool) -> None:
        state = "normal" if can_merge and not self.busy else "disabled"
        self.all_button.configure(state=state)
        self.selected_button.configure(state=state)
        self.scan_button.configure(state="normal" if self.folder.get().strip() and not self.busy else "disabled")
        self.pick_button.configure(state="disabled" if self.busy else "normal")

    def _write_log(self, msg: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _pick_folder(self) -> None:
        result = filedialog.askdirectory(title="원본 CSV·Excel 파일이 있는 폴더 선택", mustexist=True)
        if result:
            self.folder.set(result)
            self._scan()

    def _scan(self) -> None:
        location = self.folder.get().strip()
        if not location:
            messagebox.showinfo("폴더 선택", "먼저 원본 CSV·Excel 폴더를 선택해 주세요.")
            return
        try:
            report = scan_folder(Path(location))
        except Exception as exc:
            messagebox.showerror("파일 분석 실패", str(exc))
            return
        self.report = report
        self.group_map.clear()
        for entry in self.tree.get_children():
            self.tree.delete(entry)
        for idx, group in enumerate(report.groups):
            gid = f"g{idx}"
            self.group_map[gid] = group
            self.tree.insert("", "end", iid=gid, text=group.output_name,
                             values=(f"{len(group.inputs)}개 시트", "병합 가능"), tags=("group",), open=True)
            for j, item in enumerate(group.inputs):
                self.tree.insert(gid, "end", iid=f"{gid}i{j}", text=item.sheet_name,
                                 values=(item.sheet_name, item.file_name))
        if report.singles:
            sid = self.tree.insert("", "end", text="단독 파일 · 병합 보류",
                                   values=(f"{len(report.singles)}개", "파일 1개인 그룹은 병합하지 않음"), tags=("single",), open=True)
            for group in report.singles:
                self.tree.insert(sid, "end", text=group.output_name, values=("대상 1개", group.inputs[0].file_name), tags=("single",))
        if report.rejected:
            xid = self.tree.insert("", "end", text="인식 제외 파일",
                                   values=(f"{len(report.rejected)}개", "병합하지 않음"), tags=("rejected",), open=True)
            for rejected in report.rejected:
                self.tree.insert(xid, "end", text=rejected.name, values=("제외", rejected.reason), tags=("rejected",))
        for key, value in zip(("checked", "groups", "singles", "rejected"), summary_metrics(report)):
            self.metrics[key].set(value)
        self.summary.set(f"병합 대상 {sum(len(g.inputs) for g in report.groups)}개 파일 · 결과 {len(report.groups)}개")
        self.progress.configure(maximum=1, value=0)
        self._set_enabled(bool(report.groups))
        self.state_text.set("미리보기 준비 완료" if report.groups else "병합할 그룹 없음 · 단독/제외 파일 확인")
        self._write_log(summarize_scan(report))

    def _selected_groups(self) -> list[MergeGroup]:
        selected = set()
        for item in self.tree.selection():
            node = item
            while node and node not in self.group_map:
                node = self.tree.parent(node)
            if node in self.group_map:
                selected.add(node)
        return [g for gid, g in self.group_map.items() if gid in selected]

    def _start_merge(self, selection_only: bool) -> None:
        if self.busy or self.report is None:
            return
        groups = self._selected_groups() if selection_only else list(self.report.groups)
        if not groups:
            messagebox.showinfo("병합 선택", "병합할 그룹을 선택해 주세요." if selection_only else "병합 대상 그룹이 없습니다.")
            return
        output_dir = Path(self.folder.get()) / "병합결과"
        self.busy = True
        self._set_enabled(False)
        self.progress.configure(maximum=len(groups), value=0)
        self.state_text.set(f"병합 진행 0/{len(groups)}")
        self._write_log(f"병합 시작: {len(groups)}개 그룹 → {output_dir}")
        thread = threading.Thread(target=self._worker_merge, args=(groups, output_dir), daemon=True)
        thread.start()

    def _worker_merge(self, groups: list[MergeGroup], output_dir: Path) -> None:
        completed = 0
        failed = 0
        skipped = 0
        records = [f"병합 실행: {datetime.now():%Y-%m-%d %H:%M:%S}", f"결과 폴더: {output_dir}"]
        if self.report:
            records += [f"제외: {e.name} | {e.reason}" for e in self.report.rejected]
            records += [f"단독 보류: {g.inputs[0].file_name}" for g in self.report.singles]
        try:
            output_dir.mkdir(exist_ok=True, parents=True)
            for index, group in enumerate(groups, 1):
                completed, failed, skipped = self._process_group(
                    group, output_dir, index, len(groups), completed, failed, skipped, records
                )
        except Exception as exc:
            failed += 1
            item = f"병합 작업 실패: {exc}"
            records.append(item)
            self.events.put(("log", item))
        finally:
            try:
                output_dir.mkdir(exist_ok=True, parents=True)
                log_file = output_dir / f"병합작업기록_{datetime.now():%Y%m%d_%H%M%S_%f}.txt"
                log_file.write_text("\n".join(records) + "\n", encoding="utf-8-sig")
                self.events.put(("log", f"기록 파일: {log_file.name}"))
            except Exception as exc:
                self.events.put(("log", f"작업 기록 저장 실패: {exc}"))
            self.events.put(("done", completed, failed, skipped))

    def _process_group(self, group, output_dir, index, total, completed, failed, skipped, records):
        try:
            outcome = merge_group(group, output_dir)
            if outcome.status == MergeStatus.CREATED:
                completed += 1
                item = f"완료: {group.output_name} / 시트 {len(group.inputs)}개"
            else:
                skipped += 1
                item = f"건너뜀 (이미 존재): {group.output_name}"
        except Exception as exc:
            failed += 1
            item = f"실패: {group.output_name} | {exc}"
        records.append(item)
        self.events.put(("log", item))
        self.events.put(("progress", index, total))
        return completed, failed, skipped

    def _consume_events(self) -> None:
        # Cancel the previous callback when invoked directly by a test or UI.
        # A normal scheduled invocation has already left the Tcl queue.
        previous = getattr(self, "_poll_job", None)
        if previous is not None:
            try:
                self.after_cancel(previous)
            except tk.TclError:
                pass
            self._poll_job = None
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    self._write_log(event[1])
                elif event[0] == "progress":
                    self.progress.configure(maximum=event[2], value=event[1])
                    self.state_text.set(f"병합 진행 {event[1]}/{event[2]}")
                elif event[0] == "done":
                    self.busy = False
                    completed, failed, skipped = event[1:4]
                    self._set_enabled(bool(self.report and self.report.groups))
                    self.state_text.set(f"완료 {completed}개 / 실패 {failed}개 / 건너뜀 {skipped}개")
                    self._write_log(self.state_text.get())
                    if failed:
                        self._show_log()
                        messagebox.showwarning("병합 종료", self.state_text.get() + "\n작업 기록을 확인하세요.")
                    else:
                        messagebox.showinfo("병합 종료", self.state_text.get() + "\n결과는 '병합결과' 폴더에 저장했습니다.")
        except queue.Empty:
            pass
        if self.winfo_exists():
            self._poll_job = self.after(120, self._consume_events)

    def _open_results(self) -> None:
        if not self.folder.get():
            messagebox.showinfo("결과 폴더", "먼저 원본 폴더를 선택해 주세요.")
            return
        destination = Path(self.folder.get()) / "병합결과"
        destination.mkdir(parents=True, exist_ok=True)
        if os.name == "nt":
            os.startfile(str(destination))  # type: ignore[attr-defined]
        else:
            messagebox.showinfo("결과 폴더", str(destination))

    def destroy(self) -> None:
        # Prevent stale Tcl after callbacks when closing/reopening windows.
        job = getattr(self, "_poll_job", None)
        if job is not None:
            try:
                self.after_cancel(job)
            except tk.TclError:
                pass
            self._poll_job = None
        super().destroy()

    def _close(self) -> None:
        if self.busy:
            messagebox.showwarning("병합 중", "병합 중에는 종료할 수 없습니다. 작업이 완료된 후 닫아 주세요.")
        else:
            self.destroy()


def run_app() -> None:
    MergeApp().mainloop()
