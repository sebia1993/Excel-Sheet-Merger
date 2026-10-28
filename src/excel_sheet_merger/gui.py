"""Windows 11용 한국어 Tkinter CSV → XLSX GUI — 작업 스레드와 UI 이벤트 분리."""

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


def summarize_scan(report: ScanReport) -> str:
    files = sum(len(group.inputs) for group in report.groups)
    return (f"검색 파일 {report.checked}개  |  병합 그룹 {len(report.groups)}개  |  "
            f"병합 대상 {files}개  |  단독 파일 {len(report.singles)}개  |  제외 {len(report.rejected)}개")


class MergeApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("CSV 시트 자동병합  |  Windows 11")
        self.geometry("1050x740")
        self.minsize(840, 620)
        self.configure(bg="#f4f7fb")
        self.folder = tk.StringVar(value="")
        self.summary = tk.StringVar(value="CSV 파일이 있는 폴더를 선택해 주세요.")
        self.state_text = tk.StringVar(value="준비됨")
        self.report: ScanReport | None = None
        self.group_map: dict[str, MergeGroup] = {}
        self.events: queue.Queue[tuple] = queue.Queue()
        self.busy = False
        self._style()
        self._layout()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(120, self._consume_events)

    def _style(self) -> None:
        style = ttk.Style(self)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("App.TFrame", background="#f4f7fb")
        style.configure("Head.TLabel", background="#f4f7fb", foreground="#20304c", font=("Malgun Gothic", 19, "bold"))
        style.configure("Body.TLabel", background="#f4f7fb", foreground="#455468", font=("Malgun Gothic", 9))
        style.configure("Hint.TLabel", background="#f4f7fb", foreground="#64748b", font=("Malgun Gothic", 9))
        style.configure("Big.TButton", font=("Malgun Gothic", 10, "bold"), padding=(11, 9))
        style.configure("TButton", font=("Malgun Gothic", 9), padding=(8, 6))
        style.configure("Treeview", rowheight=25, font=("Malgun Gothic", 9))
        style.configure("Treeview.Heading", font=("Malgun Gothic", 9, "bold"))
        style.configure("TLabelFrame.Label", font=("Malgun Gothic", 10, "bold"))

    def _layout(self) -> None:
        root = ttk.Frame(self, padding=18, style="App.TFrame")
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="CSV → Excel 자동 병합", style="Head.TLabel").pack(anchor="w")
        ttk.Label(root, text="CSV 자동 분류 → 시간별 시트가 있는 XLSX 생성 | Excel 설치 불필요 · 원본 보존", style="Hint.TLabel").pack(anchor="w", pady=(1, 12))

        controls = ttk.LabelFrame(root, text="1  원본 폴더 선택", padding=10)
        controls.pack(fill="x", pady=(0, 10))
        controls.columnconfigure(0, weight=1)
        ttk.Entry(controls, textvariable=self.folder, font=("Malgun Gothic", 10), state="readonly").grid(row=0, column=0, sticky="ew", padx=(0, 7))
        self.pick_button = ttk.Button(controls, text="폴더 선택", command=self._pick_folder)
        self.pick_button.grid(row=0, column=1, padx=(0, 5))
        self.scan_button = ttk.Button(controls, text="파일 분석", command=self._scan)
        self.scan_button.grid(row=0, column=2)

        preview = ttk.LabelFrame(root, text="2  병합 대상 미리보기 (선택 병합은 그룹 또는 해당 하위 시트 선택)", padding=10)
        preview.pack(fill="both", expand=True, pady=(0, 10))
        ttk.Label(preview, textvariable=self.summary, style="Body.TLabel").pack(anchor="w", pady=(0, 8))
        table_container = ttk.Frame(preview)
        table_container.pack(fill="both", expand=True)
        table_container.columnconfigure(0, weight=1)
        table_container.rowconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_container, columns=("sheet", "source"), show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="병합 결과 파일 / 대상")
        self.tree.heading("sheet", text="시트 이름 / 개수")
        self.tree.heading("source", text="원본 파일명 / 제외 사유")
        self.tree.column("#0", width=340, minwidth=200)
        self.tree.column("sheet", width=140, anchor="center", minwidth=105, stretch=False)
        self.tree.column("source", width=440, minwidth=240)
        self.tree.tag_configure("group", foreground="#196a45")
        self.tree.tag_configure("single", foreground="#b27719")
        self.tree.tag_configure("rejected", foreground="#9b3535")
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=scroll.set)

        actions = ttk.LabelFrame(root, text="3  병합 실행", padding=10)
        actions.pack(fill="x", pady=(0, 10))
        action_row = ttk.Frame(actions)
        action_row.pack(fill="x")
        self.all_button = ttk.Button(action_row, text="전체 그룹 병합", style="Big.TButton", command=lambda: self._start_merge(False))
        self.all_button.pack(side="left", padx=(0, 7))
        self.selected_button = ttk.Button(action_row, text="선택 그룹만 병합", command=lambda: self._start_merge(True))
        self.selected_button.pack(side="left", padx=(0, 7))
        self.open_button = ttk.Button(action_row, text="결과 폴더 열기", command=self._open_results)
        self.open_button.pack(side="left")
        ttk.Label(action_row, textvariable=self.state_text, style="Body.TLabel").pack(side="right")
        self.progress = ttk.Progressbar(actions, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x", pady=(10, 0))
        self._set_enabled(False)

        log_box = ttk.LabelFrame(root, text="작업 기록", padding=8)
        log_box.pack(fill="both", pady=(0, 1))
        self.log = tk.Text(log_box, height=7, wrap="word", font=("Malgun Gothic", 9), background="#fbfdff", foreground="#273449")
        self.log.pack(fill="both")
        self.log.configure(state="disabled")
        self._write_log("안내: 원본 CSV를 변경하지 않고 시간별 시트가 있는 XLSX로 변환합니다. 기존 결과는 건너뜁니다.")

    def _set_enabled(self, can_merge: bool) -> None:
        state = "normal" if can_merge and not self.busy else "disabled"
        self.all_button.configure(state=state)
        self.selected_button.configure(state=state)
        self.scan_button.configure(state="disabled" if self.busy else "normal")
        self.pick_button.configure(state="disabled" if self.busy else "normal")

    def _write_log(self, msg: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _pick_folder(self) -> None:
        result = filedialog.askdirectory(title="원본 CSV 파일이 있는 폴더 선택", mustexist=True)
        if result:
            self.folder.set(result)
            self._scan()

    def _scan(self) -> None:
        location = self.folder.get().strip()
        if not location:
            messagebox.showinfo("폴더 선택", "먼저 원본 CSV 폴더를 선택해 주세요.")
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
            self.tree.insert("", "end", iid=gid, text=f"📗 {group.output_name}", values=(f"{len(group.inputs)}개 시트", "병합 가능"), tags=("group",), open=True)
            for j, item in enumerate(group.inputs):
                self.tree.insert(gid, "end", iid=f"{gid}i{j}", text=f"  └ {item.sheet_name}", values=(item.sheet_name, item.file_name))
        if report.singles:
            sid = self.tree.insert("", "end", text="⚠ 단독 파일 (병합 보류)", values=(f"{len(report.singles)}개", "1개인 그룹은 병합하지 않음"), tags=("single",), open=True)
            for group in report.singles:
                self.tree.insert(sid, "end", text=group.output_name, values=("대상 1개", group.inputs[0].file_name), tags=("single",))
        if report.rejected:
            xid = self.tree.insert("", "end", text="✕ 인식 제외 파일", values=(f"{len(report.rejected)}개", "병합하지 않음"), tags=("rejected",), open=True)
            for rejected in report.rejected:
                self.tree.insert(xid, "end", text=rejected.name, values=("제외", rejected.reason), tags=("rejected",))
        self.summary.set(summarize_scan(report))
        self._set_enabled(bool(report.groups))
        self.state_text.set("미리보기 준비 완료" if report.groups else "병합할 그룹 없음")
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
                        messagebox.showwarning("병합 종료", self.state_text.get() + "\n작업 기록을 확인하세요.")
                    else:
                        messagebox.showinfo("병합 종료", self.state_text.get() + "\n결과는 '병합결과' 폴더에 저장했습니다.")
        except queue.Empty:
            pass
        if self.winfo_exists():
            self.after(120, self._consume_events)

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

    def _close(self) -> None:
        if self.busy:
            messagebox.showwarning("병합 중", "병합 중에는 종료할 수 없습니다. 작업이 완료된 후 닫아 주세요.")
        else:
            self.destroy()


def run_app() -> None:
    app = MergeApp()
    app.mainloop()
