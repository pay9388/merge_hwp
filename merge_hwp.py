#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
한글(HWP/HWPX) 병합기 — Slim 버전 (Windows + 한글 설치 필요)
- 첫 파일을 기반 문서로 열고 나머지를 InsertFile로 뒤에 삽입(공백 첫 페이지 방지)
- 서식 유지 옵션(쪽/글자/문단/스타일)
- 병합 후 PDF 저장 옵션
- 필수 패키지: pip install pyhwpx pywin32
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

SUPPORTED_EXTS = (".hwp", ".hwpx")

# ----------------------- HWP 래핑(간결화) -----------------------
try:
    from pyhwpx import Hwp
except Exception as e:
    raise SystemExit(
        "[pyhwpx] 불러오기 실패: Windows + 한글(HWP) 설치와 'pip install pyhwpx pywin32'가 필요합니다.\n"
        f"에러: {e}"
    )

class HwpSession:
    """with HwpSession() as hwp: 형태로 쓰면 자동 quit() 보장"""
    def __enter__(self):
        self.hwp = Hwp()
        return self.hwp
    def __exit__(self, exc_type, exc, tb):
        try:
            self.hwp.quit()
        except Exception:
            pass

def insert_file(hwp, path, keep):
    """InsertFile 실행(서식 유지 옵션 dict 사용)"""
    hwp.HAction.Run("MoveDocEnd")  # 항상 문서 끝에 삽입
    p = hwp.HParameterSet.HInsertFile
    hwp.HAction.GetDefault("InsertFile", p.HSet)
    p.filename = os.path.abspath(path)
    p.KeepSection   = 1 if keep["section"]   else 0
    p.KeepCharshape = 1 if keep["char"]      else 0
    p.KeepParashape = 1 if keep["para"]      else 0
    p.KeepStyle     = 1 if keep["style"]     else 0
    if not hwp.HAction.Execute("InsertFile", p.HSet):
        raise RuntimeError(f"InsertFile 실패: {path}")

def export_pdf(hwp, pdf_path):
    """현재 문서를 PDF로 내보내기"""
    p = hwp.HParameterSet.HFileOpenSave
    hwp.HAction.GetDefault("FileSaveAs_S", p.HSet)
    p.filename = os.path.abspath(pdf_path)
    p.Format   = "PDF"
    if not hwp.HAction.Execute("FileSaveAs_S", p.HSet):
        raise RuntimeError(f"PDF 저장 실패: {pdf_path}")

def merge_to(output_path, paths, keep, save_pdf):
    """핵심 병합 함수(간결): 첫 파일 open → 나머지 InsertFile → 저장 → (선택)PDF"""
    if not paths:
        raise ValueError("병합할 파일이 없습니다.")
    first, rest = os.path.abspath(paths[0]), [os.path.abspath(p) for p in paths[1:]]
    if not os.path.exists(first):
        raise FileNotFoundError(first)

    with HwpSession() as hwp:
        hwp.open(first)                          # 공백 첫 페이지 방지
        for p in rest:
            if not os.path.exists(p):
                raise FileNotFoundError(p)
            insert_file(hwp, p, keep)            # 서식 유지 옵션 반영
        hwp.save_as(os.path.abspath(output_path))
        if save_pdf:
            base, _ = os.path.splitext(output_path)
            export_pdf(hwp, base + ".pdf")

# ----------------------- GUI(필수 요소만) -----------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HWP/HWPX 병합기")
        self.geometry("680x520")
        self.paths = []
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        # 상단: 파일 추가/제거/정렬
        top = ttk.Frame(self); top.pack(fill="x", **pad)
        ttk.Button(top, text="파일 추가(다중)", command=self.add_files).pack(side="left")
        ttk.Button(top, text="선택 제거", command=self.remove_sel).pack(side="left", padx=6)
        ttk.Button(top, text="모두 지우기", command=self.clear_all).pack(side="left")

        sort = ttk.Frame(self); sort.pack(fill="x", **pad)
        ttk.Label(sort, text="정렬:").pack(side="left")
        self.sort_var = tk.StringVar(value="name")
        ttk.Radiobutton(sort, text="파일명", value="name", variable=self.sort_var, command=self.sort_paths).pack(side="left")
        ttk.Radiobutton(sort, text="수동(위/아래)", value="manual", variable=self.sort_var).pack(side="left", padx=10)

        # 리스트 + 위/아래
        center = ttk.Frame(self); center.pack(fill="both", expand=True, **pad)
        self.listbox = tk.Listbox(center, selectmode="extended"); self.listbox.pack(side="left", fill="both", expand=True)
        btns = ttk.Frame(center); btns.pack(side="left", fill="y", padx=(8,0))
        ttk.Button(btns, text="▲ 위", command=lambda: self.move(-1)).pack(fill="x", pady=3)
        ttk.Button(btns, text="▼ 아래", command=lambda: self.move(+1)).pack(fill="x", pady=3)

        # 옵션(서식 유지 + PDF)
        opt = ttk.LabelFrame(self, text="서식 유지 옵션"); opt.pack(fill="x", **pad)
        self.keep_section   = tk.BooleanVar(value=True)
        self.keep_char      = tk.BooleanVar(value=True)
        self.keep_para      = tk.BooleanVar(value=True)
        self.keep_style     = tk.BooleanVar(value=True)
        for text,var in [("쪽 구역",self.keep_section),("글자 모양",self.keep_char),
                         ("문단 모양",self.keep_para),("스타일",self.keep_style)]:
            ttk.Checkbutton(opt, text=text, variable=var).pack(side="left", padx=6)

        pdf = ttk.Frame(self); pdf.pack(fill="x", **pad)
        self.save_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(pdf, text="병합 후 PDF로 저장", variable=self.save_pdf).pack(side="left")

        # 저장 경로
        out = ttk.Frame(self); out.pack(fill="x", **pad)
        ttk.Label(out, text="저장 파일:").pack(side="left")
        self.output_entry = ttk.Entry(out); self.output_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(out, text="선택…", command=self.pick_output).pack(side="left", padx=6)

        # 실행
        bottom = ttk.Frame(self); bottom.pack(fill="x", **pad)
        self.status = tk.StringVar(value="준비")
        ttk.Label(bottom, textvariable=self.status).pack(side="left")
        ttk.Button(bottom, text="병합 실행", command=self.on_merge).pack(side="right")


        # 👇 제작자 표시
        footer = tk.Label(
            self,
            text="마산구암고등학교 미래교육지원부 (2025)",
            anchor="center",
            fg="white",
            bg="#2c3e50",          # 네이비
            font=("맑은 고딕", 9, "bold")
        )
        footer.pack(side="bottom", fill="x")

    # ---------- 이벤트 ----------
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="병합할 파일 선택", filetypes=[("HWP/HWPX", "*.hwp *.hwpx"), ("HWP", "*.hwp"), ("HWPX", "*.hwpx")]
        )
        for f in files:
            if f.lower().endswith(SUPPORTED_EXTS) and f not in self.paths:
                self.paths.append(f)
        if self.sort_var.get() == "name": self.sort_paths()
        else: self.refresh()

    def remove_sel(self):
        for i in reversed(self.listbox.curselection()):
            del self.paths[i]
        self.refresh()

    def clear_all(self):
        self.paths.clear(); self.refresh()

    def sort_paths(self):
        self.paths.sort(key=lambda p: (os.path.basename(p).lower(), p.lower()))
        self.refresh()

    def move(self, d):
        if self.sort_var.get() != "manual":
            messagebox.showinfo("안내", "수동 정렬 모드에서만 이동 가능합니다."); return
        sel = list(self.listbox.curselection())
        if not sel: return
        a = self.paths
        rng = range(len(a)-1, -1, -1) if d>0 else range(len(a))
        sset = set(sel)
        for i in rng:
            if i in sset:
                j = i + d
                if 0 <= j < len(a) and j not in sset:
                    a[i], a[j] = a[j], a[i]
        self.refresh()
        # 선택 유지
        self.listbox.selection_clear(0,"end")
        for i in [min(max(i+d,0),len(a)-1) for i in sel]:
            self.listbox.selection_set(i)

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="병합 결과 저장", defaultextension=".hwp", filetypes=[("HWP", "*.hwp"), ("HWPX", "*.hwpx")]
        )
        if path:
            self.output_entry.delete(0,"end"); self.output_entry.insert(0,path)

    def on_merge(self):
        try:
            if not self.paths: 
                messagebox.showwarning("경고","병합할 파일을 추가하세요."); return
            out = (self.output_entry.get() or "").strip()
            if not out or not out.lower().endswith(SUPPORTED_EXTS):
                messagebox.showwarning("경고",".hwp 또는 .hwpx 확장자로 저장하세요."); return
            self.status.set("병합 중…"); self.update_idletasks()

            keep = {"section":self.keep_section.get(), "char":self.keep_char.get(),
                    "para":self.keep_para.get(), "style":self.keep_style.get()}
            merge_to(out, self.paths, keep, self.save_pdf.get())

            self.status.set("완료")
            msg = f"병합 완료\n{out}"
            if self.save_pdf.get(): msg += f"\n{os.path.splitext(out)[0]}.pdf 저장됨"
            messagebox.showinfo("완료", msg)
        except Exception as e:
            self.status.set("오류")
            messagebox.showerror("오류", f"{e}")

    def refresh(self):
        self.listbox.delete(0,"end")
        for p in self.paths: self.listbox.insert("end", os.path.basename(p))

# ----------------------- Entry -----------------------
if __name__ == "__main__":
    if os.name != "nt":
        raise SystemExit("이 프로그램은 Windows에서만 동작합니다.")
    App().mainloop()
