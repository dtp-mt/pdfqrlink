# -*- coding: utf-8 -*-
import io
import os
import threading
import traceback
import numpy as np
import fitz  # PyMuPDF
from PIL import Image
import cv2
import zxingcpp

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

# ドラッグ＆ドロップ
from tkinterdnd2 import DND_FILES, TkinterDnD


# ====== 検出ロジック ======
def detect_and_decode_qr_zxing(pil_img: Image.Image) -> list:
    img_rgb = np.array(pil_img.convert("RGB"))
    img_bgr = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2BGR)
    img_bgr = np.ascontiguousarray(img_bgr)
    barcodes = zxingcpp.read_barcodes(img_bgr)

    results = []
    for bc in barcodes:
        if bc.format != zxingcpp.BarcodeFormat.QRCode:
            continue
        pos = getattr(bc, "position", None)
        if pos is None:
            continue
        pts = np.array([
            (float(pos.top_left.x), float(pos.top_left.y)),
            (float(pos.top_right.x), float(pos.top_right.y)),
            (float(pos.bottom_right.x), float(pos.bottom_right.y)),
            (float(pos.bottom_left.x), float(pos.bottom_left.y)),
        ], dtype=np.float32)
        results.append({"text": getattr(bc, "text", ""), "points": pts})
    return results


def export_annotated_pdf(input_bytes: bytes, detections_map: dict, zoom_map: dict, password: str | None = None) -> bytes:
    doc = fitz.open(stream=input_bytes, filetype="pdf")
    if doc.needs_pass and password:
        if not doc.authenticate(password):
            doc.close()
            raise RuntimeError("PDFパスワードが正しくありません。")

    for pidx, dets in detections_map.items():
        page = doc.load_page(pidx)
        zoom = zoom_map.get(pidx, 3.0)
        for det in dets:
            pts = det["points"]
            xs = [x / zoom for x, _ in pts]
            ys = [y / zoom for _, y in pts]
            rect = fitz.Rect(min(xs), min(ys), max(xs), max(ys))

            margin = 4.0
            rect = fitz.Rect(rect.x0 - margin, rect.y0 - margin, rect.x1 + margin, rect.y1 + margin)

            # 半透明シアンの塗り
            fill_annot = page.add_rect_annot(rect)
            fill_annot.set_border(width=0)
            fill_annot.set_colors(stroke=None, fill=(0, 1, 1))
            fill_annot.set_opacity(0.3)
            fill_annot.update()

            # 赤枠
            border_annot = page.add_rect_annot(rect)
            border_annot.set_border(width=1.5)
            border_annot.set_colors(stroke=(1, 0, 0), fill=None)
            border_annot.set_opacity(1.0)
            border_annot.update()

            # URLはリンク化
            txt = det["text"] or ""
            if txt.startswith("http://") or txt.startswith("https://"):
                page.insert_link({"kind": fitz.LINK_URI, "from": rect, "uri": txt})

    out = io.BytesIO()
    doc.save(out, deflate=True)
    doc.close()
    return out.getvalue()


def parse_pages(sel: str, total: int) -> list[int]:
    if not sel or sel.strip().lower() == "all":
        return list(range(total))
    pages = set()
    for token in sel.split(","):
        token = token.strip()
        if not token:
            continue
        if "-" in token:
            a, b = token.split("-", 1)
            a = int(a)
            b = int(b)
            pages.update(range(a - 1, b))
        else:
            pages.add(int(token) - 1)
    return sorted(p for p in pages if 0 <= p < total)


# ====== GUI アプリ（DnD + 自動開始） ======
class QRPdfAnnotatorApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("QR PDF Annotator (GUI版)")
        self.geometry("780x580")

        # 入力状態
        self.pdf_path = tk.StringVar()
        self.password = tk.StringVar()
        self.zoom = tk.DoubleVar(value=3.0)
        self.page_sel = tk.StringVar(value="all")
        self.auto_run_on_drop = tk.BooleanVar(value=True)  # ★ ドロップで自動開始（既定ON）

        self._build_ui()

        # 処理用
        self._worker: threading.Thread | None = None
        self._stop_flag = False

        # 出力
        self.annotated_bytes: bytes | None = None

    # ===== UI =====
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="x", **pad)

        # PDF選択
        ttk.Label(frm, text="PDFファイル").grid(row=0, column=0, sticky="e")
        self.ent_pdf = ttk.Entry(frm, textvariable=self.pdf_path, width=60)
        self.ent_pdf.grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(frm, text="参照…", command=self.select_pdf).grid(row=0, column=2, sticky="w")
        frm.columnconfigure(1, weight=1)

        # Entryにドロップ対応
        self.ent_pdf.drop_target_register(DND_FILES)
        self.ent_pdf.dnd_bind('<<Drop>>', self._on_drop)

        # パスワード
        ttk.Label(frm, text="パスワード（必要な場合）").grid(row=1, column=0, sticky="e")
        ttk.Entry(frm, textvariable=self.password, show="*").grid(row=1, column=1, sticky="we", **pad)

        # ズーム（スライダー）
        zfrm = ttk.Frame(frm)
        zfrm.grid(row=2, column=1, sticky="we", **pad)
        ttk.Label(frm, text="レンダリング倍率").grid(row=2, column=0, sticky="e")
        self.zoom_label = ttk.Label(zfrm, text=f"{self.zoom.get():.2f}x")
        self.zoom_label.pack(side="right")
        self.zoom_scale = ttk.Scale(zfrm, from_=2.0, to=5.0, orient="horizontal",
                                    command=self._on_zoom_change, value=self.zoom.get())
        self.zoom_scale.pack(fill="x", side="left", expand=True)

        # ページ範囲
        ttk.Label(frm, text="解析ページ範囲").grid(row=3, column=0, sticky="e")
        ttk.Entry(frm, textvariable=self.page_sel).grid(row=3, column=1, sticky="we", **pad)
        ttk.Label(frm, text="例: all / 1-3,5,10-12").grid(row=3, column=2, sticky="w")

        # 自動開始トグル
        chk = ttk.Checkbutton(self, text="ドロップで自動解析する", variable=self.auto_run_on_drop)
        chk.pack(anchor="w", padx=12, pady=(0, 4))

        # 明確なドロップエリア
        self.drop_box = ttk.Label(self, text="ここに PDF をドラッグ＆ドロップ 📄👇", anchor="center",
                                  relief="groove", padding=18)
        self.drop_box.pack(fill="x", padx=10, pady=(0, 8))
        self.drop_box.drop_target_register(DND_FILES)
        self.drop_box.dnd_bind('<<Drop>>', self._on_drop)

        # ウィンドウ全体も受け付け（任意）
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self._on_drop)

        # 実行・停止・保存
        btnfrm = ttk.Frame(self)
        btnfrm.pack(fill="x", **pad)
        self.btn_run = ttk.Button(btnfrm, text="解析開始 ▶️", command=self.start_process)
        self.btn_run.pack(side="left")
        self.btn_stop = ttk.Button(btnfrm, text="停止 ⏹", command=self.stop_process, state="disabled")
        self.btn_stop.pack(side="left", padx=5)
        self.btn_save = ttk.Button(btnfrm, text="保存… 💾", command=self.save_output, state="disabled")
        self.btn_save.pack(side="left", padx=5)

        # 進捗・ログ
        pfrm = ttk.Frame(self)
        pfrm.pack(fill="x", **pad)
        self.progress = ttk.Progressbar(pfrm, maximum=100, mode="determinate")
        self.progress.pack(fill="x", expand=True)
        self.status = ttk.Label(self, text="待機中")
        self.status.pack(anchor="w", padx=12)

        ttk.Label(self, text="ログ").pack(anchor="w", padx=12)
        self.log = tk.Text(self, height=12)
        self.log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # ===== イベント =====
    def _on_drop(self, event):
        try:
            paths = self.tk.splitlist(event.data)
            if not paths:
                return
            pdfs = [p.strip("{}") for p in paths if p.strip("{}").lower().endswith(".pdf")]
            if not pdfs:
                messagebox.showwarning("注意", "PDFファイルをドロップしてください。")
                return

            # 最初のPDFを採用
            picked = pdfs[0]
            self.pdf_path.set(picked)
            self.log_write(f"ドロップ: {picked}\n")

            # 解析中なら自動開始はしない（完了後に手動で開始 or もう一度ドロップ）
            if self._worker and self._worker.is_alive():
                self.log_write("現在解析中のため、自動開始はスキップしました。\n")
                return

            # 自動開始（有効時）
            if self.auto_run_on_drop.get():
                # 少し遅延させてUI更新を反映
                self.after(150, self.start_process)

        except Exception as e:
            messagebox.showerror("エラー", f"ドロップ処理に失敗しました: {e}")

    def _on_zoom_change(self, val):
        try:
            v = float(val)
        except ValueError:
            v = 3.0
        self.zoom.set(v)
        self.zoom_label.config(text=f"{v:.2f}x")

    def select_pdf(self):
        path = filedialog.askopenfilename(
            title="PDFを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_path.set(path)

    # ===== 実行系 =====
    def start_process(self):
        # すでに実行中なら何もしない
        if self._worker and self._worker.is_alive():
            return

        path = self.pdf_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("エラー", "PDFファイルを選択してください。")
            return

        # ★ 自動開始時でもパスワード事前確認（必要ならダイアログ）
        pwd = self.password.get().strip()
        try:
            test_doc = fitz.open(path)
            try:
                if test_doc.needs_pass and not pwd:
                    ask = simpledialog.askstring("パスワード",
                                                 "このPDFにはパスワードが必要です。入力してください：",
                                                 show="*",
                                                 parent=self)
                    if not ask:
                        self.log_write("開始をキャンセルしました（パスワード未入力）。\n")
                        return
                    self.password.set(ask)
            finally:
                test_doc.close()
        except Exception:
            # 壊れたPDFなど
            pass

        # UI 切り替え
        self._stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_save.config(state="disabled")
        self.progress.config(value=0)
        self.status.config(text="解析準備中…")
        self.log_delete()
        self.log_write(f"入力PDF: {path}\n")

        self._worker = threading.Thread(target=self._process_worker, daemon=True)
        self._worker.start()
        self.after(200, self._poll_worker)

    def stop_process(self):
        self._stop_flag = True
        self.log_write("停止要求を受け付けました…\n")

    def _poll_worker(self):
        if self._worker and self._worker.is_alive():
            self.after(200, self._poll_worker)
        else:
            # 終了時のUI復帰
            self.btn_run.config(state="normal")
            self.btn_stop.config(state="disabled")
            if self.annotated_bytes:
                self.btn_save.config(state="normal")
                self.status.config(text="完了")
            else:
                self.status.config(text="中断/失敗")

    def _process_worker(self):
        try:
            pdf_path = self.pdf_path.get()
            pwd = self.password.get().strip() or None
            zoom = float(self.zoom.get())
            page_sel = self.page_sel.get().strip()

            doc_in = fitz.open(pdf_path)
            try:
                if doc_in.needs_pass:
                    if not pwd:
                        raise RuntimeError("このPDFにはパスワードが必要です。")
                    if not doc_in.authenticate(pwd):
                        raise RuntimeError("パスワードが正しくありません。")

                total_pages = len(doc_in)
                target_pages = parse_pages(page_sel, total_pages)
                if not target_pages:
                    raise RuntimeError("解析対象ページが空です。指定を見直してください。")

                self.log_write(f"ページ数: {total_pages} / 解析対象: {', '.join(str(p+1) for p in target_pages)}\n")

                detections_map: dict[int, list] = {}
                zoom_map: dict[int, float] = {}

                for i, pidx in enumerate(target_pages, start=1):
                    if self._stop_flag:
                        self.log_write("ユーザーにより停止されました。\n")
                        self.annotated_bytes = None
                        return
                    page = doc_in.load_page(pidx)
                    zoom_map[pidx] = zoom
                    mat = fitz.Matrix(zoom, zoom)
                    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csRGB)

                    pil_img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                    detections = detect_and_decode_qr_zxing(pil_img)
                    detections_map[pidx] = detections

                    self.log_write(f"Page {pidx+1}: QR {len(detections)}件\n")
                    self._set_progress(i / len(target_pages) * 100.0)
                    self._set_status(f"解析中… ({i}/{len(target_pages)})")

                with open(pdf_path, "rb") as f:
                    file_bytes = f.read()
                self.annotated_bytes = export_annotated_pdf(file_bytes, detections_map, zoom_map, password=pwd)
                self.log_write("注釈PDFの生成が完了しました。\n")
            finally:
                doc_in.close()

        except Exception as e:
            self.annotated_bytes = None
            self.log_write("エラー: " + str(e) + "\n")
            self.log_write(traceback.format_exc() + "\n")
            messagebox.showerror("エラー", str(e))

    # ===== 小物 =====
    def _set_progress(self, val):
        def _inner():
            self.progress.config(value=val)
        self.after(0, _inner)

    def _set_status(self, text):
        def _inner():
            self.status.config(text=text)
        self.after(0, _inner)

    def save_output(self):
        if not self.annotated_bytes:
            messagebox.showwarning("警告", "出力データがありません。先に解析してください。")
            return
        in_name = os.path.splitext(os.path.basename(self.pdf_path.get()))[0]
        out_path = filedialog.asksaveasfilename(
            title="注釈付きPDFを保存",
            defaultextension=".pdf",
            initialfile=f"{in_name}_annotated.pdf",
            filetypes=[("PDF files", "*.pdf")]
        )
        if out_path:
            try:
                with open(out_path, "wb") as f:
                    f.write(self.annotated_bytes)
                messagebox.showinfo("完了", "保存しました。")
            except Exception as e:
                messagebox.showerror("エラー", f"保存に失敗しました: {e}")

    def log_write(self, text: str):
        def _inner():
            self.log.insert("end", text)
            self.log.see("end")
        self.after(0, _inner)

    def log_delete(self):
        self.log.delete("1.0", "end")


if __name__ == "__main__":
    app = QRPdfAnnotatorApp()
    app.mainloop()
