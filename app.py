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
from tkinter import filedialog, messagebox
from tkinter import font as tkfont

# ttk は ttkbootstrap 版を使う
import ttkbootstrap as tb
from ttkbootstrap import ttk as ttkb
from ttkbootstrap.constants import PRIMARY, INFO, SUCCESS, DANGER, SECONDARY, INVERSE
# ドラッグ＆ドロップ
from tkinterdnd2 import DND_FILES, TkinterDnD

# 48x48 PNG (Base64埋め込み)
APP_ICON_PNG48_B64 = """
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"""

# ====== QR検出ロジック ======
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


# ====== ユーティリティ ======
def _square_rect_from_points(pts: np.ndarray, zoom: float, margin: float = 4.0) -> fitz.Rect:
    xs = [x / zoom for x, _ in pts]
    ys = [y / zoom for _, y in pts]
    base = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
    w, h = base.width, base.height
    size = max(w, h)
    cx, cy = base.x0 + w / 2.0, base.y0 + h / 2.0
    sq = fitz.Rect(cx - size / 2.0, cy - size / 2.0, cx + size / 2.0, cy + size / 2.0)
    return fitz.Rect(sq.x0 - margin, sq.y0 - margin, sq.x1 + margin, sq.y1 + margin)


def _safe_add_text_annot(page, point, contents, icon="Comment"):
    try:
        return page.add_text_annot(point, contents, icon=icon)  # 新
    except AttributeError:
        return page.addTextAnnot(point, contents, icon=icon)    # 旧


def _safe_add_freetext_annot(page, rect, text, **kwargs):
    try:
        return page.add_freetext_annot(rect, text, **kwargs)    # 新
    except AttributeError:
        return page.addFreetextAnnot(rect, text, **kwargs)      # 旧


def _safe_insert_link(page, rect, uri):
    payload = {"kind": fitz.LINK_URI, "from": rect, "uri": uri}
    try:
        return page.insert_link(payload)  # 新
    except AttributeError:
        return page.insertLink(payload)   # 旧


def _rect_valid(r: fitz.Rect) -> bool:
    return (r is not None) and (r.width > 1.0) and (r.height > 1.0)


def _text_width(text: str, fontname: str, fontsize: float) -> float:
    try:
        return fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)  # 新
    except Exception:
        try:
            return fitz.getTextlength(text, fontname=fontname, fontsize=fontsize)  # 旧
        except Exception:
            return len(text) * fontsize * 0.6  # 概算


def _append_summary_pages(
    doc: fitz.Document,
    entries: list[tuple[int, str]],
    title: str = "QR Decode Summary",
    fontname: str = "helv",
):
    # ページサイズは先頭ページを踏襲、なければA4相当
    if len(doc) > 0:
        page_rect = doc[0].rect
    else:
        page_rect = fitz.Rect(0, 0, 595, 842)  # 約A4

    margin = 36  # 0.5 inch
    col_left = page_rect.x0 + margin
    col_right = page_rect.x1 - margin
    top = page_rect.y0 + margin
    bottom = page_rect.y1 - margin
    title_fs = 16
    body_fs = 11
    line_gap = body_fs * 1.35

    def new_page():
        return doc.new_page(width=page_rect.width, height=page_rect.height)

    def write_title(p: fitz.Page):
        p.insert_text(fitz.Point(col_left, top), title, fontsize=title_fs, fontname=fontname, color=(0, 0, 0))
        p.draw_line(
            fitz.Point(col_left, top + title_fs * 0.6),
            fitz.Point(col_right, top + title_fs * 0.6),
            color=(0, 0, 0),
            width=0.7,
        )
        return top + title_fs * 1.6

    def wrap_to_width(text: str, max_width: float) -> list[str]:
        if not text:
            return [""]
        lines, buf = [], ""
        for ch in text:
            if ch == "\n":
                lines.append(buf)
                buf = ""
                continue
            test = buf + ch
            if _text_width(test, fontname, body_fs) <= max_width:
                buf = test
            else:
                if buf == "":
                    lines.append(ch)
                    buf = ""
                else:
                    lines.append(buf)
                    buf = ch
        if buf:
            lines.append(buf)
        return lines

    page = new_page()
    y = write_title(page)

    info_line = f"Total: {len(entries)}"
    page.insert_text(fitz.Point(col_left, y), info_line, fontsize=body_fs, fontname=fontname, color=(0, 0, 0))
    y += line_gap

    max_width = col_right - col_left
    for idx, txt in entries:
        prefix = f"#{idx}: "
        first_line_budget = max_width - _text_width(prefix, fontname, body_fs)
        wrapped = []
        for i, seg in enumerate(txt.split("\n")):
            seg_lines = wrap_to_width(seg, max_width if i else max(first_line_budget, 24))
            if i == 0 and seg_lines:
                head = seg_lines[0]
                seg_lines[0] = prefix + head
            wrapped.extend(seg_lines)
        if not wrapped:
            wrapped = [prefix]

        for line in wrapped:
            if y + line_gap > bottom:
                page = new_page()
                y = write_title(page)
            page.insert_text(fitz.Point(col_left, y), line, fontsize=body_fs, fontname=fontname, color=(0, 0, 0))
            y += line_gap


def _is_encrypted(doc) -> bool:
    val = getattr(doc, "is_encrypted", None)
    if val is None:
        val = getattr(doc, "isEncrypted", None)
    needs = getattr(doc, "needs_pass", None)
    if needs is None:
        needs = getattr(doc, "needsPass", False)
    try:
        return bool(val) or bool(needs)
    except Exception:
        return False


# ====== 注釈付きPDFを書き出す ======
def export_annotated_pdf(input_bytes, detections_map, zoom_map):
    doc = fitz.open(stream=input_bytes, filetype="pdf")
    try:
        if _is_encrypted(doc):
            raise RuntimeError("暗号化されているPDFは対象外です。")
        global_idx = 1
        summary_entries: list[tuple[int, str]] = []

        for pidx in sorted(detections_map.keys()):
            page = doc.load_page(pidx)
            zoom = zoom_map.get(pidx, 3.0)
            dets = detections_map.get(pidx, [])
            for det in dets:
                pts = det["points"]
                txt = (det.get("text") or "").strip()

                rect = _square_rect_from_points(pts, zoom, margin=4.0)
                fill_annot = page.add_rect_annot(rect)
                fill_annot.set_border(width=0)
                fill_annot.set_colors(stroke=None, fill=(0, 1, 1))
                fill_annot.set_opacity(0.30)
                fill_annot.update()
                border_annot = page.add_rect_annot(rect)
                border_annot.set_border(width=1.5)
                border_annot.set_colors(stroke=(1, 0, 0), fill=None)
                border_annot.set_opacity(1.0)
                border_annot.update()

                ICON_EST = 20.0
                GAP = 6.0
                offset = ICON_EST + GAP
                bubble_pt = fitz.Point(rect.x0 - offset, rect.y0 - offset)
                pagebox = page.bound()
                bx = min(max(bubble_pt.x, pagebox.x0 + 2), pagebox.x1 - 2)
                by = min(max(bubble_pt.y, pagebox.y0 + 2), pagebox.y1 - 2)
                bubble_pt = fitz.Point(bx, by)
                contents_str = f"[#{global_idx}] {txt}"
                text_annot = _safe_add_text_annot(page, bubble_pt, contents_str, icon="Comment")
                try:
                    text_annot.set_info({"content": contents_str, "title": f"QR #{global_idx}", "subject": "QR decode"})
                except Exception:
                    pass
                text_annot.set_colors(stroke=(1, 0, 0), fill=None)
                text_annot.update()

                unit = max(12.0, min(rect.width, rect.height) / 4.0)
                label_text = f"#{global_idx}"
                fontsize = max(8.0, min(13.0, unit * 0.55))
                pad = max(2.0, fontsize * 0.35)
                text_w = _text_width(label_text, fontname="helv", fontsize=fontsize)
                label_w = max(unit, text_w + pad * 2.0)
                label_h = max(unit, fontsize * 1.35)
                label_rect = fitz.Rect(rect.x0, rect.y1 - label_h, rect.x0 + label_w, rect.y1)
                try:
                    ft = _safe_add_freetext_annot(
                        page, label_rect, label_text,
                        fontsize=fontsize, fontname="helv",
                        text_color=(1, 0, 0), fill_color=(1, 1, 1),
                        align=fitz.TEXT_ALIGN_LEFT, rotate=0,
                    )
                except TypeError:
                    ft = _safe_add_freetext_annot(page, label_rect, label_text, fontsize=fontsize, text_color=(1, 0, 0))
                try:
                    ft.set_border(width=0.8)
                    if hasattr(ft, "set_opacity"):
                        ft.set_opacity(0.90)
                    ft.set_info({"title": f"QR #{global_idx}", "subject": "QR label"})
                    ft.update()
                except Exception:
                    pass

                is_url = txt.lower().startswith("http://") or txt.lower().startswith("https://")
                if is_url:
                    link_top = fitz.Rect(rect.x0, rect.y0, rect.x1, label_rect.y0)
                    link_right = fitz.Rect(label_rect.x1, label_rect.y0, rect.x1, rect.y1)
                    for lr in (link_top, link_right):
                        if _rect_valid(lr):
                            _safe_insert_link(page, lr, txt)

                summary_entries.append((global_idx, txt if txt else ""))
                global_idx += 1

        _append_summary_pages(doc, summary_entries, title="QR Decode Summary")
        out = io.BytesIO()
        doc.save(out, deflate=True)
        return out.getvalue()
    finally:
        doc.close()


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

        # 参照保持（GC対策）
        self._icon_img48 = tk.PhotoImage(data=APP_ICON_PNG48_B64)
        # 以後に作られる Toplevel にも同じアイコンを適用
        self.iconphoto(True, self._icon_img48)

        # ---- 基本設定
        self.title("PDF内QRコードに注釈追加")
        self.geometry("700x640")

        # ---- ttkbootstrap テーマ
        self.style = tb.Style(theme="minty")

        # ---- フォントを Noto Sans 系に統一（自動フォールバック）
        family = self._choose_font_family()
        self._apply_global_fonts(family)
        # Canvas用フォント（名前付きで保持）
        self._font_ui = tkfont.Font(family=family, size=11)
        self._font_small = tkfont.Font(family=family, size=9)

        # ---- 入力状態
        self.pdf_path = tk.StringVar()
        self.zoom = tk.DoubleVar(value=3.0)
        self.page_sel = tk.StringVar(value="all")
        self.auto_run_on_drop = tk.BooleanVar(value=True)  # ドロップで自動開始（既定ON）

        # ワーカー系
        self._worker: threading.Thread | None = None
        self._stop_flag = False
        self.annotated_bytes: bytes | None = None

        # 単色パレット（ドロップエリア用）
        self._pal = {"bg": "#FFFFFF", "border": "#D0D5DD", "text": "#111827", "muted": "#6B7280", "accent": "#2D7FF9"}

        # UI構築
        self._build_ui()
        self._bind_shortcuts()

    # ===== フォント適用 =====
    def _choose_font_family(self) -> str:
        # 優先順位で選択（存在しなければ次へ）
        preferred = ["Noto Sans JP", "Noto Sans", "Yu Gothic UI", "Segoe UI", "Arial"]
        available = set(tkfont.families())
        for f in preferred:
            if f in available:
                return f
        return "TkDefaultFont"

    def _apply_global_fonts(self, family: str):
        # Tk 名前付きフォントを書き換えると既存/今後のウィジェットに反映
        for name, size, weight in [
            ("TkDefaultFont", 11, "normal"),
            ("TkTextFont", 11, "normal"),
            ("TkFixedFont", 11, "normal"),
            ("TkHeadingFont", 13, "bold"),
            ("TkMenuFont", 11, "normal"),
            ("TkTooltipFont", 10, "normal"),
        ]:
            try:
                f = tkfont.nametofont(name)
                f.configure(family=family, size=size, weight=weight)
            except Exception:
                pass

    # ===== UI =====
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        # ===== 入力カード =====
        card_in = ttkb.Labelframe(self, text="入力", bootstyle=SECONDARY)
        card_in.pack(fill="x", **pad)

        # PDF選択
        row = 0
        ttkb.Label(card_in, text="PDFファイル").grid(row=row, column=0, sticky="e", padx=8, pady=8)
        self.ent_pdf = ttkb.Entry(card_in, textvariable=self.pdf_path, width=60)
        self.ent_pdf.grid(row=row, column=1, sticky="we", padx=8, pady=8)
        ttkb.Button(card_in, text="参照…", command=self.select_pdf, bootstyle=PRIMARY).grid(
            row=row, column=2, sticky="w", padx=8, pady=8
        )
        card_in.columnconfigure(1, weight=1)

        # Entryにドロップ対応
        self.ent_pdf.drop_target_register(DND_FILES)
        self.ent_pdf.dnd_bind("<<Drop>>", self._on_drop)

        # レンダリング倍率
        row += 1
        ttkb.Label(card_in, text="レンダリング倍率").grid(row=row, column=0, sticky="e", padx=8, pady=4)
        zfrm = ttkb.Frame(card_in)
        zfrm.grid(row=row, column=1, sticky="we", padx=8, pady=4)
        self.zoom_label = ttkb.Label(zfrm, text=f"{self.zoom.get():.2f}x", bootstyle=SECONDARY)
        self.zoom_label.pack(side="right")
        self.zoom_scale = ttkb.Scale(
            zfrm, from_=2.0, to=5.0, orient="horizontal",
            command=self._on_zoom_change, value=self.zoom.get()
        )
        self.zoom_scale.pack(fill="x", side="left", expand=True, padx=(0, 8))
        ttkb.Label(card_in, text="").grid(row=row, column=2, sticky="w")

        # 解析ページ範囲
        row += 1
        ttkb.Label(card_in, text="解析ページ範囲").grid(row=row, column=0, sticky="e", padx=8, pady=4)
        ttkb.Entry(card_in, textvariable=self.page_sel).grid(row=row, column=1, sticky="we", padx=8, pady=4)
        ttkb.Label(card_in, text="例: all / 1-3,5,10-12", bootstyle=SECONDARY).grid(
            row=row, column=2, sticky="w", padx=8, pady=4
        )

        # 自動開始トグル
        row += 1
        ttkb.Checkbutton(card_in, text="ドロップで自動解析する", variable=self.auto_run_on_drop).grid(
            row=row, column=1, sticky="w", padx=8, pady=(0, 8)
        )

        # ===== ドロップエリア（Canvas演出） =====
        drop_card = ttkb.Labelframe(self, text="ドラッグ＆ドロップ", bootstyle=SECONDARY)
        drop_card.pack(fill="x", **pad)

        self.drop_canvas = tk.Canvas(drop_card, height=120, bd=0, highlightthickness=0, cursor="hand2")
        self.drop_canvas.pack(fill="x", padx=10, pady=10)

        # DnD登録
        self.drop_canvas.drop_target_register(DND_FILES)
        self.drop_canvas.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_canvas.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.drop_canvas.dnd_bind("<<DragLeave>>", self._on_drag_leave)

        # ウィンドウ全体も受け付け（任意）
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._on_drop)
        self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
        self.dnd_bind("<<DragLeave>>", self._on_drag_leave)

        self._draw_drop_area(hover=False)
        self.drop_canvas.bind("<Configure>", lambda e: self._draw_drop_area(hover=False))

        # ===== 実行ボタン群（角丸が映える配色）=====
        btnfrm = ttkb.Frame(self)
        btnfrm.pack(fill="x", **pad)
        self.btn_run  = ttkb.Button(btnfrm, text="解析開始 ▶️", bootstyle=PRIMARY,  command=self.start_process)
        self.btn_stop = ttkb.Button(btnfrm, text="停止 ⏹",     bootstyle=DANGER,   command=self.stop_process, state="disabled")
        self.btn_save = ttkb.Button(btnfrm, text="保存… 💾",    bootstyle=SUCCESS,  command=self.save_output, state="disabled")
        self.btn_run.pack(side="left")
        self.btn_stop.pack(side="left", padx=6)
        self.btn_save.pack(side="left", padx=6)

        # ===== 進捗・ログ =====
        pfrm = ttkb.Labelframe(self, text="進捗", bootstyle=SECONDARY)
        pfrm.pack(fill="x", **pad)
        self.progress = ttkb.Progressbar(pfrm, maximum=100, mode="determinate", bootstyle=f"{INFO}-striped")
        self.progress.pack(fill="x", padx=10, pady=10)
        self.progress_text = ttkb.Label(pfrm, text="0%", bootstyle=SECONDARY)
        self.progress_text.place(in_=self.progress, relx=0.5, rely=0.5, anchor="center")

        self.status = ttkb.Label(self, text="待機中", bootstyle=SECONDARY)
        self.status.pack(anchor="w", padx=12)

        log_card = ttkb.Labelframe(self, text="ログ", bootstyle=SECONDARY)
        log_card.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(log_card, height=12, relief="flat", wrap="word")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    # ===== ドロップエリア描画 =====
    def _draw_drop_area(self, hover: bool):
        p = self._pal
        c = self.drop_canvas
        c.delete("all")
        bg = p["bg"]
        fg = p["accent"] if hover else p["border"]
        c.configure(bg=bg)
        w = c.winfo_width() or c.winfo_reqwidth()
        h = c.winfo_height() or c.winfo_reqheight()
        pad = 10
        c.create_rectangle(pad, pad, w - pad, h - pad, dash=(6, 4), outline=fg, width=2, fill=bg)
        c.create_text(w / 2, h / 2 - 8, text="ここに PDF をドラッグ＆ドロップ ⤵️",
                      fill=p["text"], font=self._font_ui)
        c.create_text(w / 2, h / 2 + 16, text="または「参照…」をクリック",
                      fill=p["muted"], font=self._font_small)

    def _on_drag_enter(self, event):
        self._draw_drop_area(hover=True)

    def _on_drag_leave(self, event):
        self._draw_drop_area(hover=False)

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
            picked = pdfs[0]
            self.pdf_path.set(picked)
            self.log_write(f"ドロップ: {picked}\n")
            if self._worker and self._worker.is_alive():
                self.log_write("現在解析中のため、自動開始はスキップしました。\n")
                return
            if self.auto_run_on_drop.get():
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
        if self._worker and self._worker.is_alive():
            return
        path = self.pdf_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("エラー", "PDFファイルを選択してください。")
            return

        try:
            test_doc = fitz.open(path)
            try:
                if _is_encrypted(test_doc):
                    self.log_write("暗号化されているPDFは対象外のためスキップします。\n")
                    messagebox.showinfo("対象外", "暗号化されているPDFは解析対象外です。")
                    return
            finally:
                test_doc.close()
        except Exception:
            pass

        self._stop_flag = False
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.btn_save.config(state="disabled")
        self.progress.config(value=0)
        try:
            self.progress_text.config(text="0%")
        except Exception:
            pass
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
            self.btn_run.config(state="normal")
            self.btn_stop.config(state="disabled")
            if self.annotated_bytes:
                self.btn_save.config(state="normal")
                self.status.config(text="完了")
                self.after(150, self.save_output)
            else:
                self.status.config(text="中断/失敗")

    def _process_worker(self):
        try:
            pdf_path = self.pdf_path.get()
            zoom = float(self.zoom.get())
            page_sel = self.page_sel.get().strip()
            doc_in = fitz.open(pdf_path)
            try:
                if _is_encrypted(doc_in):
                    raise RuntimeError("暗号化されているPDFは対象外です。")
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
                self.annotated_bytes = export_annotated_pdf(file_bytes, detections_map, zoom_map)
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
            try:
                self.progress_text.config(text=f"{val:.0f}%")
            except Exception:
                pass
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

    def _bind_shortcuts(self):
        self.bind_all("<Control-o>", lambda e: self.select_pdf())
        self.bind_all("<Control-r>", lambda e: self.start_process())
        self.bind_all("<Control-s>", lambda e: self.save_output())
        self.bind_all("<Escape>",    lambda e: self.stop_process())


if __name__ == "__main__":
    app = QRPdfAnnotatorApp()
    app.mainloop()
