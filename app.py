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
            (float(pos.top_left.x),     float(pos.top_left.y)),
            (float(pos.top_right.x),    float(pos.top_right.y)),
            (float(pos.bottom_right.x), float(pos.bottom_right.y)),
            (float(pos.bottom_left.x),  float(pos.bottom_left.y)),
        ], dtype=np.float32)
        results.append({"text": getattr(bc, "text", ""), "points": pts})
    return results


# ====== ユーティリティ ======
def _square_rect_from_points(pts: np.ndarray, zoom: float, margin: float = 4.0) -> fitz.Rect:
    """ZXingの4点から矩形を作り、中心を保ったまま正方形化して少しマージンを取る"""
    xs = [x / zoom for x, _ in pts]
    ys = [y / zoom for _, y in pts]
    base = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
    w, h = base.width, base.height
    size = max(w, h)
    cx, cy = base.x0 + w / 2.0, base.y0 + h / 2.0
    sq = fitz.Rect(cx - size / 2.0, cy - size / 2.0, cx + size / 2.0, cy + size / 2.0)
    # 追加マージン
    return fitz.Rect(sq.x0 - margin, sq.y0 - margin, sq.x1 + margin, sq.y1 + margin)


def _safe_add_text_annot(page, point, contents, icon="Comment"):
    """Text注釈（付箋）を新旧API名に対応して追加"""
    try:
        return page.add_text_annot(point, contents, icon=icon)  # 新
    except AttributeError:
        return page.addTextAnnot(point, contents, icon=icon)    # 旧


def _safe_add_freetext_annot(page, rect, text, **kwargs):
    """FreeText注釈（テキストボックス）を新旧API名に対応して追加"""
    try:
        return page.add_freetext_annot(rect, text, **kwargs)    # 新
    except AttributeError:
        return page.addFreetextAnnot(rect, text, **kwargs)      # 旧


def _safe_insert_link(page, rect, uri):
    """URIリンクを新旧API名に対応して追加"""
    payload = {"kind": fitz.LINK_URI, "from": rect, "uri": uri}
    try:
        return page.insert_link(payload)  # 新
    except AttributeError:
        return page.insertLink(payload)   # 旧


def _rect_valid(r: fitz.Rect) -> bool:
    return (r is not None) and (r.width > 1.0) and (r.height > 1.0)

# ====== ここから 追加のヘルパー ======

def _text_width(text: str, fontname: str, fontsize: float) -> float:
    """文字列の描画幅を取得（新旧APIを吸収）。失敗時は概算。"""
    try:
        return fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)  # 新
    except Exception:
        try:
            return fitz.getTextlength(text, fontname=fontname, fontsize=fontsize)  # 旧
        except Exception:
            return len(text) * fontsize * 0.6  # 概算

def _append_summary_pages(doc: fitz.Document,
                          entries: list[tuple[int, str]],
                          title: str = "QR Decode Summary",
                          fontname: str = "helv"):
    """
    通し番号とデコード結果の一覧ページを文末に追加。
    - entries: [(#番号, テキスト), ...]
    - 複数ページに自動改ページ
    """
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
    line_gap = body_fs * 1.35  # 行送りの目安

    def new_page():
        return doc.new_page(width=page_rect.width, height=page_rect.height)

    def write_title(p: fitz.Page):
        p.insert_text(
            fitz.Point(col_left, top),
            title,
            fontsize=title_fs,
            fontname=fontname,
            color=(0, 0, 0)
        )
        # 罫線（任意）
        p.draw_line(
            fitz.Point(col_left, top + title_fs * 0.6),
            fitz.Point(col_right, top + title_fs * 0.6),
            color=(0, 0, 0),
            width=0.7
        )
        return top + title_fs * 1.6  # 次のY

    def wrap_to_width(text: str, max_width: float) -> list[str]:
        """1行分を最大幅に収まるようラップ（URL等の無空白も対応）"""
        if not text:
            return [""]
        lines = []
        buf = ""
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
                    # 1文字でも溢れる場合は強制改行
                    lines.append(ch)
                    buf = ""
                else:
                    lines.append(buf)
                    buf = ch
        if buf:
            lines.append(buf)
        return lines

    # ページ生成 & タイトル
    page = new_page()
    y = write_title(page)

    # ヘッダ情報（任意表示）
    info_line = f"Total: {len(entries)}"
    page.insert_text(
        fitz.Point(col_left, y),
        info_line,
        fontsize=body_fs,
        fontname=fontname,
        color=(0, 0, 0)
    )
    y += line_gap

    # エントリを順に書く
    max_width = col_right - col_left
    for idx, txt in entries:
        prefix = f"#{idx}: "
        # プレフィックスと本体を一緒にラップ（1行目だけprefix付きで詰める）
        first_line_budget = max_width - _text_width(prefix, fontname, body_fs)
        wrapped = []
        for i, seg in enumerate(txt.split("\n")):
            seg_lines = wrap_to_width(seg, max_width if i else max(first_line_budget, 24))
            if i == 0 and seg_lines:
                # 先頭行にprefixを融合
                head = seg_lines[0]
                seg_lines[0] = prefix + head
            else:
                # 改行を挟んだ文はそのまま（prefixなし）
                pass
            wrapped.extend(seg_lines)
        if not wrapped:
            wrapped = [prefix]  # 空文字対策

        for line in wrapped:
            # 改ページ判定
            if y + line_gap > bottom:
                page = new_page()
                y = write_title(page)
            page.insert_text(
                fitz.Point(col_left, y),
                line,
                fontsize=body_fs,
                fontname=fontname,
                color=(0, 0, 0)
            )
            y += line_gap

# ====== 追加ヘルパー ここまで ======


# ====== 注釈付きPDFを書き出す ======
def export_annotated_pdf(input_bytes, detections_map, zoom_map, password=None):
    """
    - 検出枠：正方形＋半透明シアン塗り＋赤枠
    - コメント（Text注釈）：枠の左上“外側”に付箋（本文 = [#n] デコード文字列）
    - 通し番号（FreeText注釈）：枠の左下“内側”に、枠サイズの約1/4の正方形で #n を赤字表示（背景は白）
    - URLはリンク化。ただし左下のラベル領域は避けるため、リンクを2分割（上帯＋右側）
    - ★ 最終ページに「#n と デコード結果」の一覧を追加
    """
    doc = fitz.open(stream=input_bytes, filetype="pdf")
    try:
        if doc.needs_pass and password:
            if not doc.authenticate(password):
                raise RuntimeError("PDFパスワードが正しくありません。")

        global_idx = 1  # 文書全体の通し番号
        summary_entries: list[tuple[int, str]] = []  # ★ 一覧ページ用に収集

        # ページは昇順で処理
        for pidx in sorted(detections_map.keys()):
            page = doc.load_page(pidx)
            zoom = zoom_map.get(pidx, 3.0)
            dets = detections_map.get(pidx, [])

            for det in dets:
                pts = det["points"]
                txt = (det.get("text") or "").strip()

                # === 正方形枠 ===
                rect = _square_rect_from_points(pts, zoom, margin=4.0)

                # 半透明シアンの塗り
                fill_annot = page.add_rect_annot(rect)
                fill_annot.set_border(width=0)
                fill_annot.set_colors(stroke=None, fill=(0, 1, 1))
                fill_annot.set_opacity(0.30)
                fill_annot.update()

                # 赤枠
                border_annot = page.add_rect_annot(rect)
                border_annot.set_border(width=1.5)
                border_annot.set_colors(stroke=(1, 0, 0), fill=None)
                border_annot.set_opacity(1.0)
                border_annot.update()

                # ===== コメント（Text注釈）…枠の左上“外側” =====
                ICON_EST = 20.0  # 付箋アイコンの概寸（ページ座標）
                GAP      = 6.0
                offset   = ICON_EST + GAP
                bubble_pt = fitz.Point(rect.x0 - offset, rect.y0 - offset)

                # ページ範囲にクランプ
                pagebox = page.bound()
                bx = min(max(bubble_pt.x, pagebox.x0 + 2), pagebox.x1 - 2)
                by = min(max(bubble_pt.y, pagebox.y0 + 2), pagebox.y1 - 2)
                bubble_pt = fitz.Point(bx, by)

                contents_str = f"[#{global_idx}] {txt}"
                text_annot = _safe_add_text_annot(page, bubble_pt, contents_str, icon="Comment")
                try:
                    text_annot.set_info({
                        "content": contents_str,
                        "title":   f"QR #{global_idx}",
                        "subject": "QR decode",
                    })
                except Exception:
                    pass
                text_annot.set_colors(stroke=(1, 0, 0), fill=None)
                text_annot.update()

                # ===== 通し番号（FreeText注釈）…枠の左下“内側”（1/4サイズ・白背景） =====
                unit = max(12.0, min(rect.width, rect.height) / 4.0)
                margin = 2.0
                label_rect = fitz.Rect(
                    rect.x0 + margin,
                    rect.y1 - margin - unit,
                    rect.x0 + margin + unit,
                    rect.y1 - margin
                )
                fontsize = max(8.0, min(13.0, unit * 0.55))
                try:
                    ft = _safe_add_freetext_annot(
                        page, label_rect, f"#{global_idx}",
                        fontsize=fontsize,
                        fontname="helv",
                        text_color=(1, 0, 0),
                        fill_color=(1, 1, 1),      # ★ 白背景
                        align=fitz.TEXT_ALIGN_LEFT,
                        rotate=0,
                    )
                except TypeError:
                    ft = _safe_add_freetext_annot(
                        page, label_rect, f"#{global_idx}",
                        fontsize=fontsize,
                        text_color=(1, 0, 0),
                    )
                try:
                    ft.set_border(width=0.8)
                    if hasattr(ft, "set_opacity"):
                        ft.set_opacity(0.90)
                    ft.set_info({"title": f"QR #{global_idx}", "subject": "QR label"})
                    ft.update()
                except Exception:
                    pass

                # ===== URLリンク化（左下のラベル領域だけ避ける＝2分割） =====
                is_url = txt.lower().startswith("http://") or txt.lower().startswith("https://")
                if is_url:
                    link_top   = fitz.Rect(rect.x0, rect.y0, rect.x1, label_rect.y0)
                    link_right = fitz.Rect(label_rect.x1, label_rect.y0, rect.x1, rect.y1)
                    for lr in (link_top, link_right):
                        if _rect_valid(lr):
                            _safe_insert_link(page, lr, txt)

                # ★ 一覧ページ用に追記
                summary_entries.append((global_idx, txt if txt else ""))

                # 次のQRへ
                global_idx += 1

        # ===== ★ 文末にサマリーページを追加 =====
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
        self.drop_box = ttk.Label(self, text="ここに PDF をドラッグ＆ドロップ 📄👇",
                                  anchor="center", relief="groove", padding=18)
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
        # すでに実行中なら何もしない
        if self._worker and self._worker.is_alive():
            return
        path = self.pdf_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror("エラー", "PDFファイルを選択してください。")
            return

        # ★ パスワード事前確認（必要時）
        pwd = self.password.get().strip()
        try:
            test_doc = fitz.open(path)
            try:
                if test_doc.needs_pass and not pwd:
                    ask = simpledialog.askstring("パスワード",
                                                 "このPDFにはパスワードが必要です。入力してください：",
                                                 show="*", parent=self)
                    if not ask:
                        self.log_write("開始をキャンセルしました（パスワード未入力）。\n")
                        return
                    self.password.set(ask)
            finally:
                test_doc.close()
        except Exception:
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
