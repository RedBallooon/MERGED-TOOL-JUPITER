# MERGE JUPITER TOOL - desktop app
# pip install fastexcel polars xlsxwriter tkinterdnd2 pyinstaller
import os
import sys
import time
import queue
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext

import fastexcel
import polars as pl
import xlsxwriter

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

APP_NAME = "Merge Jupiter Tool"
ICON_FILE = "icon.ico"          # icon.ico hoặc icon.png đặt cạnh file này (thêm sau)
ICON_PNG = "icon.png"
FACTORY_ORDER = ["VH", "VH2", "JV", "JV2"]
ROWS_TO_DROP = 3                # số dòng đầu bị xóa trước khi merge (dòng kế tiếp làm header)
EXCEL_ROW_LIMIT = 1_048_576


def resource_path(name: str) -> str:
    """Đường dẫn tài nguyên, chạy được cả khi đóng gói bằng PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def sort_by_factory(df: pl.DataFrame) -> pl.DataFrame:
    """Sắp xếp theo FACTORY_ORDER nếu có cột Factory (giữ nguyên thứ tự gốc trong từng nhóm)."""
    if "Factory" not in df.columns:
        return df
    key = (
        pl.col("Factory").cast(pl.String).str.strip_chars()
        .replace_strict(
            {name: i for i, name in enumerate(FACTORY_ORDER)},
            default=len(FACTORY_ORDER),
            return_dtype=pl.Int32,
        )
    )
    return (
        df.with_columns(key.alias("__k"))
        .sort("__k", maintain_order=True)
        .drop("__k")
    )


# --- Logic merge ---------------------------------------------------------
def run_merge(root_folder: str, log, set_progress) -> str | None:
    """Trả về đường dẫn file kết quả (hoặc None nếu không có dữ liệu)."""
    output_file = os.path.join(
        os.path.dirname(root_folder), "MERGED_FILE", "Merged_report.xlsx"
    )
    output_abspath = os.path.abspath(output_file)

    # Liệt kê file trước để có thanh tiến trình
    targets = []
    for current_root, dirs, files in os.walk(root_folder):
        dirs[:] = [d for d in dirs if d.casefold() != "merged_file"]
        for filename in files:
            if filename.startswith("~$") or not filename.casefold().endswith((".xlsx", ".xls")):
                continue
            fp = os.path.join(current_root, filename)
            if os.path.abspath(fp) == output_abspath:
                continue
            targets.append((current_root, fp))

    total = len(targets)
    log(f"Tìm thấy {total} file Excel.")
    if total == 0:
        return None

    frames, matched, failed, multi_sheet = [], [], [], []
    t0 = time.time()

    for idx, (current_root, file_path) in enumerate(targets, start=1):
        file_start = time.time()
        try:
            xl = fastexcel.read_excel(file_path)
            sheet = xl.sheet_names[0]
            if len(xl.sheet_names) > 1:
                multi_sheet.append((file_path, xl.sheet_names))
                log(f"[CHÚ Ý] File có {len(xl.sheet_names)} sheet, chỉ đọc '{sheet}': {file_path}")

            df = xl.load_sheet_by_name(sheet, header_row=ROWS_TO_DROP).to_polars()
            matched.append(file_path)

            if not df.is_empty():
                folder_name = os.path.basename(current_root)
                c_file = "Source_File" if "Source_File" not in df.columns else "Source_File_"
                c_sheet = "Source_Sheet" if "Source_Sheet" not in df.columns else "Source_Sheet_"
                df = df.select(
                    pl.lit(folder_name).alias(c_file),
                    pl.lit(sheet).alias(c_sheet),
                    pl.all(),
                )
                frames.append(df)

            log(f"[OK] ({time.time() - file_start:.1f}s): {file_path}")
        except Exception as e:
            failed.append((file_path, str(e)))
            log(f"[LỖI] Không đọc được {file_path}: {e}")

        set_progress(idx / total * 90)  # 90% cho giai đoạn đọc

    log(f"\nTổng thời gian đọc: {time.time() - t0:.1f}s")

    if not frames:
        log("Không tìm thấy dữ liệu phù hợp để gộp.")
        return None

    try:
        merged = pl.concat(frames, how="diagonal_relaxed")
    except Exception as e:
        log(f"[CẢNH BÁO] Gộp theo kiểu thất bại ({e}); chuyển toàn bộ sang String.")
        frames = [f.cast(pl.String, strict=False) for f in frames]
        merged = pl.concat(frames, how="diagonal")

    merged = sort_by_factory(merged)

    log("Đang ghi file Excel...")
    write_start = time.time()
    os.makedirs(os.path.dirname(output_abspath), exist_ok=True)
    chunk_size = EXCEL_ROW_LIMIT - 1

    xlsx_options = {
        "strings_to_urls": False,
        "strings_to_formulas": False,
        "strings_to_numbers": False,
        "nan_inf_to_errors": True,
    }

    with xlsxwriter.Workbook(output_abspath, xlsx_options) as wb:
        for i, start in enumerate(range(0, merged.height, chunk_size), start=1):
            sheet_name = "Merged_Data" if i == 1 else f"Merged_Data_{i}"
            merged.slice(start, chunk_size).write_excel(
                workbook=wb,
                worksheet=sheet_name,
                table_style=None,
                autofit=False,
                header_format={"bold": False},
                dtype_formats={
                    pl.Datetime: "yyyy-mm-dd hh:mm:ss",
                    pl.Date: "yyyy-mm-dd",
                },
            )

    set_progress(100)
    log(f"Thời gian ghi: {time.time() - write_start:.1f}s")
    log(f"Đã đọc thành công {len(matched)} file.")
    log(f"Tổng số dòng dữ liệu: {merged.height:,}")
    log(f"Đã lưu kết quả tại: {output_abspath}")

    if multi_sheet:
        log(f"\n{len(multi_sheet)} file có nhiều hơn 1 sheet (chỉ đọc sheet đầu):")
        for p, sheets in multi_sheet:
            log(f"  - {p} (sheet có: {sheets})")
    if failed:
        log(f"\n{len(failed)} file bị lỗi:")
        for p, err in failed:
            log(f"  - {p}: {err}")

    return output_abspath


# --- Giao diện -----------------------------------------------------------
class App:
    def __init__(self):
        self.root = TkinterDnD.Tk() if HAS_DND else tk.Tk()
        self.root.title(APP_NAME)
        self.root.geometry("760x560")
        self.root.minsize(620, 460)

        ico = resource_path(ICON_FILE)
        png = resource_path(ICON_PNG)
        try:
            if os.path.exists(ico):
                self.root.iconbitmap(ico)
            elif os.path.exists(png):
                self._icon_img = tk.PhotoImage(file=png)  # giữ tham chiếu
                self.root.iconphoto(True, self._icon_img)
        except tk.TclError:
            pass

        self.q: queue.Queue = queue.Queue()
        self.result_path = None
        self.running = False

        self._build_ui()
        self.root.after(100, self._poll)

    def _build_ui(self):
        pad = {"padx": 12}

        msg = ("KÉO & THẢ thư mục chứa các subfolder vào đây"
               if HAS_DND else
               "(Chưa cài tkinterdnd2 -> dùng Browse hoặc dán đường dẫn)")
        self.drop_area = tk.Label(self.root, text=msg, relief="groove", bd=2,
                                  bg="#eef3fb", font=("Segoe UI", 11), height=5)
        self.drop_area.pack(fill="x", pady=(12, 6), **pad)

        if HAS_DND:
            self.drop_area.drop_target_register(DND_FILES)
            self.drop_area.dnd_bind("<<Drop>>", self._on_drop)

        row = tk.Frame(self.root)
        row.pack(fill="x", **pad)
        self.path_var = tk.StringVar()
        tk.Entry(row, textvariable=self.path_var).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse...", width=10, command=self._browse).pack(side="left", padx=(6, 0))

        btns = tk.Frame(self.root)
        btns.pack(fill="x", pady=8, **pad)
        self.run_btn = tk.Button(btns, text="Bắt đầu merge", width=16,
                                 bg="#2f6fdd", fg="white", command=self._start)
        self.run_btn.pack(side="left")
        self.open_btn = tk.Button(btns, text="Mở file kết quả", width=16,
                                  state="disabled", command=self._open_result)
        self.open_btn.pack(side="left", padx=6)
        self.folder_btn = tk.Button(btns, text="Mở thư mục", width=12,
                                    state="disabled", command=self._open_folder)
        self.folder_btn.pack(side="left")

        self.progress = ttk.Progressbar(self.root, maximum=100)
        self.progress.pack(fill="x", **pad)

        self.status = tk.StringVar(value="Sẵn sàng.")
        tk.Label(self.root, textvariable=self.status, anchor="w").pack(fill="x", **pad)

        self.log_box = scrolledtext.ScrolledText(self.root, height=12, state="disabled",
                                                 font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, pady=(0, 12), **pad)

    # -- chọn thư mục
    def _accept(self, path: str):
        path = path.strip().strip('"')
        if os.path.isfile(path):
            path = os.path.dirname(path)
        if os.path.isdir(path):
            self.path_var.set(os.path.abspath(path))
            self.status.set("Đã chọn thư mục. Bấm 'Bắt đầu merge'.")
        else:
            self.status.set(f"Đường dẫn không hợp lệ: {path}")

    def _on_drop(self, event):
        items = self.root.tk.splitlist(event.data)
        if items:
            self._accept(items[0])

    def _browse(self):
        p = filedialog.askdirectory(parent=self.root, title="Chọn thư mục")
        if p:
            self._accept(p)

    # -- chạy
    def _start(self):
        if self.running:
            return
        folder = self.path_var.get().strip().strip('"')
        if not os.path.isdir(folder):
            self.status.set("Bạn chưa chọn thư mục hợp lệ.")
            return

        self.running = True
        self.result_path = None
        self.run_btn.config(state="disabled")
        self.open_btn.config(state="disabled")
        self.folder_btn.config(state="disabled")
        self.progress["value"] = 0
        self.log_box.config(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.config(state="disabled")
        self.status.set("Đang xử lý...")
        self._log(f"Thư mục nguồn: {folder}")

        threading.Thread(target=self._worker, args=(folder,), daemon=True).start()

    def _worker(self, folder):
        try:
            result = run_merge(
                folder,
                log=lambda m: self.q.put(("log", m)),
                set_progress=lambda v: self.q.put(("progress", v)),
            )
            self.q.put(("done", result))
        except Exception:
            self.q.put(("log", traceback.format_exc()))
            self.q.put(("error", None))

    def _poll(self):
        try:
            while True:
                kind, val = self.q.get_nowait()
                if kind == "log":
                    self._log(val)
                elif kind == "progress":
                    self.progress["value"] = val
                elif kind == "done":
                    self._finish(val)
                elif kind == "error":
                    self.running = False
                    self.run_btn.config(state="normal")
                    self.status.set("Có lỗi xảy ra, xem log bên dưới.")
        except queue.Empty:
            pass
        self.root.after(100, self._poll)

    def _finish(self, result):
        self.running = False
        self.run_btn.config(state="normal")
        if result:
            self.result_path = result
            self.open_btn.config(state="normal")
            self.folder_btn.config(state="normal")
            self.status.set("Hoàn tất!")
        else:
            self.status.set("Không có dữ liệu để gộp.")

    def _log(self, msg):
        self.log_box.config(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def _open_result(self):
        if self.result_path and os.path.exists(self.result_path):
            os.startfile(self.result_path)

    def _open_folder(self):
        if self.result_path:
            os.startfile(os.path.dirname(self.result_path))

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()