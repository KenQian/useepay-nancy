import logging
import os
import queue
import threading
import traceback

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
except Exception:  # pragma: no cover
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
    ScrolledText = None

from merchant_anomaly_detector import run_merchant_analyzer


class QueueLogHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__(level=logging.INFO)
        self.log_queue = log_queue
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S'))

    def emit(self, record):
        try:
            self.log_queue.put_nowait(self.format(record))
        except Exception:
            pass


class MerchantAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("商户交易异常检测")
        self.root.geometry("960x720")
        self.root.minsize(820, 600)

        self.log_queue = queue.Queue()
        self.log_handler = QueueLogHandler(self.log_queue)
        self.worker_thread = None
        self.result = None

        self.file_var = tk.StringVar()
        self.activity_var = tk.StringVar(value="请选择一个 Excel 文件，然后开始检测。")
        self.result_var = tk.StringVar(value="")
        self.activity_spinner_frames = ("|", "/", "-", "\\")
        self.activity_spinner_index = 0
        self.activity_spinner_after_id = None

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.poll_log_queue)

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main, text="选择一个 Excel 文件，然后点击开始检测。").pack(anchor=tk.W, pady=(4, 12))

        file_frame = ttk.LabelFrame(main, text="源文件", padding=10)
        file_frame.pack(fill=tk.X)

        ttk.Entry(file_frame, textvariable=self.file_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="选择文件", command=self.choose_file).pack(side=tk.LEFT, padx=(8, 0))

        action_frame = ttk.LabelFrame(main, text="检测操作", padding=10)
        action_frame.pack(fill=tk.X, pady=(12, 8))

        self.run_button = tk.Button(
            action_frame,
            text="开始检测",
            command=self.start_run,
            font=("Segoe UI", 11, "bold"),
            width=20,
            height=2,
            padx=18,
            pady=12,
            relief=tk.RAISED,
            bd=2,
        )
        self.run_button.pack(anchor=tk.W)

        activity_row = ttk.Frame(action_frame)
        activity_row.pack(fill=tk.X, pady=(8, 0))
        self.activity_icon = ttk.Label(activity_row, text="", font=("Segoe UI", 12, "bold"), foreground="#1d4ed8")
        self.activity_icon.pack(side=tk.LEFT, anchor=tk.N, padx=(0, 8))
        ttk.Label(activity_row, textvariable=self.activity_var, wraplength=880).pack(anchor=tk.W)

        result_frame = ttk.LabelFrame(main, text="处理结果", padding=10)
        result_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(result_frame, textvariable=self.result_var, wraplength=840, justify=tk.LEFT).pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(main, text="执行日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, height=20, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def choose_file(self):
        file_path = filedialog.askopenfilename(
            title="请选择要分析的 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if file_path:
            self.file_var.set(file_path)
            self.activity_var.set("已选择文件，点击“开始检测”。")

    def start_run(self):
        input_file = self.file_var.get().strip()
        if not input_file:
            messagebox.showwarning("未选择文件", "请先选择要分析的 Excel 文件。")
            return
        if not os.path.isfile(input_file):
            messagebox.showerror("文件不存在", "请选择有效的 Excel 文件。")
            return
        if self.worker_thread and self.worker_thread.is_alive():
            return

        self.result = None
        self.result_var.set("")
        self.activity_var.set("正在检测，请稍候。")
        self._append_log("")
        self._append_log(f"==== 开始检测：{input_file} ====")

        self._set_run_button_enabled(False)
        self._start_activity_spinner()

        logger = logging.getLogger()
        if self.log_handler not in logger.handlers:
            logger.addHandler(self.log_handler)

        self.worker_thread = threading.Thread(target=self._run_worker, args=(input_file,), daemon=True)
        self.worker_thread.start()

    def _run_worker(self, input_file):
        try:
            result = run_merchant_analyzer(input_file)
            self.result = ("success", result)
        except Exception as exc:
            traceback_text = traceback.format_exc() if os.environ.get("FX_TOOL_DEBUG_TRACEBACK") == "1" else ""
            self.result = ("error", {"exception": exc, "traceback": traceback_text, "input_file": input_file})

    def poll_log_queue(self):
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)
            self._update_status_from_log(message)

        if self.result is not None:
            outcome, payload = self.result
            self.result = None
            self._stop_activity_spinner()
            logger = logging.getLogger()
            if self.log_handler in logger.handlers:
                logger.removeHandler(self.log_handler)

            if outcome == "success":
                self.activity_var.set("检测完成。请打开 result 文件夹查看输出文件和日志。")
                result_lines = [f"日志文件：\n{payload['log_path']}"]
                if payload['output_file']:
                    result_lines.insert(0, f"输出文件：\n{payload['output_file']}\n")
                else:
                    result_lines.insert(0, f"结果：\n{payload['message']}\n")
                self.result_var.set("\n".join(result_lines))
            else:
                error_message = self._format_error_message(payload["exception"], payload["input_file"])
                self.activity_var.set("检测失败。请查看错误信息和日志。")
                self.result_var.set(error_message)
                if payload["traceback"]:
                    self._append_log(payload["traceback"])
                messagebox.showerror("处理失败", error_message)

            self._set_run_button_enabled(True)

        self.root.after(100, self.poll_log_queue)

    def _append_log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_status_from_log(self, message):
        if "Loading source workbook..." in message:
            self.activity_var.set("正在加载源文件...")
        elif "Pre-processing source data..." in message:
            self.activity_var.set("正在预处理数据...")
        elif "Analyzing merchants" in message:
            self.activity_var.set("正在检测商户交易异常...")
        elif "Processing merchants:" in message:
            self.activity_var.set("正在检测商户交易异常...")
        elif "Formatting anomaly report..." in message:
            self.activity_var.set("正在整理结果...")
        elif "Saving anomaly report workbook..." in message:
            self.activity_var.set("正在保存Excel文件...")
        elif "Report:" in message or "No significant anomalies found" in message:
            self.activity_var.set("检测完成。")
        elif "Starting Merchant Analyzer..." in message:
            self.activity_var.set("开始检测...")

    def _format_error_message(self, exc, input_file):
        result_dir = os.path.join(os.path.dirname(os.path.abspath(input_file)), 'result') if input_file else ""
        log_hint = f"\n\n请查看日志：\n{result_dir}" if result_dir else ""
        return f"处理失败：\n{exc}{log_hint}"

    def _set_run_button_enabled(self, enabled):
        if enabled:
            self.run_button.configure(state=tk.NORMAL, background="#ffffff", activebackground="#bfdbfe", fg="#0f172a", activeforeground="#0f172a")
        else:
            self.run_button.configure(state=tk.DISABLED, background="#f3f4f6", activebackground="#f3f4f6", fg="#8a8f98", activeforeground="#8a8f98", disabledforeground="#8a8f98")

    def _start_activity_spinner(self):
        if self.activity_spinner_after_id is not None:
            return
        self._animate_activity_spinner()

    def _animate_activity_spinner(self):
        frame = self.activity_spinner_frames[self.activity_spinner_index]
        self.activity_spinner_index = (self.activity_spinner_index + 1) % len(self.activity_spinner_frames)
        self.activity_icon.configure(text=frame)
        self.activity_spinner_after_id = self.root.after(120, self._animate_activity_spinner)

    def _stop_activity_spinner(self):
        if self.activity_spinner_after_id is not None:
            self.root.after_cancel(self.activity_spinner_after_id)
            self.activity_spinner_after_id = None
        self.activity_spinner_index = 0
        self.activity_icon.configure(text="")

    def on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("处理中", "程序仍在运行，关闭窗口不会终止后台处理。\n确定要关闭窗口吗？"):
                return
        self._stop_activity_spinner()
        logger = logging.getLogger()
        if self.log_handler in logger.handlers:
            logger.removeHandler(self.log_handler)
        self.root.destroy()


def main():
    if tk is None or filedialog is None or messagebox is None or ttk is None or ScrolledText is None:
        raise RuntimeError("无法启动图形界面，请确认系统支持 tkinter。")

    root = tk.Tk()
    MerchantAnalyzerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
