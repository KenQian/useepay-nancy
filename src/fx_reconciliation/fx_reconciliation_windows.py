import logging
import os
import queue
import threading
import traceback

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter.scrolledtext import ScrolledText
except Exception:  # pragma: no cover - platform/environment dependent
    tk = None
    filedialog = None
    messagebox = None
    ttk = None
    ScrolledText = None

from fx_reconciliation_core import run_fx_reconciliation


class QueueLogHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__(level=logging.INFO)
        self.log_queue = log_queue
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt='%H:%M:%S'))

    def emit(self, record):
        try:
            self.log_queue.put_nowait(self.format(record))
        except Exception:  # pragma: no cover - queue/UI fallback
            pass


class FxReconciliationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("FX 对账工具")
        self.root.geometry("900x650")
        self.root.minsize(760, 520)

        self.log_queue = queue.Queue()
        self.log_handler = QueueLogHandler(self.log_queue)
        self.worker_thread = None
        self.result = None

        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择包含源文件的文件夹。")
        self.latest_var = tk.StringVar(value="等待开始。")
        self.result_var = tk.StringVar(value="")

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.poll_log_queue)

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main, text="FX 对账工具", font=("Segoe UI", 16, "bold"))
        title.pack(anchor=tk.W)

        subtitle = ttk.Label(main, text="双击启动后，选择源文件夹并点击开始处理。")
        subtitle.pack(anchor=tk.W, pady=(4, 12))

        folder_frame = ttk.LabelFrame(main, text="源文件夹", padding=10)
        folder_frame.pack(fill=tk.X)

        entry = ttk.Entry(folder_frame, textvariable=self.folder_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(folder_frame, text="选择文件夹", command=self.choose_folder).pack(side=tk.LEFT, padx=(8, 0))

        action_frame = ttk.Frame(main)
        action_frame.pack(fill=tk.X, pady=(12, 8))

        self.run_button = ttk.Button(action_frame, text="开始处理", command=self.start_run)
        self.run_button.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(12, 0))

        status_frame = ttk.LabelFrame(main, text="当前状态", padding=10)
        status_frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(status_frame, textvariable=self.status_var, font=("Segoe UI", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(status_frame, textvariable=self.latest_var, wraplength=840).pack(anchor=tk.W, pady=(6, 0))

        result_frame = ttk.LabelFrame(main, text="处理结果", padding=10)
        result_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(result_frame, textvariable=self.result_var, wraplength=840, justify=tk.LEFT).pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(main, text="执行日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, height=20, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def choose_folder(self):
        folder = filedialog.askdirectory(title="请选择包含源文件的文件夹")
        if folder:
            self.folder_var.set(folder)
            self.status_var.set("已选择文件夹，点击“开始处理”。")
            self.latest_var.set(folder)

    def start_run(self):
        source_root = self.folder_var.get().strip()
        if not source_root:
            messagebox.showwarning("未选择文件夹", "请先选择包含源文件的文件夹。")
            return
        if not os.path.isdir(source_root):
            messagebox.showerror("文件夹不存在", "请选择有效的文件夹。")
            return
        if self.worker_thread and self.worker_thread.is_alive():
            return

        self.result = None
        self.result_var.set("")
        self.status_var.set("正在处理，请稍候。")
        self.latest_var.set("准备启动...")
        self._append_log("")
        self._append_log(f"==== 开始处理：{source_root} ====")

        self.run_button.state(["disabled"])
        self.progress.start(10)

        logger = logging.getLogger()
        if self.log_handler not in logger.handlers:
            logger.addHandler(self.log_handler)

        self.worker_thread = threading.Thread(target=self._run_worker, args=(source_root,), daemon=True)
        self.worker_thread.start()

    def _run_worker(self, source_root):
        try:
            result = run_fx_reconciliation(source_root)
            self.result = ("success", result)
        except Exception as exc:  # pragma: no cover - background error path
            traceback_text = traceback.format_exc() if os.environ.get("FX_TOOL_DEBUG_TRACEBACK") == "1" else ""
            self.result = ("error", {"exception": exc, "traceback": traceback_text, "source_root": source_root})

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
            self.progress.stop()
            self.run_button.state(["!disabled"])
            logger = logging.getLogger()
            if self.log_handler in logger.handlers:
                logger.removeHandler(self.log_handler)

            if outcome == "success":
                self.status_var.set("处理完成。")
                self.latest_var.set("请打开 result 文件夹查看输出文件和日志。")
                self.result_var.set(
                    "输出文件：\n"
                    f"{payload['final_path']}\n\n"
                    "日志文件：\n"
                    f"{payload['log_path']}"
                )
                messagebox.showinfo("处理完成", self.result_var.get())
            else:
                error_message = self._format_error_message(payload["exception"], payload["source_root"])
                self.status_var.set("处理失败。")
                self.latest_var.set(str(payload["exception"]))
                self.result_var.set(error_message)
                if payload["traceback"]:
                    self._append_log(payload["traceback"])
                messagebox.showerror("处理失败", error_message)

        self.root.after(100, self.poll_log_queue)

    def _append_log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_status_from_log(self, message):
        if "Phase complete:" in message:
            phase = message.split("Phase complete:", 1)[1].split("elapsed=", 1)[0].strip()
            self.status_var.set(f"当前阶段：{phase}")
            self.latest_var.set(message)
        elif "Processing 渠道订单 rows:" in message:
            self.status_var.set("正在处理渠道订单...")
            self.latest_var.set(message)
        elif "Writing 账户流水 rows:" in message:
            self.status_var.set("正在写入账户流水...")
            self.latest_var.set(message)
        elif "Saving workbook to WIP path..." in message:
            self.status_var.set("正在保存Excel文件...")
            self.latest_var.set(message)
        elif "COMPLETED. File:" in message:
            self.status_var.set("处理完成。")
            self.latest_var.set(message)
        elif "Starting FX Settlement Automation..." in message:
            self.status_var.set("开始处理...")
            self.latest_var.set(message)

    def _format_error_message(self, exc, source_root):
        result_dir = os.path.join(source_root, 'result') if source_root else ""
        log_hint = f"\n\n请查看日志：\n{result_dir}" if result_dir else ""
        return f"处理失败：\n{exc}{log_hint}"

    def on_close(self):
        if self.worker_thread and self.worker_thread.is_alive():
            if not messagebox.askyesno("处理中", "程序仍在运行，关闭窗口不会终止后台处理。\n确定要关闭窗口吗？"):
                return
        logger = logging.getLogger()
        if self.log_handler in logger.handlers:
            logger.removeHandler(self.log_handler)
        self.root.destroy()


def main():
    if tk is None or filedialog is None or messagebox is None or ttk is None or ScrolledText is None:
        raise RuntimeError("无法启动图形界面，请确认系统支持 tkinter。")

    root = tk.Tk()
    app = FxReconciliationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
