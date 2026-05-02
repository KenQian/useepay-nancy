import logging
import json
import os
import queue
import threading
import traceback
from pathlib import Path

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

from finalize_fx_summary_report import finalize_fx_summary_report
from prepare_fx_summary_workbook import prepare_fx_summary_workbook


STATE_FILE_PATH = Path.home() / ".fx_summary_workflow_state.json"
STATE_LAST_FOLDER_KEY = "last_selected_folder"


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
        self.root.title("生成各通道需换汇情况汇总")
        self.root.geometry("960x760")
        self.root.minsize(820, 620)

        self.log_queue = queue.Queue()
        self.log_handler = QueueLogHandler(self.log_queue)
        self.worker_thread = None
        self.pending_result = None
        self.current_stage = None
        self.current_panel = 1

        self.prepared_result = None
        self.manual_input_items = []
        self.manual_item_vars = []
        self.stepper_labels = {}
        self.step_activity_rows = {}
        self.step_activity_messages = {}
        self.step_activity_icons = {}
        self.step_activity_spinner_frames = {}
        self.step_activity_spinner_indices = {}
        self.step_activity_spinner_after_ids = {}
        self.step_panels = {}
        self.session_completed = False
        self.button_enabled = {}

        self.last_selected_folder = self._load_last_selected_folder()
        self.folder_var = tk.StringVar()
        self.result_var = tk.StringVar(value="")
        self.manual_hint_var = tk.StringVar(value="将在生成工作簿后显示需要补充的数据")

        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.root.after(100, self.poll_log_queue)

    def _build_ui(self):
        self._configure_styles()

        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        subtitle = ttk.Label(main, text="请按以下步骤完成操作")
        subtitle.pack(anchor=tk.W, pady=(0, 12))

        stepper_frame = ttk.Frame(main)
        stepper_frame.pack(fill=tk.X, pady=(0, 12))
        self._build_stepper(stepper_frame)

        self.panel_container = ttk.Frame(main)
        self.panel_container.pack(fill=tk.X, pady=(0, 8))

        self._build_step1_panel()
        self._build_step2_panel()
        self._build_step3_panel()

        result_frame = ttk.LabelFrame(main, text="处理结果", padding=10)
        result_frame.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(result_frame, textvariable=self.result_var, wraplength=900, justify=tk.LEFT).pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(main, text="执行日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(log_frame, wrap=tk.WORD, height=22, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.show_step_panel(1)
        self._set_stepper_state(active_step=1)

    def _build_step1_panel(self):
        panel = ttk.LabelFrame(self.panel_container, text="1 生成工作簿", padding=10)
        self.step_panels[1] = panel

        folder_frame = ttk.Frame(panel)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        entry = ttk.Entry(folder_frame, textvariable=self.folder_var)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(folder_frame, text="选择", command=self.choose_folder).pack(side=tk.LEFT, padx=(8, 0))

        self.prepare_button = tk.Button(
            panel,
            text="开始生成工作簿",
            command=self._on_prepare_click,
            font=("Segoe UI", 11, "bold"),
            width=20,
            height=2,
            padx=18,
            pady=12,
            relief=tk.RAISED,
            bd=2,
        )
        self.prepare_button.pack(anchor=tk.W)
        self.button_enabled[self.prepare_button] = False
        self._bind_button_hover(self.prepare_button)

        self._build_step_activity_row(panel, 1)

    def _build_step2_panel(self):
        panel = ttk.LabelFrame(self.panel_container, text="2 补充数据", padding=10)
        self.step_panels[2] = panel

        ttk.Label(
            panel,
            textvariable=self.manual_hint_var,
            wraplength=900,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        self.manual_items_frame = ttk.Frame(panel)
        self.manual_items_frame.pack(fill=tk.X, pady=(8, 0))

    def _build_step3_panel(self):
        panel = ttk.LabelFrame(self.panel_container, text="3 生成汇总", padding=10)
        self.step_panels[3] = panel

        self.finalize_button = tk.Button(
            panel,
            text="生成最终汇总",
            command=self._on_finalize_click,
            font=("Segoe UI", 11, "bold"),
            width=20,
            height=2,
            padx=18,
            pady=12,
            relief=tk.RAISED,
            bd=2,
        )
        self.finalize_button.pack(anchor=tk.W)
        self.button_enabled[self.finalize_button] = False
        self._bind_button_hover(self.finalize_button)

        self._build_step_activity_row(panel, 3)

    def _configure_styles(self):
        style = ttk.Style()
        try:
            style.theme_use(style.theme_use())
        except Exception:  # pragma: no cover - theme availability varies
            pass

        style.configure(
            "StepperPending.TLabel",
            font=("Segoe UI", 26, "bold"),
            foreground="#6b7280",
            padding=(8, 6),
        )
        style.configure(
            "StepperActive.TLabel",
            font=("Segoe UI", 26, "bold"),
            foreground="#1d4ed8",
            background="#dbeafe",
            padding=(10, 6),
        )
        style.configure(
            "StepperDone.TLabel",
            font=("Segoe UI", 26, "bold"),
            foreground="#166534",
            background="#dcfce7",
            padding=(10, 6),
        )
        style.configure(
            "StepActivity.TLabel",
            font=("Segoe UI", 10),
            foreground="#475569",
            padding=(0, 6),
        )
        style.configure(
            "StepSpinner.TLabel",
            font=("Segoe UI", 12, "bold"),
            foreground="#1d4ed8",
            padding=(0, 4),
        )

    def _build_stepper(self, parent):
        step_texts = {
            1: "1 生成工作簿",
            2: "2 补充数据",
            3: "3 生成汇总",
        }
        for idx in (1, 2, 3):
            label = ttk.Label(parent, text=step_texts[idx], style="StepperPending.TLabel")
            label.pack(side=tk.LEFT)
            self.stepper_labels[idx] = label
            if idx < 3:
                ttk.Label(parent, text="→", style="StepperPending.TLabel").pack(side=tk.LEFT, padx=6)

    def _build_step_activity_row(self, parent, step_number):
        row = ttk.Frame(parent)
        icon = ttk.Label(row, text="", style="StepSpinner.TLabel")
        icon.pack(side=tk.LEFT, anchor=tk.N, padx=(0, 8))
        text_frame = ttk.Frame(row)
        text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        message = tk.StringVar(value="")
        label = ttk.Label(text_frame, textvariable=message, style="StepActivity.TLabel")
        label.pack(anchor=tk.W)
        self.step_activity_rows[step_number] = row
        self.step_activity_messages[step_number] = message
        self.step_activity_icons[step_number] = icon
        self.step_activity_spinner_frames[step_number] = ("|", "/", "-", "\\")
        self.step_activity_spinner_indices[step_number] = 0
        self.step_activity_spinner_after_ids[step_number] = None

    def _start_spinner(self, step_number):
        after_id = self.step_activity_spinner_after_ids.get(step_number)
        if after_id is not None:
            return
        self._animate_spinner(step_number)

    def _animate_spinner(self, step_number):
        frames = self.step_activity_spinner_frames[step_number]
        index = self.step_activity_spinner_indices[step_number]
        self.step_activity_icons[step_number].configure(text=frames[index])
        self.step_activity_spinner_indices[step_number] = (index + 1) % len(frames)
        self.step_activity_spinner_after_ids[step_number] = self.root.after(
            120,
            lambda: self._animate_spinner(step_number),
        )

    def _stop_spinner(self, step_number):
        after_id = self.step_activity_spinner_after_ids.get(step_number)
        if after_id is not None:
            self.root.after_cancel(after_id)
            self.step_activity_spinner_after_ids[step_number] = None
        self.step_activity_spinner_indices[step_number] = 0
        self.step_activity_icons[step_number].configure(text="")

    def show_step_panel(self, step_number):
        for panel_number, panel in self.step_panels.items():
            if panel_number == step_number:
                if not panel.winfo_ismapped():
                    panel.pack(fill=tk.X)
            elif panel.winfo_ismapped():
                panel.pack_forget()
        self.current_panel = step_number

    def _show_step_activity(self, step_number, message):
        for current_step, row in self.step_activity_rows.items():
            if current_step == step_number:
                self.step_activity_messages[current_step].set(message)
                if not row.winfo_ismapped():
                    row.pack(fill=tk.X, pady=(10, 0))
                self._start_spinner(current_step)
            else:
                self._hide_step_activity(current_step)

    def _hide_step_activity(self, step_number):
        row = self.step_activity_rows.get(step_number)
        if row is None:
            return
        self._stop_spinner(step_number)
        self.step_activity_messages[step_number].set("")
        if row.winfo_ismapped():
            row.pack_forget()

    def _clear_step_activity(self):
        for step_number in self.step_activity_rows:
            self._hide_step_activity(step_number)

    def _set_stepper_state(self, active_step=None, completed_steps=None):
        completed_steps = set(completed_steps or [])
        for step_number, label in self.stepper_labels.items():
            if step_number in completed_steps:
                label.configure(style="StepperDone.TLabel")
            elif step_number == active_step:
                label.configure(style="StepperActive.TLabel")
            else:
                label.configure(style="StepperPending.TLabel")

    def _bind_button_hover(self, button):
        button.bind("<Enter>", lambda _event, btn=button: self._set_button_hover(btn, True))
        button.bind("<Leave>", lambda _event, btn=button: self._set_button_hover(btn, False))
        self._sync_button_interaction_state(button)

    def _set_button_hover(self, button, hovering):
        if self._is_button_disabled(button):
            button.configure(cursor="", background="#f3f4f6", activebackground="#f3f4f6")
            return
        button.configure(
            cursor="",
            background="#dbeafe" if hovering else "#ffffff",
            activebackground="#bfdbfe",
        )

    def _sync_button_interaction_state(self, button):
        if self._is_button_disabled(button):
            button.configure(
                cursor="",
                background="#f3f4f6",
                activebackground="#f3f4f6",
                fg="#8a8f98",
                activeforeground="#8a8f98",
                disabledforeground="#8a8f98",
            )
        else:
            button.configure(
                cursor="",
                background="#ffffff",
                activebackground="#bfdbfe",
                fg="#0f172a",
                activeforeground="#0f172a",
                disabledforeground="#8a8f98",
            )

    def _is_button_disabled(self, button):
        return not self.button_enabled.get(button, True)

    def _set_button_enabled(self, button, enabled):
        self.button_enabled[button] = enabled
        self._sync_button_interaction_state(button)

    def choose_folder(self):
        folder = filedialog.askdirectory(
            title="请选择包含源文件的文件夹",
            initialdir=self._get_folder_dialog_initialdir(),
        )
        if folder:
            self.folder_var.set(folder)
            self.last_selected_folder = folder
            self._save_last_selected_folder(folder)
            self.result_var.set(f"已选择源文件夹：\n{folder}")
            if not self.session_completed and self.worker_thread is None:
                self._set_button_enabled(self.prepare_button, True)
                self.root.update_idletasks()

    def _on_prepare_click(self):
        if self._is_button_disabled(self.prepare_button):
            return
        self.start_prepare()

    def _on_finalize_click(self):
        if self._is_button_disabled(self.finalize_button):
            return
        self.start_finalize()

    def start_prepare(self):
        if self.session_completed:
            return
        source_root = self.folder_var.get().strip()
        if not source_root:
            messagebox.showwarning("未选择文件夹", "请先选择包含源文件的文件夹。")
            return
        if not os.path.isdir(source_root):
            messagebox.showerror("文件夹不存在", "请选择有效的文件夹。")
            return
        self.last_selected_folder = source_root
        self._save_last_selected_folder(source_root)
        if self.worker_thread and self.worker_thread.is_alive():
            return

        self.prepared_result = None
        self.manual_input_items = []
        self._render_manual_input_items([])
        self.show_step_panel(1)
        self._set_stepper_state(active_step=1)
        self.result_var.set("正在生成工作簿，请稍候。")
        self._show_step_activity(1, "正在生成工作簿...")
        self._append_log("")
        self._append_log(f"==== 第 1 步开始：{source_root} ====")
        self._start_worker("prepare", source_root)

    def start_finalize(self):
        if not self.prepared_result:
            messagebox.showwarning("缺少处理中工作簿", "请先执行第 1 步。")
            return
        if not self._can_finalize():
            messagebox.showwarning("未完成确认", "请先勾选所有需要人工补充的项目。")
            return
        if self.worker_thread and self.worker_thread.is_alive():
            return

        workbook_path = self.prepared_result["workbook_path"]
        if not os.path.isfile(workbook_path):
            messagebox.showerror("文件不存在", f"处理中工作簿不存在：\n{workbook_path}")
            return

        self.show_step_panel(3)
        self._set_stepper_state(active_step=3, completed_steps={1, 2})
        self.result_var.set(
            "处理中工作簿已确认完成。\n\n"
            "正在生成最终汇总，请稍候。"
        )
        self._show_step_activity(3, "正在生成最终汇总...")
        self._append_log("")
        self._append_log(f"==== 第 3 步开始：{workbook_path} ====")
        self._start_worker(
            "finalize",
            {
                "workbook_path": workbook_path,
                "log_path": self.prepared_result.get("log_path"),
            },
        )

    def _start_worker(self, stage, stage_input):
        self.pending_result = None
        self.current_stage = stage
        self._set_running(True)

        logger = logging.getLogger()
        if self.log_handler not in logger.handlers:
            logger.addHandler(self.log_handler)

        self.worker_thread = threading.Thread(
            target=self._run_worker,
            args=(stage, stage_input),
            daemon=True,
        )
        self.worker_thread.start()

    def _run_worker(self, stage, stage_input):
        try:
            if stage == "prepare":
                result = prepare_fx_summary_workbook(stage_input)
            elif stage == "finalize":
                result = finalize_fx_summary_report(
                    stage_input["workbook_path"],
                    log_path=stage_input.get("log_path"),
                )
            else:  # pragma: no cover - defensive branch
                raise ValueError(f"Unsupported stage: {stage}")
            self.pending_result = ("success", stage, result)
        except Exception as exc:  # pragma: no cover - background error path
            traceback_text = traceback.format_exc() if os.environ.get("FX_TOOL_DEBUG_TRACEBACK") == "1" else ""
            self.pending_result = (
                "error",
                stage,
                {
                    "exception": exc,
                    "traceback": traceback_text,
                    "stage_input": stage_input,
                },
            )

    def poll_log_queue(self):
        self._drain_log_queue()

        if self.pending_result is not None:
            outcome, stage, payload = self.pending_result
            self.pending_result = None
            self._set_running(False)
            logger = logging.getLogger()
            if self.log_handler in logger.handlers:
                logger.removeHandler(self.log_handler)
            self._drain_log_queue()

            if outcome == "success":
                if stage == "prepare":
                    self._handle_prepare_success(payload)
                else:
                    self._handle_finalize_success(payload)
            else:
                self._handle_error(stage, payload)

        self.root.after(100, self.poll_log_queue)

    def _get_folder_dialog_initialdir(self):
        candidates = [self.folder_var.get().strip(), self.last_selected_folder]
        for candidate in candidates:
            if candidate and os.path.isdir(candidate):
                return candidate
        return os.path.expanduser("~")

    def _load_last_selected_folder(self):
        try:
            with STATE_FILE_PATH.open("r", encoding="utf-8") as file_obj:
                payload = json.load(file_obj)
        except FileNotFoundError:
            return ""
        except Exception:  # pragma: no cover - persistence should not block UI
            return ""

        folder = payload.get(STATE_LAST_FOLDER_KEY, "")
        if folder and os.path.isdir(folder):
            return folder
        return ""

    def _save_last_selected_folder(self, folder):
        try:
            STATE_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
            with STATE_FILE_PATH.open("w", encoding="utf-8") as file_obj:
                json.dump({STATE_LAST_FOLDER_KEY: folder}, file_obj, ensure_ascii=False, indent=2)
        except Exception:  # pragma: no cover - persistence should not block UI
            pass

    def _drain_log_queue(self):
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self._append_log(message)
            self._update_status_from_log(message)

    def _handle_prepare_success(self, payload):
        self.prepared_result = payload
        self.manual_input_items = payload.get("manual_input_items", [])
        self._render_manual_input_items(self.manual_input_items)
        self._update_finalize_button_state()
        self._clear_step_activity()

        workbook_path = payload["workbook_path"]
        log_path = payload["log_path"]
        final_report_path = payload.get("final_report_path", "")
        if self.manual_input_items:
            self.show_step_panel(2)
            self._set_stepper_state(active_step=2, completed_steps={1})
            manual_lines = "\n".join(f"- {item['display_label']}" for item in self.manual_input_items)
            self.result_var.set(
                "工作簿已生成：\n"
                f"{workbook_path}\n\n"
                "请补充以下数据：\n"
                f"{manual_lines}\n\n"
                "完成后勾选当前页面中的项目，系统将进入下一步。\n\n"
                "日志文件：\n"
                f"{log_path}"
            )
        else:
            self.show_step_panel(3)
            self._set_stepper_state(active_step=3, completed_steps={1, 2})
            self.result_var.set(
                "工作簿已生成：\n"
                f"{workbook_path}\n\n"
                "无需补充数据，请点击“生成最终汇总”。\n\n"
                "最终文件：\n"
                f"{final_report_path}\n\n"
                "日志文件：\n"
                f"{log_path}"
            )

    def _handle_finalize_success(self, payload):
        final_path = payload["final_path"]
        log_path = payload["log_path"]
        self.session_completed = True
        self.show_step_panel(3)
        self._set_stepper_state(completed_steps={1, 2, 3})
        self._clear_step_activity()
        self.result_var.set(
            "最终文件：\n"
            f"{final_path}\n\n"
            "日志文件：\n"
            f"{log_path}\n\n"
            "如需重新执行，请关闭后重新打开窗口。"
        )
        self.prepared_result = None
        self.manual_input_items = []
        self._render_manual_input_items([])
        self._update_finalize_button_state()

    def _handle_error(self, stage, payload):
        stage_name = "第 1 步" if stage == "prepare" else "第 3 步"
        error_message = self._format_error_message(payload["exception"], stage_name, payload["stage_input"])
        if stage == "prepare":
            self.show_step_panel(1)
            self._set_stepper_state(active_step=1)
            self._show_step_activity(1, "生成工作簿失败")
        else:
            self.show_step_panel(3)
            self._set_stepper_state(active_step=3, completed_steps={1, 2})
            self._show_step_activity(3, "生成最终汇总失败")
        self.result_var.set(error_message)
        if payload["traceback"]:
            self._append_log(payload["traceback"])
        messagebox.showerror(f"{stage_name}失败", error_message)
        self._update_finalize_button_state()

    def _render_manual_input_items(self, items):
        for child in self.manual_items_frame.winfo_children():
            child.destroy()
        self.manual_item_vars = []

        if not items:
            if self.prepared_result is None:
                self.manual_hint_var.set("将在生成工作簿后显示需要补充的数据")
            else:
                self.manual_hint_var.set("无需补充数据，系统将直接进入生成汇总步骤")
            return

        self.manual_hint_var.set("请在处理完以下工作表后逐项勾选确认。")
        for item in items:
            variable = tk.BooleanVar(value=False)
            variable.trace_add("write", self._on_manual_confirmation_changed)
            checkbox = ttk.Checkbutton(
                self.manual_items_frame,
                text=item["display_label"],
                variable=variable,
            )
            checkbox.pack(anchor=tk.W, pady=2)
            self.manual_item_vars.append(variable)

    def _on_manual_confirmation_changed(self, *_args):
        if self.manual_input_items:
            if self._can_finalize():
                self.show_step_panel(3)
                self._set_stepper_state(active_step=3, completed_steps={1, 2})
                self.result_var.set(
                    "已完成补充数据确认。\n\n"
                    "现在可以生成最终汇总。"
                )
            else:
                self.show_step_panel(2)
                self._set_stepper_state(active_step=2, completed_steps={1})
        self._update_finalize_button_state()

    def _can_finalize(self):
        if not self.prepared_result:
            return False
        if not self.manual_input_items:
            return True
        return all(variable.get() for variable in self.manual_item_vars)

    def _update_finalize_button_state(self):
        if self.session_completed:
            self._set_button_enabled(self.finalize_button, False)
            return
        if self.worker_thread and self.worker_thread.is_alive():
            self._set_button_enabled(self.finalize_button, False)
            return
        if self._can_finalize():
            self._set_button_enabled(self.finalize_button, True)
        else:
            self._set_button_enabled(self.finalize_button, False)

    def _set_running(self, running):
        if running:
            self._set_button_enabled(self.prepare_button, False)
            self._set_button_enabled(self.finalize_button, False)
            return
        if self.session_completed:
            self._set_button_enabled(self.prepare_button, False)
        elif self.prepared_result is None and self.folder_var.get().strip():
            self._set_button_enabled(self.prepare_button, True)
        else:
            self._set_button_enabled(self.prepare_button, False)
        self._update_finalize_button_state()

    def _append_log(self, message):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _update_status_from_log(self, message):
        if "Starting FX summary workbook preparation..." in message:
            self.show_step_panel(1)
            self._set_stepper_state(active_step=1)
            self._show_step_activity(1, "正在生成工作簿...")
        elif "Phase complete:" in message and self.current_stage == "prepare":
            phase = message.split("Phase complete:", 1)[1].split("elapsed=", 1)[0].strip()
            self._show_step_activity(1, f"正在生成工作簿: {phase}")
        elif "Starting FX consolidation post-processing..." in message:
            self.show_step_panel(3)
            self._set_stepper_state(active_step=3, completed_steps={1, 2})
            self._show_step_activity(3, "正在生成最终汇总...")
        elif self.current_stage == "finalize":
            if "Publishing" in message or "Building" in message or "Rebuilding" in message or "Saving finalized workbook updates" in message:
                self._show_step_activity(3, message.split(" - ", 2)[-1] if " - " in message else message)
        elif "Completed final FX summary report:" in message:
            self._set_stepper_state(completed_steps={1, 2, 3})
            self._clear_step_activity()

    def _format_error_message(self, exc, stage_name, stage_input):
        if isinstance(stage_input, dict):
            stage_path = stage_input.get("workbook_path") or stage_input.get("source_root") or ""
        else:
            stage_path = stage_input
        if stage_name == "第 1 步":
            result_dir = os.path.join(stage_path, "result") if stage_path else ""
        else:
            result_dir = os.path.dirname(stage_path) if stage_path else ""
        log_hint = f"\n\n请查看日志目录：\n{result_dir}" if result_dir else ""
        return f"{stage_name}失败：\n{exc}{log_hint}"

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
