import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

from translator import OfflineTranslator
from ppt_handler import PPTProcessor


class TranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Offline PPT Translator")
        self.root.geometry("900x700")
        self.root.minsize(760, 600)

        self.log_queue = queue.Queue()
        self.is_running = False
        self.worker_thread = None

        self.translator = None
        self.processor = None

        self._build_ui()
        self._load_translator()
        self._poll_log_queue()

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        title_frame = ttk.Frame(main)
        title_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(
            title_frame,
            text="离线 PPT 文本替换翻译器",
            font=("Microsoft YaHei", 16, "bold")
        ).pack(side=tk.LEFT)

        self.status_var = tk.StringVar(value="状态：初始化中...")
        ttk.Label(
            title_frame,
            textvariable=self.status_var,
            foreground="green"
        ).pack(side=tk.RIGHT)

        settings_frame = ttk.LabelFrame(main, text="设置", padding=12)
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        settings_frame.columnconfigure(1, weight=1)

        ttk.Label(settings_frame, text="输入文件夹:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.input_dir_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.input_dir_var).grid(row=0, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(settings_frame, text="浏览...", command=self.select_input_dir).grid(row=0, column=2, padx=5, pady=5)

        ttk.Label(settings_frame, text="输出文件夹:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.output_dir_var = tk.StringVar()
        ttk.Entry(settings_frame, textvariable=self.output_dir_var).grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        ttk.Button(settings_frame, text="浏览...", command=self.select_output_dir).grid(row=1, column=2, padx=5, pady=5)

        ttk.Label(settings_frame, text="源语言:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.source_lang = ttk.Combobox(settings_frame, width=10, state="readonly", values=["en"])
        self.source_lang.set("en")
        self.source_lang.grid(row=2, column=1, sticky="w", padx=5, pady=5)

        ttk.Label(settings_frame, text="目标语言:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.target_lang = ttk.Combobox(settings_frame, width=10, state="readonly", values=["zh"])
        self.target_lang.set("zh")
        self.target_lang.grid(row=3, column=1, sticky="w", padx=5, pady=5)

        action_frame = ttk.Frame(main)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        self.start_btn = ttk.Button(action_frame, text="开始翻译", command=self.start_translation)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(action_frame, text="快速测试", command=self.quick_test).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="扫描文件", command=self.scan_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(action_frame, text="打开输出文件夹", command=self.open_output_dir).pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(action_frame, mode="indeterminate", length=220)
        self.progress.pack(side=tk.RIGHT, padx=5)

        test_frame = ttk.LabelFrame(main, text="快速测试", padding=12)
        test_frame.pack(fill=tk.X, pady=(0, 10))
        test_frame.columnconfigure(0, weight=1)
        test_frame.columnconfigure(2, weight=1)

        self.test_input_var = tk.StringVar(value="Business Plan 2024")
        ttk.Entry(test_frame, textvariable=self.test_input_var).grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        ttk.Button(test_frame, text="翻译 →", command=self.quick_test).grid(row=0, column=1, padx=5, pady=5)

        self.test_output_var = tk.StringVar()
        ttk.Entry(test_frame, textvariable=self.test_output_var, state="readonly").grid(row=0, column=2, sticky="ew", padx=5, pady=5)

        stats_frame = ttk.LabelFrame(main, text="统计", padding=12)
        stats_frame.pack(fill=tk.X, pady=(0, 10))

        self.stats_var = tk.StringVar(value="等待开始...")
        ttk.Label(stats_frame, textvariable=self.stats_var).pack(anchor="w")

        log_frame = ttk.LabelFrame(main, text="日志", padding=12)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _load_translator(self):
        try:
            self.translator = OfflineTranslator()
            self.processor = PPTProcessor(self.translator, self.enqueue_log)

            info = self.translator.get_status_info()

            self.status_var.set("状态：翻译引擎已就绪")
            self.enqueue_log("[INFO] 翻译引擎加载成功")
            self.enqueue_log(f"[INFO] phrases file: {info['phrases_path']}")
            self.enqueue_log(f"[INFO] lexicon file: {info['lexicon_path']}")
            self.enqueue_log(f"[INFO] 短语数: {info['phrases_count']}")
            self.enqueue_log(f"[INFO] 单词数: {info['lexicon_count']}")
            self.enqueue_log(f"[INFO] Argos 可用: {info['argos_available']}")
            self.enqueue_log(f"[INFO] Argos 模型已安装: {info['argos_installed']}")

        except Exception as e:
            self.status_var.set("状态：翻译引擎加载失败")
            self.enqueue_log(f"[ERROR] 加载翻译引擎失败: {e}")
            messagebox.showerror("错误", f"加载翻译引擎失败:\n{e}")

    def enqueue_log(self, message):
        self.log_queue.put(message)

    def _poll_log_queue(self):
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log_queue)

    def select_input_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.input_dir_var.set(directory)
            if not self.output_dir_var.get():
                self.output_dir_var.set(os.path.join(directory, "translated_output"))

    def select_output_dir(self):
        directory = filedialog.askdirectory()
        if directory:
            self.output_dir_var.set(directory)

    def scan_files(self):
        input_dir = self.input_dir_var.get().strip()
        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showwarning("提示", "请先选择有效的输入文件夹")
            return

        pptx_files = list(Path(input_dir).glob("*.pptx"))
        ppt_files = list(Path(input_dir).glob("*.ppt"))

        self.enqueue_log("=" * 60)
        self.enqueue_log(f"[SCAN] 输入目录: {input_dir}")
        self.enqueue_log(f"[SCAN] 找到 .pptx: {len(pptx_files)}")
        self.enqueue_log(f"[SCAN] 找到 .ppt : {len(ppt_files)}")

        for file in pptx_files:
            self.enqueue_log(f"  [PPTX] {file.name}")
        for file in ppt_files:
            self.enqueue_log(f"  [PPT ] {file.name}")

        self.stats_var.set(f"扫描结果：pptx={len(pptx_files)}，ppt={len(ppt_files)}")

    def quick_test(self):
        if not self.translator:
            messagebox.showerror("错误", "翻译引擎未加载")
            return

        text = self.test_input_var.get().strip()
        if not text:
            return

        result = self.translator.translate(text)
        self.test_output_var.set(result)
        self.enqueue_log(f"[TEST] {text} -> {result}")

    def open_output_dir(self):
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showwarning("提示", "请先选择输出文件夹")
            return

        if not os.path.isdir(output_dir):
            messagebox.showwarning("提示", "输出文件夹不存在")
            return

        try:
            os.startfile(output_dir)
        except Exception as e:
            messagebox.showerror("错误", f"打开输出文件夹失败:\n{e}")

    def start_translation(self):
        if self.is_running:
            return

        if not self.processor:
            messagebox.showerror("错误", "处理器未初始化")
            return

        input_dir = self.input_dir_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showwarning("提示", "请选择有效的输入文件夹")
            return

        if not output_dir:
            messagebox.showwarning("提示", "请选择输出文件夹")
            return

        os.makedirs(output_dir, exist_ok=True)

        pptx_files = list(Path(input_dir).glob("*.pptx"))
        ppt_files = list(Path(input_dir).glob("*.ppt"))

        if not pptx_files and not ppt_files:
            messagebox.showwarning("提示", "输入文件夹下未找到 .pptx 或 .ppt 文件")
            return

        msg = (
            f"找到 .pptx 文件 {len(pptx_files)} 个\n"
            f"找到 .ppt 文件 {len(ppt_files)} 个\n\n"
            f"是否开始？"
        )
        if not messagebox.askyesno("确认", msg):
            return

        self.is_running = True
        self.start_btn.config(state="disabled")
        self.progress.start()
        self.log_text.delete("1.0", tk.END)

        self.enqueue_log("=" * 60)
        self.enqueue_log("[INFO] 开始批量翻译")
        self.enqueue_log(f"[INFO] 输入目录: {input_dir}")
        self.enqueue_log(f"[INFO] 输出目录: {output_dir}")
        self.enqueue_log(f"[INFO] 计划处理 .pptx: {len(pptx_files)}")
        self.enqueue_log(f"[INFO] 扫描到 .ppt: {len(ppt_files)} (跳过)")
        self.enqueue_log("=" * 60)

        self.worker_thread = threading.Thread(
            target=self._worker_translate,
            args=(input_dir, output_dir),
            daemon=True
        )
        self.worker_thread.start()

    def _worker_translate(self, input_dir, output_dir):
        try:
            stats = self.processor.translate_directory(input_dir, output_dir)
            self.root.after(0, lambda: self._on_translation_done(stats))
        except Exception as e:
            self.root.after(0, lambda: self._on_translation_error(str(e)))

    def _on_translation_done(self, stats):
        self.is_running = False
        self.start_btn.config(state="normal")
        self.progress.stop()

        summary = (
            f"完成：总文件 {stats['total_files']}，"
            f"成功 {stats['success_files']}，"
            f"跳过 {stats['skipped_files']}，"
            f"失败 {stats['failed_files']}，"
            f"总页数 {stats['total_slides']}，"
            f"翻译文本数 {stats['translated_items']}"
        )
        self.stats_var.set(summary)

        self.enqueue_log("=" * 60)
        self.enqueue_log("[INFO] 处理完成")
        self.enqueue_log(summary)

        if stats["errors"]:
            self.enqueue_log("[INFO] 错误列表：")
            for err in stats["errors"]:
                self.enqueue_log(f"  - {err}")

        messagebox.showinfo("完成", summary)

    def _on_translation_error(self, error_message):
        self.is_running = False
        self.start_btn.config(state="normal")
        self.progress.stop()

        self.enqueue_log(f"[ERROR] 处理失败: {error_message}")
        messagebox.showerror("错误", f"翻译过程中出错:\n{error_message}")


def main():
    root = tk.Tk()
    app = TranslatorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
