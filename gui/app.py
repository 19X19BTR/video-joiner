"""
gui/app.py - 视频排列拼接工具 GUI
"""

import os
import sys
import subprocess
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.constants import (
    DEFAULT_INPUT_ROOT, DEFAULT_OUTPUT_ROOT, DEFAULT_BGM_DIR, WARN_PERM_COUNT, HIDDEN_SI
)
from core.config import load_config, save_config
from core.scanner import scan_videos, probe_video, build_video_lookup
from core.permuter import perm_to_key, perm_to_filename, generate_all_permutations
from core.encoder import (
    concat_videos, scan_bgm_files, measure_audio_loudness,
    calculate_bgm_volume, add_bgm_to_video, BgmManager
)
from core.excel import (
    get_excel_path, load_existing_status, create_excel, update_excel_status
)


class VideoJoinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎬 视频排列拼接工具")
        self.root.geometry("780x680")
        self.root.resizable(True, True)
        self.root.minsize(680, 580)

        self.groups = {}
        self.group_vars = {}  # {group_name: BooleanVar}
        self.is_running = False
        self.stop_flag = False
        self.bgm_volume = 1.0  # BGM音量倍数（预览确认后更新）

        self.config = load_config()
        self.input_root = self.config.get('input_root', DEFAULT_INPUT_ROOT)
        self.output_root = self.config.get('output_root', DEFAULT_OUTPUT_ROOT)

        self.BG = "#f5f5f5"
        self.CARD_BG = "#ffffff"
        self.ACCENT = "#4472C4"
        self.SUCCESS = "#2e7d32"
        self.WARN = "#f57f17"
        self.FAIL = "#c62828"

        self.root.configure(bg=self.BG)

        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('Card.TFrame', background=self.CARD_BG)
        self.style.configure('Title.TLabel', font=('Microsoft YaHei UI', 11, 'bold'),
                             background=self.CARD_BG)
        self.style.configure('Accent.TButton', font=('Microsoft YaHei UI', 10),
                             padding=(12, 6))

        self._build_ui()

    # ─── UI 构建 ───

    def _build_ui(self):
        title_frame = tk.Frame(self.root, bg=self.ACCENT, height=50)
        title_frame.pack(fill='x')
        title_frame.pack_propagate(False)
        tk.Label(title_frame, text="🎬  视频排列拼接工具",
                 font=('Microsoft YaHei UI', 16, 'bold'),
                 fg='white', bg=self.ACCENT).pack(side='left', padx=20, pady=8)
        tk.Button(title_frame, text="⚙ 设置", font=('Microsoft YaHei UI', 10),
                  fg='white', bg=self.ACCENT, relief='flat',
                  activebackground='#3a62a8', activeforeground='white',
                  cursor='hand2', padx=12, pady=2,
                  command=self._open_settings).pack(side='right', padx=20, pady=8)

        main = tk.Frame(self.root, bg=self.BG)
        main.pack(fill='both', expand=True, padx=16, pady=10)

        self._build_folder_card(main)
        self._build_params_card(main)
        self._build_preview_card(main)
        self._build_action_card(main)
        self._build_log_area(main)

    def _build_folder_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame')
        card.pack(fill='x', pady=(0, 8))
        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill='x', padx=16, pady=12)

        ttk.Label(inner, text="📁 文件夹设置", style='Title.TLabel').pack(anchor='w')

        row1 = tk.Frame(inner, bg=self.CARD_BG)
        row1.pack(fill='x', pady=(8, 4))
        tk.Label(row1, text="素材目录：", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, width=10, anchor='e').pack(side='left')
        self.input_var = tk.StringVar(value=self.config.get('last_input', ''))
        tk.Entry(row1, textvariable=self.input_var, font=('Consolas', 10),
                 state='readonly', readonlybackground='#fafafa').pack(
            side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(row1, text="浏览...", style='Accent.TButton',
                   command=self._browse_input).pack(side='right')
        tk.Button(row1, text="📂", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=4, cursor='hand2',
                  command=lambda: self._open_folder(self.input_var.get())).pack(
            side='right', padx=(0, 4))

        tk.Label(inner, text=f"素材根目录: {self.input_root}",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#999').pack(
            anchor='w', padx=(78, 0))

        row2 = tk.Frame(inner, bg=self.CARD_BG)
        row2.pack(fill='x', pady=4)
        tk.Label(row2, text="输出目录：", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, width=10, anchor='e').pack(side='left')
        self.output_var = tk.StringVar(value=self.config.get('last_output', ''))
        tk.Entry(row2, textvariable=self.output_var, font=('Consolas', 10),
                 state='readonly', readonlybackground='#fafafa').pack(
            side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(row2, text="浏览...", style='Accent.TButton',
                   command=self._browse_output).pack(side='right')
        tk.Button(row2, text="📂", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=4, cursor='hand2',
                  command=lambda: self._open_folder(self.output_var.get())).pack(
            side='right', padx=(0, 4))

        tk.Label(inner, text=f"输出根目录: {self.output_root}（自动按日期分文件夹）",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#999').pack(
            anchor='w', padx=(78, 0))

        # ── 组规则 ──
        sep = ttk.Separator(inner, orient='horizontal')
        sep.pack(fill='x', pady=(10, 6))

        rule_top = tk.Frame(inner, bg=self.CARD_BG)
        rule_top.pack(fill='x')
        tk.Label(rule_top, text="组规则：", font=('Microsoft YaHei UI', 10, 'bold'),
                 bg=self.CARD_BG).pack(side='left')
        tk.Label(rule_top, text="（留空则自动按首字母分组 A1.mp4 B1.mp4...）",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#999').pack(
            side='left', padx=(4, 0))

        self.rules_var = tk.StringVar(value=self.config.get('rules', ''))
        rules_entry = tk.Entry(inner, textvariable=self.rules_var,
                               font=('Microsoft YaHei UI', 10))
        rules_entry.pack(fill='x', pady=(4, 0))

        tk.Label(inner, text="格式: A=开头,B=痛点,C=结尾,D=效果  （组名=关键词，逗号分隔）",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#666').pack(
            anchor='w')
        tk.Label(inner, text="文件名包含关键词即归入该组，如 '开头_产品介绍.mp4' → A组",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#999').pack(
            anchor='w')

    def _build_params_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame')
        card.pack(fill='x', pady=(0, 8))
        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill='x', padx=16, pady=12)

        ttk.Label(inner, text="⚙️ 参数设置", style='Title.TLabel').pack(anchor='w')

        row = tk.Frame(inner, bg=self.CARD_BG)
        row.pack(fill='x', pady=(8, 0))

        f1 = tk.LabelFrame(row, text="排列模式", font=('Microsoft YaHei UI', 9),
                           bg=self.CARD_BG, padx=10, pady=6)
        f1.pack(side='left', padx=(0, 16))
        self.mode_var = tk.StringVar(value=self.config.get('mode', 'ordered'))
        tk.Radiobutton(f1, text="有序 (A→B→C→D)", variable=self.mode_var,
                       value='ordered', bg=self.CARD_BG,
                       font=('Microsoft YaHei UI', 9), command=self._on_mode_change).pack(anchor='w')
        tk.Radiobutton(f1, text="无序 (随机组序)", variable=self.mode_var,
                       value='unordered', bg=self.CARD_BG,
                       font=('Microsoft YaHei UI', 9), command=self._on_mode_change).pack(anchor='w')
        tk.Radiobutton(f1, text="部分固定 (自定义)", variable=self.mode_var,
                       value='partial', bg=self.CARD_BG,
                       font=('Microsoft YaHei UI', 9), command=self._on_mode_change).pack(anchor='w')

        f2 = tk.LabelFrame(row, text="拼接数量", font=('Microsoft YaHei UI', 9),
                           bg=self.CARD_BG, padx=10, pady=6)
        f2.pack(side='left', padx=(0, 24))
        self.count_var = tk.StringVar(value=str(self.config.get('count', 100)))
        count_spin = tk.Spinbox(f2, from_=1, to=99999, width=8,
                                textvariable=self.count_var,
                                font=('Microsoft YaHei UI', 10),
                                buttonbackground='#ddd')
        count_spin.pack()
        self.count_hint = tk.Label(f2, text="（表格自动生成全部，此处控制拼接视频数）",
                                   font=('Microsoft YaHei UI', 8),
                                   bg=self.CARD_BG, fg='#888')
        self.count_hint.pack()

        # ── 部分固定设置 ──
        self.partial_frame = tk.LabelFrame(row, text="固定组别设置", font=('Microsoft YaHei UI', 9),
                                           bg=self.CARD_BG, padx=10, pady=6)
        self.partial_frame.pack(side='left', padx=(0, 16))
        self.partial_frame.pack_forget()  # 默认隐藏

        self.partial_fixed_vars = {}  # 各组是否固定的复选框变量
        self.partial_fixed_frame = tk.Frame(self.partial_frame, bg=self.CARD_BG)
        self.partial_fixed_frame.pack()
        tk.Label(self.partial_fixed_frame, text="选择要固定的组：", font=('Microsoft YaHei UI', 9),
                 bg=self.CARD_BG, fg='#666').pack(side='left')

        f3 = tk.LabelFrame(row, text="输出分辨率", font=('Microsoft YaHei UI', 9),
                           bg=self.CARD_BG, padx=10, pady=6)
        f3.pack(side='left', padx=(0, 16))
        self.res_var = tk.StringVar(
            value=self.config.get('resolution', '自动 (以第一个视频为准)'))
        ttk.Combobox(f3, textvariable=self.res_var,
                     values=['自动 (以第一个视频为准)', '1920x1080', '1280x720',
                             '854x480', '3840x2160'],
                     state='readonly', width=24,
                     font=('Microsoft YaHei UI', 9)).pack()

        # ── BGM 设置 ──
        f4 = tk.LabelFrame(row, text="背景音乐", font=('Microsoft YaHei UI', 9),
                           bg=self.CARD_BG, padx=10, pady=6)
        f4.pack(side='left')
        self.use_bgm_var = tk.BooleanVar(value=self.config.get('use_bgm', False))
        tk.Checkbutton(f4, text="添加BGM", variable=self.use_bgm_var,
                       font=('Microsoft YaHei UI', 9), bg=self.CARD_BG,
                       command=self._on_bgm_toggle).pack(anchor='w')
        self.bgm_info_label = tk.Label(f4, text="未选择", font=('Microsoft YaHei UI', 8),
                                       bg=self.CARD_BG, fg='#999')
        self.bgm_info_label.pack(anchor='w')
        self.bgm_select_btn = tk.Button(f4, text="选择BGM...", font=('Microsoft YaHei UI', 8),
                                        relief='flat', bg='#e8eaf6', padx=6, cursor='hand2',
                                        command=self._open_bgm_picker)
        self.bgm_select_btn.pack(anchor='w', pady=(2, 0))

        # BGM状态
        self.selected_bgm_files = []  # 用户选中的BGM文件列表
        self.bgm_dir = self.config.get('bgm_dir', DEFAULT_BGM_DIR)

    def _build_preview_card(self, parent):
        card = ttk.Frame(parent, style='Card.TFrame')
        card.pack(fill='x', pady=(0, 8))
        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill='x', padx=16, pady=12)

        top_row = tk.Frame(inner, bg=self.CARD_BG)
        top_row.pack(fill='x')
        ttk.Label(top_row, text="📋 视频分组预览", style='Title.TLabel').pack(side='left')
        ttk.Button(top_row, text="🔄 重新扫描", style='Accent.TButton',
                   command=self._scan_folder).pack(side='right')

        self.preview_frame = tk.Frame(inner, bg=self.CARD_BG)
        self.preview_frame.pack(fill='x', pady=(8, 0))

        tk.Label(self.preview_frame, text="请先选择素材目录",
                 font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, fg='#999').pack()

    def _build_action_card(self, parent):
        card = tk.Frame(parent, bg=self.CARD_BG, relief='flat',
                        highlightthickness=1, highlightbackground='#ddd')
        card.pack(fill='x', pady=(0, 8))
        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill='x', padx=16, pady=12)

        btn_row = tk.Frame(inner, bg=self.CARD_BG)
        btn_row.pack(fill='x')

        self.start_btn = tk.Button(
            btn_row, text="▶  开始拼接", font=('Microsoft YaHei UI', 12, 'bold'),
            bg=self.ACCENT, fg='white', activebackground='#3a62a8',
            activeforeground='white', relief='flat', padx=30, pady=8,
            cursor='hand2', command=self._start_process)
        self.start_btn.pack(side='left', padx=(0, 12))

        self.excel_btn = tk.Button(
            btn_row, text="📊 仅生成表格", font=('Microsoft YaHei UI', 10),
            bg='#e8eaf6', fg='#333', activebackground='#c5cae9',
            relief='flat', padx=16, pady=8, cursor='hand2',
            command=self._generate_excel_only)
        self.excel_btn.pack(side='left', padx=(0, 12))

        self.stop_btn = tk.Button(
            btn_row, text="⏹ 停止", font=('Microsoft YaHei UI', 10),
            bg='#ffcdd2', fg='#c62828', activebackground='#ef9a9a',
            relief='flat', padx=16, pady=8, cursor='hand2',
            state='disabled', command=self._stop_process)
        self.stop_btn.pack(side='left')

        prog_frame = tk.Frame(inner, bg=self.CARD_BG)
        prog_frame.pack(fill='x', pady=(10, 0))

        self.progress_var = tk.DoubleVar(value=0)
        ttk.Progressbar(prog_frame, variable=self.progress_var,
                        maximum=100, length=400).pack(
            fill='x', side='left', expand=True, padx=(0, 10))

        self.progress_label = tk.Label(prog_frame, text="0/0 (0%)",
                                       font=('Consolas', 10), bg=self.CARD_BG, fg='#555')
        self.progress_label.pack(side='right')

        self.result_label = tk.Label(inner, text="", font=('Microsoft YaHei UI', 10),
                                     bg=self.CARD_BG)
        self.result_label.pack(anchor='w', pady=(6, 0))

    def _build_log_area(self, parent):
        log_frame = tk.Frame(parent, bg=self.BG)
        log_frame.pack(fill='both', expand=True)

        tk.Label(log_frame, text="📝 运行日志", font=('Microsoft YaHei UI', 9, 'bold'),
                 bg=self.BG, fg='#666').pack(anchor='w')

        self.log_text = tk.Text(log_frame, height=8, font=('Consolas', 9),
                                bg='#1e1e1e', fg='#d4d4d4',
                                insertbackground='white', relief='flat',
                                wrap='word', state='disabled')
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        self.log_text.pack(fill='both', expand=True, pady=(4, 0))

        self.log_text.tag_configure('info', foreground='#d4d4d4')
        self.log_text.tag_configure('success', foreground='#6a9955')
        self.log_text.tag_configure('warn', foreground='#dcdcaa')
        self.log_text.tag_configure('error', foreground='#f44747')
        self.log_text.tag_configure('title', foreground='#569cd6',
                                    font=('Consolas', 9, 'bold'))

    # ─── 日志 ───

    def _log(self, msg, tag='info'):
        self.log_text.configure(state='normal')
        self.log_text.insert('end', msg + '\n', tag)
        self.log_text.see('end')
        self.log_text.configure(state='disabled')

    def _log_clear(self):
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', 'end')
        self.log_text.configure(state='disabled')

    # ─── 设置 ───

    def _open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("⚙ 设置")
        win.geometry("560x260")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.configure(bg=self.CARD_BG)

        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 560) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 220) // 2
        win.geometry(f"560x260+{x}+{y}")

        tk.Label(win, text="⚙  路径设置", font=('Microsoft YaHei UI', 14, 'bold'),
                 bg=self.CARD_BG, fg=self.ACCENT).pack(pady=(16, 12))

        row1 = tk.Frame(win, bg=self.CARD_BG)
        row1.pack(fill='x', padx=24, pady=6)
        tk.Label(row1, text="素材根目录：", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, width=12, anchor='e').pack(side='left')
        input_var = tk.StringVar(value=self.input_root)
        tk.Entry(row1, textvariable=input_var, font=('Consolas', 10),
                 width=42).pack(side='left', padx=(0, 8))
        ttk.Button(row1, text="浏览",
                   command=lambda: self._settings_browse(input_var)).pack(side='right')
        tk.Button(row1, text="📂", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=4, cursor='hand2',
                  command=lambda: self._open_folder(input_var.get())).pack(
            side='right', padx=(0, 4))

        row2 = tk.Frame(win, bg=self.CARD_BG)
        row2.pack(fill='x', padx=24, pady=6)
        tk.Label(row2, text="输出根目录：", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, width=12, anchor='e').pack(side='left')
        output_var = tk.StringVar(value=self.output_root)
        tk.Entry(row2, textvariable=output_var, font=('Consolas', 10),
                 width=42).pack(side='left', padx=(0, 8))
        ttk.Button(row2, text="浏览",
                   command=lambda: self._settings_browse(output_var)).pack(side='right')
        tk.Button(row2, text="📂", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=4, cursor='hand2',
                  command=lambda: self._open_folder(output_var.get())).pack(
            side='right', padx=(0, 4))

        row3 = tk.Frame(win, bg=self.CARD_BG)
        row3.pack(fill='x', padx=24, pady=6)
        tk.Label(row3, text="BGM 目录：", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG, width=12, anchor='e').pack(side='left')
        bgm_var = tk.StringVar(value=self.bgm_dir)
        tk.Entry(row3, textvariable=bgm_var, font=('Consolas', 10),
                 width=42).pack(side='left', padx=(0, 8))
        ttk.Button(row3, text="浏览",
                   command=lambda: self._settings_browse(bgm_var)).pack(side='right')
        tk.Button(row3, text="📂", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=4, cursor='hand2',
                  command=lambda: self._open_folder(bgm_var.get())).pack(
            side='right', padx=(0, 4))

        btn_frame = tk.Frame(win, bg=self.CARD_BG)
        btn_frame.pack(fill='x', padx=24, pady=(12, 0))

        def save_and_close():
            self.input_root = input_var.get()
            self.output_root = output_var.get()
            self.bgm_dir = bgm_var.get()
            self.config['input_root'] = self.input_root
            self.config['output_root'] = self.output_root
            self.config['bgm_dir'] = self.bgm_dir
            save_config(self.config)
            win.destroy()
            self._log(f"设置已保存: 素材={self.input_root}, 输出={self.output_root}, BGM={self.bgm_dir}", 'title')

        tk.Button(btn_frame, text="保存", font=('Microsoft YaHei UI', 10, 'bold'),
                  bg=self.ACCENT, fg='white', relief='flat', padx=24, pady=4,
                  cursor='hand2', command=save_and_close).pack(side='right')
        tk.Button(btn_frame, text="取消", font=('Microsoft YaHei UI', 10),
                  bg='#e0e0e0', fg='#333', relief='flat', padx=16, pady=4,
                  cursor='hand2', command=win.destroy).pack(side='right', padx=(0, 8))

    def _open_folder(self, path):
        """打开文件夹，不存在则提示。"""
        if path and os.path.isdir(path):
            os.startfile(path)
        else:
            messagebox.showwarning("提示", "文件夹不存在或未设置！")

    def _settings_browse(self, var):
        path = filedialog.askdirectory(title="选择目录")
        if path:
            var.set(path)

    # ─── BGM ───

    def _on_bgm_toggle(self):
        """BGM开关切换。"""
        if self.use_bgm_var.get():
            if not self.selected_bgm_files:
                all_bgm = scan_bgm_files(self.bgm_dir)
                if all_bgm:
                    self.selected_bgm_files = all_bgm
                    self.bgm_info_label.config(
                        text=f"已加载 {len(all_bgm)} 首（目录）", fg='#2e7d32')
                else:
                    self.bgm_info_label.config(text="⚠️ BGM目录为空", fg='#e65100')
        else:
            self.bgm_info_label.config(text="未选择", fg='#999')
        self.config['use_bgm'] = self.use_bgm_var.get()
        save_config(self.config)

    def _open_bgm_picker(self):
        """打开BGM选择弹窗。"""
        all_bgm = scan_bgm_files(self.bgm_dir)
        if not all_bgm:
            messagebox.showwarning("提示", f"BGM目录为空或不存在：\n{self.bgm_dir}")
            return

        win = tk.Toplevel(self.root)
        win.title("🎵 选择背景音乐")
        win.geometry("500x420")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.configure(bg=self.CARD_BG)

        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 500) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 420) // 2
        win.geometry(f"500x420+{x}+{y}")

        tk.Label(win, text="🎵  选择背景音乐", font=('Microsoft YaHei UI', 14, 'bold'),
                 bg=self.CARD_BG, fg=self.ACCENT).pack(pady=(16, 8))
        tk.Label(win, text=f"目录: {self.bgm_dir}", font=('Microsoft YaHei UI', 8),
                 bg=self.CARD_BG, fg='#999').pack()

        btn_row = tk.Frame(win, bg=self.CARD_BG)
        btn_row.pack(fill='x', padx=24, pady=(12, 4))

        bgm_vars = []

        def select_all():
            for var, _ in bgm_vars:
                var.set(True)

        def deselect_all():
            for var, _ in bgm_vars:
                var.set(False)

        tk.Button(btn_row, text="全选", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=8, cursor='hand2',
                  command=select_all).pack(side='left')
        tk.Button(btn_row, text="全不选", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=8, cursor='hand2',
                  command=deselect_all).pack(side='left', padx=(8, 0))
        tk.Label(btn_row, text=f"共 {len(all_bgm)} 首",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#666').pack(side='right')

        list_frame = tk.Frame(win, bg=self.CARD_BG)
        list_frame.pack(fill='both', expand=True, padx=24, pady=(4, 8))

        canvas = tk.Canvas(list_frame, bg=self.CARD_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=self.CARD_BG)

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        selected_names = {f['name'] for f in self.selected_bgm_files}

        for bgm in all_bgm:
            var = tk.BooleanVar(value=(bgm['name'] in selected_names or len(selected_names) == 0))
            bgm_vars.append((var, bgm))
            tk.Checkbutton(scrollable, text=bgm['name'], variable=var,
                           font=('Microsoft YaHei UI', 9), bg=self.CARD_BG,
                           anchor='w').pack(fill='x', padx=8)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def confirm():
            self.selected_bgm_files = [bgm for var, bgm in bgm_vars if var.get()]
            count = len(self.selected_bgm_files)
            if count > 0:
                self.bgm_info_label.config(
                    text=f"已选 {count} 首", fg='#2e7d32')
                self.use_bgm_var.set(True)
            else:
                self.bgm_info_label.config(text="未选择", fg='#999')
                self.use_bgm_var.set(False)
            self.config['use_bgm'] = self.use_bgm_var.get()
            save_config(self.config)
            win.destroy()

        tk.Button(win, text="确认选择", font=('Microsoft YaHei UI', 10, 'bold'),
                  bg=self.ACCENT, fg='white', relief='flat', padx=24, pady=6,
                  cursor='hand2', command=confirm).pack(pady=(0, 16))

    def _show_bgm_preview(self, video_path: str, bgm_path: str, bgm_name: str,
                          video_duration: float, callback):
        """BGM音量预览窗口。"""
        preview_dur = min(8.0, video_duration)

        win = tk.Toplevel(self.root)
        win.title("🎵 BGM 音量预览")
        win.geometry("560x480")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        win.configure(bg=self.CARD_BG)

        win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 560) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 480) // 2
        win.geometry(f"560x480+{x}+{y}")

        tk.Label(win, text="🎵  BGM 音量预览", font=('Microsoft YaHei UI', 14, 'bold'),
                 bg=self.CARD_BG, fg=self.ACCENT).pack(pady=(16, 8))

        analysis_frame = tk.LabelFrame(win, text="📊 音量分析", font=('Microsoft YaHei UI', 10),
                                       bg=self.CARD_BG, padx=16, pady=10)
        analysis_frame.pack(fill='x', padx=24, pady=(4, 8))

        video_lufs_var = tk.StringVar(value="分析中...")
        bgm_lufs_var = tk.StringVar(value="分析中...")
        recommend_var = tk.StringVar(value="计算中...")

        tk.Label(analysis_frame, text="原视频响度:", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG).grid(row=0, column=0, sticky='w', pady=2)
        tk.Label(analysis_frame, textvariable=video_lufs_var, font=('Consolas', 10, 'bold'),
                 bg=self.CARD_BG, fg='#1565c0').grid(row=0, column=1, sticky='w', padx=(8, 0), pady=2)

        tk.Label(analysis_frame, text="BGM 响度:", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG).grid(row=1, column=0, sticky='w', pady=2)
        tk.Label(analysis_frame, textvariable=bgm_lufs_var, font=('Consolas', 10, 'bold'),
                 bg=self.CARD_BG, fg='#e65100').grid(row=1, column=1, sticky='w', padx=(8, 0), pady=2)

        tk.Label(analysis_frame, text="推荐音量:", font=('Microsoft YaHei UI', 10),
                 bg=self.CARD_BG).grid(row=2, column=0, sticky='w', pady=2)
        tk.Label(analysis_frame, textvariable=recommend_var, font=('Consolas', 10, 'bold'),
                 bg=self.CARD_BG, fg='#2e7d32').grid(row=2, column=1, sticky='w', padx=(8, 0), pady=2)

        slider_frame = tk.LabelFrame(win, text="🎛️ 音量调节", font=('Microsoft YaHei UI', 10),
                                     bg=self.CARD_BG, padx=16, pady=10)
        slider_frame.pack(fill='x', padx=24, pady=(0, 8))

        vol_var = tk.DoubleVar(value=1.0)
        vol_label_var = tk.StringVar(value="1.00")

        vol_row = tk.Frame(slider_frame, bg=self.CARD_BG)
        vol_row.pack(fill='x')

        tk.Label(vol_row, text="0.3", font=('Consolas', 8), bg=self.CARD_BG, fg='#999').pack(side='left')
        vol_slider = tk.Scale(vol_row, from_=0.3, to=5.0, resolution=0.05,
                             orient='horizontal', variable=vol_var,
                             font=('Microsoft YaHei UI', 8), bg=self.CARD_BG,
                             troughcolor='#e8eaf6', length=300,
                             showvalue=False,
                             command=lambda v: vol_label_var.set(f"{float(v):.2f}"))
        vol_slider.pack(side='left', fill='x', expand=True, padx=8)
        tk.Label(vol_row, text="5.0", font=('Consolas', 8), bg=self.CARD_BG, fg='#999').pack(side='left')

        vol_display = tk.Label(slider_frame, textvariable=vol_label_var,
                               font=('Consolas', 14, 'bold'), bg=self.CARD_BG, fg=self.ACCENT)
        vol_display.pack()

        auto_vol = [1.0]
        def reset_auto():
            vol_var.set(auto_vol[0])
            vol_label_var.set(f"{auto_vol[0]:.2f}")

        tk.Button(slider_frame, text="↩ 恢复自动推荐值", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=8, cursor='hand2',
                  command=reset_auto).pack(pady=(6, 0))

        preview_btn_frame = tk.Frame(win, bg=self.CARD_BG)
        preview_btn_frame.pack(fill='x', padx=24, pady=(0, 8))

        preview_status = tk.Label(preview_btn_frame, text="",
                                  font=('Microsoft YaHei UI', 9),
                                  bg=self.CARD_BG, fg='#666')
        preview_status.pack()

        last_preview = [None]

        def do_preview():
            preview_status.config(text="⏳ 正在生成预览...", fg='#e65100')
            preview_btn.config(state='disabled')

            def _generate():
                import tempfile
                vol = vol_var.get()
                tmp_dir = tempfile.gettempdir()
                tmp_path = os.path.join(tmp_dir, "bgm_preview.mp4")

                if last_preview[0] and os.path.exists(last_preview[0]):
                    try:
                        os.remove(last_preview[0])
                    except OSError:
                        pass

                ok = add_bgm_to_video(
                    video_path, bgm_path, tmp_path,
                    video_duration=preview_dur, bgm_volume=vol)

                if ok and os.path.exists(tmp_path):
                    last_preview[0] = tmp_path
                    self.root.after(0, lambda: preview_status.config(
                        text=f"✅ 预览已生成（{preview_dur:.0f}秒），正在打开播放...", fg='#2e7d32'))
                    try:
                        os.startfile(tmp_path)
                    except Exception:
                        subprocess.Popen(['start', '', tmp_path], shell=True)
                    self.root.after(0, lambda: preview_btn.config(state='normal'))
                else:
                    self.root.after(0, lambda: preview_status.config(
                        text="❌ 预览生成失败", fg='#c62828'))
                    self.root.after(0, lambda: preview_btn.config(state='normal'))

            threading.Thread(target=_generate, daemon=True).start()

        preview_btn = tk.Button(preview_btn_frame, text="🔊  生成试听预览",
                                font=('Microsoft YaHei UI', 11, 'bold'),
                                bg='#e8eaf6', fg='#333', relief='flat',
                                padx=24, pady=6, cursor='hand2',
                                command=do_preview)
        preview_btn.pack(pady=(6, 0))

        tk.Label(preview_btn_frame, text=f"预览时长: 前 {preview_dur:.0f} 秒 | BGM: {bgm_name}",
                 font=('Microsoft YaHei UI', 8), bg=self.CARD_BG, fg='#999').pack(pady=(4, 0))

        def _analyze():
            v_info = measure_audio_loudness(video_path, duration_limit=30)
            b_info = measure_audio_loudness(bgm_path, duration_limit=30)

            if v_info and b_info:
                v_db = v_info['mean_volume']
                b_db = b_info['mean_volume']

                self.root.after(0, lambda: video_lufs_var.set(f"{v_db:.1f} dB"))
                self.root.after(0, lambda: bgm_lufs_var.set(f"{b_db:.1f} dB"))
                self.root.after(0, lambda: recommend_var.set("1.00  (原始混音已平衡)"))
            else:
                self.root.after(0, lambda: video_lufs_var.set("分析失败"))
                self.root.after(0, lambda: bgm_lufs_var.set("分析失败"))
                self.root.after(0, lambda: recommend_var.set("1.00  (默认值)"))

        threading.Thread(target=_analyze, daemon=True).start()

        btn_frame = tk.Frame(win, bg=self.CARD_BG)
        btn_frame.pack(fill='x', padx=24, pady=(8, 16))

        def on_confirm():
            win.destroy()
            callback('confirm', vol_var.get())

        def on_skip_bgm():
            win.destroy()
            callback('skip', 0)

        tk.Button(btn_frame, text="✅ 满意，开始批量拼接",
                  font=('Microsoft YaHei UI', 10, 'bold'),
                  bg=self.ACCENT, fg='white', relief='flat', padx=20, pady=6,
                  cursor='hand2', command=on_confirm).pack(side='left')

        tk.Button(btn_frame, text="🔄 重听", font=('Microsoft YaHei UI', 10),
                  bg='#e8eaf6', fg='#333', relief='flat', padx=16, pady=6,
                  cursor='hand2', command=do_preview).pack(side='left', padx=(12, 0))

        tk.Button(btn_frame, text="❌ 不加BGM", font=('Microsoft YaHei UI', 10),
                  bg='#ffcdd2', fg='#c62828', relief='flat', padx=16, pady=6,
                  cursor='hand2', command=on_skip_bgm).pack(side='right')

        def on_close():
            win.destroy()
            callback('cancel', 0)
        win.protocol("WM_DELETE_WINDOW", on_close)

        self.root.wait_window(win)

    # ─── 文件夹浏览 ───

    def _browse_input(self):
        path = filedialog.askdirectory(title="选择素材文件夹", initialdir=self.input_root)
        if path:
            self.input_var.set(path)
            folder_name = os.path.basename(path)
            self.output_var.set(os.path.join(self.output_root, folder_name))
            self._scan_folder()

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出文件夹", initialdir=self.output_root)
        if path:
            self.output_var.set(path)

    # ─── 规则解析 ───

    def _parse_rules(self) -> dict:
        text = self.rules_var.get().strip()
        if not text:
            return {}
        rules = {}
        text = text.replace('，', ',')
        for part in text.split(','):
            part = part.strip()
            if '=' in part:
                name, keyword = part.split('=', 1)
                name = name.strip().upper()
                keyword = keyword.strip()
                if name and keyword:
                    rules[name] = keyword
        return rules

    # ─── 扫描 ───

    def _scan_folder(self):
        input_dir = self.input_var.get()
        if not input_dir:
            messagebox.showwarning("提示", "请先选择素材目录！")
            return

        for widget in self.preview_frame.winfo_children():
            widget.destroy()

        try:
            rules = self._parse_rules()
            self.groups = scan_videos(input_dir, rules=rules if rules else None)
        except FileNotFoundError:
            messagebox.showerror("错误", "素材目录不存在！")
            return

        if not self.groups:
            rules = self._parse_rules()
            hint = "未找到包含关键词的视频" if rules else "未找到符合命名规则的视频 (A1.mp4, B1.mp4, ...)"
            tk.Label(self.preview_frame, text=f"⚠️ {hint}",
                     font=('Microsoft YaHei UI', 10),
                     bg=self.CARD_BG, fg='#e65100').pack()
            return

        group_keys = sorted(self.groups.keys())
        total_videos = sum(len(v) for v in self.groups.values())
        max_perms = 1
        for v in self.groups.values():
            max_perms *= len(v)

        cards_frame = tk.Frame(self.preview_frame, bg=self.CARD_BG)
        cards_frame.pack(fill='x')
        colors = ['#1565c0', '#2e7d32', '#e65100', '#6a1b9a', '#00695c',
                  '#c62828', '#37474f', '#ad1457']
        for i, key in enumerate(group_keys):
            videos = self.groups[key]
            color = colors[i % len(colors)]
            cf = tk.Frame(cards_frame, bg=color, padx=10, pady=6)
            cf.pack(side='left', padx=(0, 6), pady=2)
            tk.Label(cf, text=f" {key} 组 ",
                     font=('Microsoft YaHei UI', 10, 'bold'),
                     fg='white', bg=color).pack()
            names = ', '.join(v['name'] for v in videos)
            tk.Label(cf, text=names, font=('Microsoft YaHei UI', 8),
                     fg='#e0e0e0', bg=color).pack()

        select_frame = tk.Frame(self.preview_frame, bg=self.CARD_BG)
        select_frame.pack(fill='x', pady=(8, 0))
        tk.Label(select_frame, text="选择参与排列的组别：",
                 font=('Microsoft YaHei UI', 10, 'bold'),
                 bg=self.CARD_BG, fg='#333').pack(side='left')

        self.group_vars = {}
        for i, key in enumerate(group_keys):
            var = tk.BooleanVar(value=True)
            self.group_vars[key] = var
            color = colors[i % len(colors)]
            cb = tk.Checkbutton(select_frame, text=f"{key}组", variable=var,
                                font=('Microsoft YaHei UI', 10, 'bold'),
                                bg=self.CARD_BG, fg=color, selectcolor='#e8eaf6',
                                activebackground=self.CARD_BG,
                                command=self._on_group_selection_change)
            cb.pack(side='left', padx=(12, 0))
            if key not in self.partial_fixed_vars:
                self.partial_fixed_vars[key] = tk.BooleanVar(value=self.config.get(f'fixed_{key}', False))

        stale_keys = [k for k in self.partial_fixed_vars if k not in group_keys]
        for k in stale_keys:
            del self.partial_fixed_vars[k]

        btn_frame = tk.Frame(select_frame, bg=self.CARD_BG)
        btn_frame.pack(side='left', padx=(16, 0))
        tk.Button(btn_frame, text="全选", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=6, cursor='hand2',
                  command=lambda: self._toggle_all_groups(True)).pack(side='left')
        tk.Button(btn_frame, text="全不选", font=('Microsoft YaHei UI', 8),
                  relief='flat', bg='#e8eaf6', padx=6, cursor='hand2',
                  command=lambda: self._toggle_all_groups(False)).pack(side='left', padx=(4, 0))

        self.perm_count_label = tk.Label(self.preview_frame, text="",
                                         font=('Microsoft YaHei UI', 10, 'bold'),
                                         bg=self.CARD_BG, fg='#333')
        self.perm_count_label.pack(anchor='w', pady=(6, 0))

        self._update_perm_count()
        self._show_existing_info(input_dir)

        if self.mode_var.get() == 'partial':
            self._update_partial_fixed_ui()

        self._log(f"扫描完成: {len(group_keys)} 组, {total_videos} 个视频, "
                  f"排列 {max_perms:,} 种", 'title')

    def _on_group_selection_change(self):
        selected_keys = set(self._get_selected_groups().keys())
        for key, var in self.partial_fixed_vars.items():
            if key not in selected_keys and var.get():
                var.set(False)
        self._update_perm_count()
        if self.mode_var.get() == 'partial':
            self._update_partial_fixed_ui()

    def _toggle_all_groups(self, value: bool):
        for var in self.group_vars.values():
            var.set(value)
        self._on_group_selection_change()

    def _on_mode_change(self):
        if self.mode_var.get() == 'partial':
            self.partial_frame.pack(side='left', padx=(0, 16))
            self._update_partial_fixed_ui()
        else:
            self.partial_frame.pack_forget()

    def _update_partial_fixed_ui(self):
        for widget in self.partial_fixed_frame.winfo_children():
            widget.destroy()

        tk.Label(self.partial_fixed_frame, text="固定组：", font=('Microsoft YaHei UI', 9),
                 bg=self.CARD_BG, fg='#666').pack(side='left')

        if not self.groups:
            return

        selected_keys = set(self._get_selected_groups().keys())
        colors = ['#1565c0', '#2e7d32', '#e65100', '#6a1b9a', '#00695c',
                  '#c62828', '#37474f', '#ad1457']

        for i, key in enumerate(sorted(self.groups.keys())):
            if key not in selected_keys:
                continue
            if key not in self.partial_fixed_vars:
                self.partial_fixed_vars[key] = tk.BooleanVar(value=False)
            color = colors[i % len(colors)]
            cb = tk.Checkbutton(self.partial_fixed_frame, text=key, variable=self.partial_fixed_vars[key],
                                font=('Microsoft YaHei UI', 9, 'bold'),
                                bg=self.CARD_BG, fg=color, selectcolor='#e8eaf6',
                                activebackground=self.CARD_BG,
                                command=self._update_perm_count)
            cb.pack(side='left', padx=(8, 0))

    def _get_selected_groups(self) -> dict:
        return {k: v for k, v in self.groups.items()
                if k in self.group_vars and self.group_vars[k].get()}

    def _update_perm_count(self):
        selected = self._get_selected_groups()
        if not selected:
            self.perm_count_label.config(text="⚠️ 请至少选择一组", fg='#e65100')
            return
        max_perms = 1
        for v in selected.values():
            max_perms *= len(v)
        if max_perms < 2:
            self.perm_count_label.config(
                text=f"选中 {len(selected)} 组 · 排列数: {max_perms}（至少需要2组）",
                fg='#e65100')
        else:
            self.perm_count_label.config(
                text=f"选中 {len(selected)} 组 · 所有排列共 {max_perms:,} 种",
                fg='#333')

    def _show_existing_info(self, input_dir: str):
        folder_name = os.path.basename(input_dir)
        output_dir = os.path.join(self.output_root, folder_name)
        info_parts = []
        if os.path.exists(output_dir):
            for f in sorted(Path(output_dir).glob('*.xlsx')):
                status = load_existing_status(str(f))
                if status:
                    done = sum(1 for v in status.values() if '✅' in v[0])
                    info_parts.append(f"{f.stem}: {len(status)}条 ({done}✅)")

        if info_parts:
            tk.Label(self.preview_frame,
                     text="📊 已有表格: " + " | ".join(info_parts),
                     font=('Microsoft YaHei UI', 10),
                     bg=self.CARD_BG, fg='#1565c0').pack(anchor='w', pady=(4, 0))

    # ─── 参数解析 ───

    def _get_resolution(self):
        res_str = self.res_var.get()
        if '自动' in res_str:
            return None, None
        try:
            return tuple(map(int, res_str.split('x')))
        except ValueError:
            return None, None

    def _get_params(self):
        input_dir = self.input_var.get()
        output_dir = self.output_var.get()

        if not input_dir:
            messagebox.showwarning("提示", "请选择素材目录！")
            return None
        if not self.groups:
            messagebox.showwarning("提示", "未扫描到视频，请检查素材目录！")
            return None

        selected_groups = self._get_selected_groups()
        if not selected_groups:
            messagebox.showwarning("提示", "请至少选择一组视频！")
            return None
        if len(selected_groups) < 2:
            messagebox.showwarning("提示", "至少需要选择两组才能排列！")
            return None

        if not output_dir:
            folder_name = os.path.basename(input_dir)
            output_dir = os.path.join(self.output_root, folder_name)
            self.output_var.set(output_dir)

        try:
            count = int(self.count_var.get())
        except ValueError:
            messagebox.showwarning("提示", "拼接数量必须是数字！")
            return None
        if count < 1:
            messagebox.showwarning("提示", "拼接数量必须 >= 1！")
            return None

        mode = self.mode_var.get()
        ordered = (mode == 'ordered')
        fixed_groups = []
        if mode == 'partial':
            fixed_groups = sorted([k for k, v in self.partial_fixed_vars.items() if v.get()])
            if not fixed_groups:
                messagebox.showwarning("提示", "部分固定模式下请至少选择一组固定！")
                return None
            if len(fixed_groups) >= len(selected_groups):
                messagebox.showwarning("提示", "固定组不能等于或超过总组数！\n至少留一组可变。")
                return None

        res_w, res_h = self._get_resolution()

        os.makedirs(output_dir, exist_ok=True)

        use_bgm = self.use_bgm_var.get()
        bgm_files = self.selected_bgm_files if use_bgm else []
        if use_bgm and not bgm_files:
            bgm_files = scan_bgm_files(self.bgm_dir)
            if not bgm_files:
                messagebox.showwarning("提示", "已勾选BGM但未选择音乐文件！")
                return None

        self.config['last_input'] = input_dir
        self.config['last_output'] = output_dir
        self.config['mode'] = self.mode_var.get()
        self.config['count'] = count
        self.config['resolution'] = self.res_var.get()
        self.config['rules'] = self.rules_var.get()
        self.config['use_bgm'] = use_bgm
        for key, var in self.partial_fixed_vars.items():
            self.config[f'fixed_{key}'] = var.get()
        save_config(self.config)

        return selected_groups, ordered, fixed_groups, count, res_w, res_h, output_dir, bgm_files

    # ─── 仅生成表格 ───

    def _generate_excel_only(self):
        params = self._get_params()
        if not params:
            return

        groups, ordered, fixed_groups, _, _, _, output_dir, _ = params
        date_str = os.path.basename(self.input_var.get())
        selected_keys = list(groups.keys())
        excel_suffix = f"固定{''.join(fixed_groups)}" if fixed_groups else None
        excel_path = get_excel_path(output_dir, date_str, ordered, selected_keys, excel_suffix)

        existing = load_existing_status(excel_path)
        if existing:
            done_count = sum(1 for v in existing.values() if '✅' in v[0])
            answer = messagebox.askyesno(
                "表格已存在",
                f"已存在表格: {os.path.basename(excel_path)}\n"
                f"共 {len(existing)} 条排列 ({done_count} 条已完成)\n\n"
                f"重新生成会覆盖现有表格，确认？")
            if not answer:
                return

        self._log_clear()
        self._log("📊 仅生成表格模式", 'title')

        perms = generate_all_permutations(groups, ordered, fixed_groups)

        if len(perms) > WARN_PERM_COUNT:
            answer = messagebox.askyesno(
                "确认",
                f"将生成 {len(perms):,} 种排列组合，数量较大。\n\n确认继续？")
            if not answer:
                return

        create_excel(perms, groups, excel_path)

        self._log(f"✅ 表格已生成: {excel_path}", 'success')
        self._log(f"   共 {len(perms):,} 种排列", 'success')
        messagebox.showinfo("完成",
                            f"✅ 表格已生成！\n\n"
                            f"排列数: {len(perms):,}\n"
                            f"文件: {excel_path}")

    # ─── 开始拼接 ───

    def _start_process(self):
        params = self._get_params()
        if not params:
            return

        groups, ordered, fixed_groups, count, res_w, res_h, output_dir, bgm_files = params
        date_str = os.path.basename(self.input_var.get())
        selected_keys = list(groups.keys())
        excel_suffix = f"固定{''.join(fixed_groups)}" if fixed_groups else None
        excel_path = get_excel_path(output_dir, date_str, ordered, selected_keys, excel_suffix)

        existing = load_existing_status(excel_path)
        if not existing:
            max_perms = 1
            for v in groups.values():
                max_perms *= len(v)
            if max_perms > WARN_PERM_COUNT:
                answer = messagebox.askyesno(
                    "确认",
                    f"将生成 {max_perms:,} 种排列组合，数量较大。\n\n确认继续？")
                if not answer:
                    return

        self.is_running = True
        self.stop_flag = False
        self.start_btn.config(state='disabled', bg='#9e9e9e')
        self.excel_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        self.result_label.config(text="")

        if bgm_files:
            self._run_bgm_preview_flow(groups, ordered, fixed_groups, count,
                                       res_w, res_h, output_dir, bgm_files,
                                       date_str, excel_path)
        else:
            t = threading.Thread(target=self._run_process,
                                 args=(groups, ordered, fixed_groups, count,
                                       res_w, res_h, output_dir, None),
                                 daemon=True)
            t.start()

    def _run_bgm_preview_flow(self, groups, ordered, fixed_groups, count,
                              res_w, res_h, output_dir, bgm_files,
                              date_str, excel_path):
        """BGM预览流程。"""
        selected_groups = self._get_selected_groups()
        group_keys = sorted(selected_groups.keys())
        first_video = selected_groups[group_keys[0]][0]
        first_bgm = bgm_files[0]
        video_info = probe_video(first_video['path'])

        def on_preview_done(action, volume):
            if action == 'cancel':
                self._log("⏹ 用户取消拼接", 'warn')
                self._reset_buttons()
                return

            if action == 'skip':
                bgm_display = " | BGM: 已跳过"
                self._log_clear()
                self._log("🎬 开始拼接任务...", 'title')
                mode_display = '有序' if ordered else ('无序' if not fixed_groups else f"部分固定(固定{''.join(fixed_groups)})")
                self._log(f"模式: {mode_display}{bgm_display} | 本次拼接: {count} 条")
                t = threading.Thread(target=self._run_process,
                                     args=(groups, ordered, fixed_groups, count,
                                           res_w, res_h, output_dir, None, None),
                                     daemon=True)
                t.start()
                return

            # action == 'confirm'
            bgm_volumes = {bgm['path']: volume for bgm in bgm_files}

            self._log_clear()
            self._log("🎬 开始拼接任务...", 'title')
            mode_display = '有序' if ordered else ('无序' if not fixed_groups else f"部分固定(固定{''.join(fixed_groups)})")
            self._log(f"模式: {mode_display} | BGM音量: {volume:.2f} | 本次拼接: {count} 条")

            self._launch_batch(
                groups, ordered, fixed_groups, count,
                res_w, res_h, output_dir, bgm_files, bgm_volumes)

        self._show_bgm_preview(
            video_path=first_video['path'],
            bgm_path=first_bgm['path'],
            bgm_name=first_bgm['name'],
            video_duration=video_info['duration'],
            callback=on_preview_done
        )

    def _launch_batch(self, groups, ordered, fixed_groups, count,
                      res_w, res_h, output_dir, bgm_files, bgm_volumes):
        t = threading.Thread(target=self._run_process,
                             args=(groups, ordered, fixed_groups, count,
                                   res_w, res_h, output_dir, bgm_files, bgm_volumes),
                             daemon=True)
        t.start()

    def _stop_process(self):
        self.stop_flag = True
        self._log("⏹ 用户请求停止...", 'warn')
        self.stop_btn.config(state='disabled')

    def _run_process(self, groups, ordered, fixed_groups, count, res_w, res_h, output_dir, bgm_files=None, bgm_volumes=None):
        """后台线程：检查表格 → 继续未完成 或 生成全部。"""
        try:
            self.root.after(0, lambda: self._log_clear())
            self.root.after(0, lambda: self._log("🎬 开始拼接任务...", 'title'))
            mode_display = '有序' if ordered else ('无序' if not fixed_groups else f"部分固定(固定{''.join(fixed_groups)})")
            bgm_display = f" | BGM: {len(bgm_files)}首" if bgm_files else ""
            self.root.after(0, lambda md=mode_display, bd=bgm_display: self._log(
                f"模式: {md}{bd} | 本次拼接: {count} 条"))

            input_dir = self.input_var.get()
            date_str = os.path.basename(input_dir)
            selected_keys = list(groups.keys())
            excel_suffix = f"固定{''.join(fixed_groups)}" if fixed_groups else None
            excel_path = get_excel_path(output_dir, date_str, ordered, selected_keys, excel_suffix)

            existing = load_existing_status(excel_path)

            if existing:
                pending_items = {k: v for k, v in existing.items()
                                 if '✅' not in v[0]}
                done_count = len(existing) - len(pending_items)

                self.root.after(0, lambda: self._log(
                    f"📊 已有表格: {os.path.basename(excel_path)}", 'info'))
                self.root.after(0, lambda: self._log(
                    f"   ✅ 已完成: {done_count}  ⏳待拼接/❌失败: {len(pending_items)}", 'info'))

                if not pending_items:
                    self.root.after(0, lambda: self._log(
                        "🎉 全部已拼接完成！", 'success'))
                    self.root.after(0, lambda: self.result_label.config(
                        text="✅ 全部已完成", fg=self.SUCCESS))
                    return

                video_lookup = build_video_lookup(groups)
                pending_perms = []
                row_map = {}

                for key, (status, row_num, _, names) in pending_items.items():
                    ordered_videos = [(video_lookup[n]['group'], video_lookup[n])
                                      for n in names if n in video_lookup]
                    if len(ordered_videos) != len(names):
                        self.root.after(0, lambda k=key: self._log(
                            f"  ⚠️ 无法解析排列 {k}，跳过", 'warn'))
                        continue

                    perm = {
                        'order': [g for g, _ in ordered_videos],
                        'videos': ordered_videos,
                        'concat_str': ' → '.join(v['name'] for _, v in ordered_videos),
                    }
                    pending_perms.append(perm)
                    row_map[perm_to_key(perm)] = row_num

                encode_count = min(count, len(pending_perms))
                encode_perms = pending_perms[:encode_count]

                self.root.after(0, lambda: self._log(
                    f"🔄 本次拼接 {encode_count} 条（剩余 {len(pending_perms)} 条）...", 'title'))

                success, fail, skip = self._encode_loop(
                    encode_perms, groups, excel_path, output_dir, date_str,
                    row_map=row_map, target_w=res_w, target_h=res_h,
                    bgm_files=bgm_files, bgm_volumes=bgm_volumes)

            else:
                self.root.after(0, lambda: self._log(
                    "未找到表格，正在生成所有排列...", 'title'))

                perms = generate_all_permutations(groups, ordered, fixed_groups)
                create_excel(perms, groups, excel_path)

                self.root.after(0, lambda: self._log(
                    f"📊 表格已创建: {os.path.basename(excel_path)} "
                    f"({len(perms):,} 种排列)", 'success'))

                encode_count = min(count, len(perms))
                encode_perms = perms[:encode_count]

                self.root.after(0, lambda: self._log(
                    f"🔄 本次拼接 {encode_count} 条", 'title'))

                success, fail, skip = self._encode_loop(
                    encode_perms, groups, excel_path, output_dir, date_str,
                    target_w=res_w, target_h=res_h,
                    bgm_files=bgm_files, bgm_volumes=bgm_volumes)

            result_text = (f"✅ 成功: {success}   ⏭️ 跳过: {skip}   "
                           f"❌ 失败: {fail}")
            self.root.after(0, lambda: self.result_label.config(
                text=result_text, fg=self.SUCCESS if fail == 0 else self.FAIL))
            self.root.after(0, lambda: self._log("=" * 50, 'title'))
            self.root.after(0, lambda: self._log(f"🎉 {result_text}", 'success'))
            self.root.after(0, lambda: self._log(
                f"📁 视频: {output_dir}", 'success'))
            self.root.after(0, lambda: self._log(
                f"📊 表格: {excel_path}", 'success'))

        except Exception as e:
            self.root.after(0, lambda: self._log(f"💥 严重错误: {e}", 'error'))
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
        finally:
            self.is_running = False
            self.root.after(0, self._reset_buttons)

    def _encode_loop(self, perms, groups, excel_path, output_dir, date_str,
                     row_map=None, target_w=None, target_h=None,
                     bgm_files=None, bgm_volumes=None):
        """编码循环。"""
        success = fail = skip = 0
        total = len(perms)

        bgm_mgr = BgmManager(bgm_files, bgm_volumes) if bgm_files else None
        use_bgm = bgm_mgr is not None

        for idx, perm in enumerate(perms):
            if self.stop_flag:
                self.root.after(0, lambda: self._log("⏹ 已停止", 'warn'))
                break

            key = perm_to_key(perm)
            output_name = perm_to_filename(perm, date_str)
            output_path = os.path.join(output_dir, output_name)

            if row_map:
                excel_row = row_map.get(key, idx + 2)
            else:
                excel_row = idx + 2

            done = success + fail + skip
            pct = (done / total * 100) if total else 0
            self.root.after(0, lambda p=pct, d=done, c=total, s=perm['concat_str']:
                self._update_progress(p, d, c, s))

            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                update_excel_status(excel_path, excel_row, '✅ 已完成',
                                    output_name, groups)
                skip += 1
                self.root.after(0, lambda n=output_name: self._log(
                    f"  ⏭️ {n} 已存在，跳过", 'warn'))
                continue

            video_list = [vid for _, vid in perm['videos']]
            try:
                ok = concat_videos(video_list, output_path, target_w, target_h)
                if ok and use_bgm:
                    bgm, bgm_vol = bgm_mgr.next()
                    info = probe_video(output_path)
                    bgm_output = output_path + '.bgm.mp4'
                    bgm_ok = add_bgm_to_video(output_path, bgm['path'],
                                              bgm_output, info['duration'],
                                              bgm_volume=bgm_vol)
                    if bgm_ok:
                        os.replace(bgm_output, output_path)
                        self.root.after(0, lambda n=output_name, b=bgm['name'], v=bgm_vol: self._log(
                            f"  🎵 {n} + BGM: {b} (vol={v:.2f})", 'info'))
                    else:
                        if os.path.exists(bgm_output):
                            os.remove(bgm_output)
                        self.root.after(0, lambda n=output_name: self._log(
                            f"  ⚠️ {n} BGM添加失败，保留原音频", 'warn'))
                if ok:
                    update_excel_status(excel_path, excel_row, '✅ 已完成',
                                        output_name, groups)
                    success += 1
                    self.root.after(0, lambda n=output_name: self._log(
                        f"  ✅ {n}", 'success'))
                else:
                    update_excel_status(excel_path, excel_row, '❌ 失败',
                                        '', groups)
                    fail += 1
                    self.root.after(0, lambda n=output_name: self._log(
                        f"  ❌ {n} 拼接失败", 'error'))
            except Exception as e:
                update_excel_status(excel_path, excel_row, '❌ 异常', '', groups)
                fail += 1
                self.root.after(0, lambda n=output_name, err=str(e): self._log(
                    f"  ❌ {n}: {err}", 'error'))

        return success, fail, skip

    def _update_progress(self, pct, done, total, current_str):
        self.progress_var.set(pct)
        self.progress_label.config(text=f"{done}/{total} ({pct:.0f}%)")
        if current_str:
            self._log(f"  [{done+1}/{total}] {current_str}")

    def _reset_buttons(self):
        self.start_btn.config(state='normal', bg=self.ACCENT)
        self.excel_btn.config(state='normal')
        self.stop_btn.config(state='disabled')


# ═══════════════════════════════════════════════════════
# 启动
# ═══════════════════════════════════════════════════════

def setup_ffmpeg_path():
    """将打包目录中的 ffmpeg 加入 PATH（PyInstaller 打包后生效）。"""
    if getattr(sys, 'frozen', False):
        internal_dir = os.path.join(os.path.dirname(sys.executable), '_internal')
    else:
        internal_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if os.path.isdir(internal_dir):
        os.environ['PATH'] = internal_dir + os.pathsep + os.environ.get('PATH', '')


def check_ffmpeg():
    try:
        subprocess.run(['ffmpeg', '-version'], capture_output=True, check=True, startupinfo=HIDDEN_SI)
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False


def main():
    setup_ffmpeg_path()
    if not check_ffmpeg():
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "缺少依赖",
            "未检测到 ffmpeg！\n\n"
            "请安装 ffmpeg 并确保在 PATH 中：\n"
            "  winget install ffmpeg")
        return

    root = tk.Tk()
    try:
        root.iconbitmap(default='')
    except Exception:
        pass
    VideoJoinerApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
