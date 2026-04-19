import tkinter as tk
from tkinter import scrolledtext, ttk
from pathlib import Path
from typing import Iterable
import subprocess
from auto_package.core.compress import find_rar_exe, run_rar_archive, run_double_compress, run_triple_compress
from auto_package.core.extract import find_winrar_exe, run_auto_extract
from auto_package.core.utils import decode_drop_path, _commit_outputs_atomic
from auto_package.config.settings import _save_upload_paths
from auto_package.utils.logging_config import get_logger

logger = get_logger()

try:
    import windnd
except ImportError:
    windnd = None

class Home:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self._rar = find_rar_exe()
        self._winrar = find_winrar_exe()
        self._build_ui()
        self._hook_drop()
    
    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}

        if self._rar:
            rar_line = f"已找到: {self._rar}"
        else:
            rar_line = "未找到 WinRAR（Rar.exe）。请安装 WinRAR 或检查安装路径。"

        frm_top = tk.Frame(self.parent)
        frm_top.pack(fill=tk.X, **pad)

        self.lbl_hint = tk.Label(
            frm_top,
            text="将文件/文件夹拖放到下方区域，或点击「选择文件夹」",
            justify=tk.LEFT,
            wraplength=520,
        )
        self.lbl_hint.pack(anchor=tk.W)

        self.lbl_rar = tk.Label(frm_top, text=rar_line, fg="#333", justify=tk.LEFT)
        self.lbl_rar.pack(anchor=tk.W, pady=(4, 0))

        self.chk_double = tk.Checkbutton(
            frm_top,
            text="二次伪装后缀压缩",
            variable=self.app._var_double,
            justify=tk.LEFT,
            wraplength=520,
            anchor=tk.W,
        )
        self.chk_double.pack(anchor=tk.W, pady=(6, 0))

        self.chk_triple = tk.Checkbutton(
            frm_top,
            text="三次分卷压缩",
            variable=self.app._var_triple,
            justify=tk.LEFT,
            wraplength=520,
            anchor=tk.W,
        )
        self.chk_triple.pack(anchor=tk.W, pady=(2, 0))

        # 百度网盘上传选项
        self._var_upload = tk.BooleanVar(value=False)
        upload_frame = tk.Frame(frm_top)
        upload_frame.pack(anchor=tk.W, pady=(6, 0))
        self.chk_upload = tk.Checkbutton(
            upload_frame,
            text="上传到百度网盘",
            variable=self.app._var_upload,
            justify=tk.LEFT,
        )
        self.chk_upload.pack(side=tk.LEFT)
        
        tk.Label(upload_frame, text="上传路径:").pack(side=tk.LEFT, padx=(12, 4))
        self._upload_path_var = tk.StringVar(value=self.app._upload_path)
        self._upload_path_combo = ttk.Combobox(
            upload_frame,
            textvariable=self._upload_path_var,
            values=self.app._upload_path_history,
            width=20,
            state="normal",
        )
        self._upload_path_combo.pack(side=tk.LEFT)
        self._upload_path_combo.bind("<<ComboboxSelected>>", self._on_upload_path_changed)
        self._upload_path_combo.bind("<FocusOut>", self._on_upload_path_changed)
        self._upload_path_combo.bind("<Return>", self._on_upload_path_changed)
        
        # 管理按钮
        self.btn_manage = tk.Button(
            upload_frame,
            text="管理",
            width=6,
            command=self._open_upload_path_page
        )
        self.btn_manage.pack(side=tk.LEFT, padx=(8, 0))

        # 拖放区域容器
        drop_container = tk.Frame(self.parent)
        drop_container.pack(fill=tk.BOTH, expand=True, **pad)

        # 压缩拖放区域
        compress_frame = tk.Frame(drop_container)
        compress_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 6))
        tk.Label(compress_frame, text="压缩拖放区", font=("Segoe UI", 10, "bold"), fg="#666").pack(anchor=tk.W, pady=(0, 4))
        self.drop_frame = tk.Frame(compress_frame, relief=tk.SUNKEN, bd=2, bg="#e8eef5", width=250, height=250)
        self.drop_frame.pack(fill=tk.BOTH, expand=True)
        self.drop_frame.pack_propagate(False)

        self.lbl_drop = tk.Label(
            self.drop_frame,
            text="拖放文件/文件夹",
            bg="#e8eef5",
            fg="#456",
            font=("Segoe UI", 12),
        )
        self.lbl_drop.pack(expand=True)

        # 解压拖放区域
        extract_frame = tk.Frame(drop_container)
        extract_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(6, 0))
        tk.Label(extract_frame, text="解压拖放区", font=("Segoe UI", 10, "bold"), fg="#666").pack(anchor=tk.W, pady=(0, 4))
        self.extract_drop_frame = tk.Frame(extract_frame, relief=tk.SUNKEN, bd=2, bg="#f5e8ee", width=250, height=250)
        self.extract_drop_frame.pack(fill=tk.BOTH, expand=True)
        self.extract_drop_frame.pack_propagate(False)

        self.lbl_extract_drop = tk.Label(
            self.extract_drop_frame,
            text="拖放要解压的文件",
            bg="#f5e8ee",
            fg="#644",
            font=("Segoe UI", 12),
        )
        self.lbl_extract_drop.pack(expand=True)

        btn_row = tk.Frame(self.parent)
        btn_row.pack(fill=tk.X, **pad)

        tk.Button(btn_row, text="选择文件夹…", command=self.app._pick_folder).pack(
            side=tk.LEFT
        )
        tk.Button(btn_row, text="清空日志", command=self.app._clear_log).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        tk.Button(btn_row, text="打开日志文件夹", command=self.app._open_log_folder).pack(
            side=tk.LEFT, padx=(8, 0)
        )
        
        # 按钮容器
        btn_right = tk.Frame(btn_row)
        btn_right.pack(side=tk.RIGHT)
        
        self.btn_cancel = tk.Button(
            btn_right, text="取消压缩", command=self.app._cancel_current, state=tk.DISABLED
        )
        self.btn_cancel.pack(side=tk.LEFT, padx=(0, 8))
        
        self.btn_cancel_extract = tk.Button(
            btn_right, text="取消解压", command=self.app._cancel_extract, state=tk.DISABLED
        )
        self.btn_cancel_extract.pack(side=tk.LEFT)

        self.log = scrolledtext.ScrolledText(self.parent, height=8, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 12))

        self.status = tk.Label(self.parent, textvariable=self.app._status_var, anchor=tk.W)
        self.status.pack(fill=tk.X, padx=12, pady=(0, 10))

        self.app._log_line(
            "就绪。默认：每进行一次压缩都会生成 5 位随机文件名（英文大小写+数字）的 .rar。"
            " 二次：在第一次产物基础上改伪装后缀再打包，外层 .rar 也会是新的随机名。"
            " 三次（分卷）：单卷目标≈总/2+10MB，最大≤2049MB，至少 2 卷；过小将取消。"
        )
        self.app._log_line("解压：拖放文件到右侧区域，自动识别并解压，支持去伪装。")
        self.app._set_status(0.0)
    
    def _on_upload_path_changed(self, event=None):
        new_path = self._upload_path_var.get().strip()
        if new_path and new_path != self.app._upload_path:
            old_path = self.app._upload_path
            self.app._upload_path = new_path
            if new_path not in self.app._upload_path_history:
                self.app._upload_path_history.insert(0, new_path)
                self.app._upload_path_history = self.app._upload_path_history[:10]
                self._upload_path_combo["values"] = self.app._upload_path_history
                logger.debug('新增上传路径到历史: %s', new_path)
            else:
                logger.debug('更新当前上传路径: %s -> %s', old_path, new_path)
            _save_upload_paths(self.app._upload_path, self.app._upload_path_history)
    
    def refresh_upload_path_combo(self):
        """刷新上传路径下拉框，与上传路径管理页面同步"""
        self._upload_path_combo["values"] = self.app._upload_path_history
        if self.app._upload_path in self.app._upload_path_history:
            self._upload_path_var.set(self.app._upload_path)
    
    def _open_upload_path_page(self):
        """打开上传路径管理页面"""
        if hasattr(self.app, "_notebook") and self.app._notebook:
            for i in range(self.app._notebook.index("end")):
                if self.app._notebook.tab(i, "text") == "上传路径":
                    self.app._notebook.select(i)
                    break
    
    def _hook_drop(self):
        if windnd is None:
            self.app._log_line("未安装 windnd，拖放不可用。请运行: pip install -r requirements.txt")
            logger.warning('未安装 windnd，拖放不可用')
            return
        try:
            # 压缩拖放区域
            kw = {"func": self.app._on_drop_files, "force_unicode": True}
            windnd.hook_dropfiles(self.drop_frame, **kw)
            windnd.hook_dropfiles(self.lbl_drop, **kw)
            # 解压拖放区域
            kw_extract = {"func": self.app._on_drop_extract, "force_unicode": True}
            windnd.hook_dropfiles(self.extract_drop_frame, **kw_extract)
            windnd.hook_dropfiles(self.lbl_extract_drop, **kw_extract)
            logger.debug('拖放注册成功')
        except Exception as e:
            self.app._log_line(f"拖放注册失败: {e}")
            logger.warning('拖放注册失败: %s', e)
