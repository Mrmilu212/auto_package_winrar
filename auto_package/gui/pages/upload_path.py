import tkinter as tk
from tkinter import ttk, messagebox
from auto_package.config.settings import _save_upload_paths
from auto_package.utils.logging_config import get_logger

logger = get_logger()


class UploadPath:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self._build_ui()
        self._refresh_paths()
    
    def _build_ui(self):
        pad = {"padx": 12, "pady": 8}
        
        # 标题
        title_frame = tk.Frame(self.parent)
        title_frame.pack(fill=tk.X, **pad)
        tk.Label(
            title_frame,
            text="上传路径管理",
            font=("Segoe UI", 12, "bold"),
            fg="#333"
        ).pack(anchor=tk.W)
        
        # 说明文字
        desc_frame = tk.Frame(self.parent)
        desc_frame.pack(fill=tk.X, **pad)
        tk.Label(
            desc_frame,
            text="管理百度网盘上传路径，点击路径行选中，点击单选按钮设置默认路径",
            fg="#666",
            justify=tk.LEFT,
            wraplength=500
        ).pack(anchor=tk.W)
        
        # 路径列表容器
        list_frame = tk.Frame(self.parent)
        list_frame.pack(fill=tk.BOTH, expand=True, **pad)
        
        # 路径列表
        self._paths_frame = tk.Frame(list_frame)
        self._paths_frame.pack(fill=tk.BOTH, expand=True)
        
        # 按钮区域
        btn_frame = tk.Frame(self.parent)
        btn_frame.pack(fill=tk.X, **pad)
        
        # 左侧按钮
        left_btn_frame = tk.Frame(btn_frame)
        left_btn_frame.pack(side=tk.LEFT)
        
        self.btn_add = tk.Button(
            left_btn_frame,
            text="新增",
            width=8,
            command=self._add_path
        )
        self.btn_add.pack(side=tk.LEFT, padx=(0, 8))
        
        self.btn_remove = tk.Button(
            left_btn_frame,
            text="删除",
            width=8,
            command=self._remove_path
        )
        self.btn_remove.pack(side=tk.LEFT, padx=(0, 8))
        
        self.btn_edit = tk.Button(
            left_btn_frame,
            text="编辑",
            width=8,
            command=self._edit_path,
            state=tk.DISABLED
        )
        self.btn_edit.pack(side=tk.LEFT, padx=(0, 8))
        
        # 右侧按钮
        right_btn_frame = tk.Frame(btn_frame)
        right_btn_frame.pack(side=tk.RIGHT)
        
        self.btn_up = tk.Button(
            right_btn_frame,
            text="上移",
            width=8,
            command=self._move_up
        )
        self.btn_up.pack(side=tk.LEFT, padx=(8, 0))
        
        self.btn_down = tk.Button(
            right_btn_frame,
            text="下移",
            width=8,
            command=self._move_down
        )
        self.btn_down.pack(side=tk.LEFT, padx=(8, 0))
        
        # 选中的路径索引
        self._selected_index = None
        self._path_widgets = []
        
        # 共享的默认路径变量
        self._default_var = tk.IntVar(value=0)
    
    def _refresh_paths(self):
        # 清空现有路径列表
        for widget in self._paths_frame.winfo_children():
            widget.destroy()
        self._path_widgets = []
        
        # 重新创建路径列表
        for i, path in enumerate(self.app._upload_path_history):
            path_frame = tk.Frame(self._paths_frame, relief=tk.RAISED, bd=1, bg="#f0f0f0")
            path_frame.pack(fill=tk.X, pady=2)
            
            # 单选按钮 - 设置默认路径
            def make_radio_handler(idx):
                return lambda: self._set_default_path(idx)
            
            radio = tk.Radiobutton(
                path_frame,
                variable=self._default_var,
                value=i,
                command=make_radio_handler(i)
            )
            radio.pack(side=tk.LEFT, padx=8, pady=4)
            
            # 路径文本
            label = tk.Label(
                path_frame,
                text=path,
                anchor=tk.W,
                bg="#f0f0f0"
            )
            label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=4)
            
            # 绑定点击事件
            def on_click(e, idx=i):
                self._select_path(idx)
            
            # 绑定双击事件
            def on_double_click(e, idx=i):
                self._select_path(idx)
                self._edit_path(idx)
            
            path_frame.bind("<Button-1>", on_click)
            label.bind("<Button-1>", on_click)
            path_frame.bind("<Double-Button-1>", on_double_click)
            label.bind("<Double-Button-1>", on_double_click)
            
            # 保存控件引用
            self._path_widgets.append({
                "frame": path_frame,
                "label": label,
                "path": path
            })
        
        # 设置默认路径对应的单选按钮状态
        default_index = self.app._upload_path_history.index(self.app._upload_path) if self.app._upload_path in self.app._upload_path_history else 0
        self._default_var.set(default_index)
        
        # 更新按钮状态
        self._update_button_states()
    
    def _select_path(self, index):
        """选择路径"""
        self._selected_index = index
        # 更新UI显示
        for i, widget in enumerate(self._path_widgets):
            if i == index:
                widget["frame"].config(bg="#e0f0ff")
                widget["label"].config(bg="#e0f0ff")
            else:
                widget["frame"].config(bg="#f0f0f0")
                widget["label"].config(bg="#f0f0f0")
        self._update_button_states()
    
    def _add_path(self):
        """新增路径"""
        dialog = tk.Toplevel(self.parent)
        dialog.title("新增上传路径")
        dialog.geometry("400x120")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        frame = tk.Frame(dialog, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="请输入百度网盘路径:", anchor=tk.W).pack(fill=tk.X, pady=(0, 8))
        entry_var = tk.StringVar(value="/")
        entry = tk.Entry(frame, textvariable=entry_var, width=40)
        entry.pack(fill=tk.X, pady=(0, 12))
        entry.focus_set()
        
        btn_frame = tk.Frame(frame)
        btn_frame.pack(side=tk.RIGHT)
        
        def on_ok():
            path = entry_var.get().strip()
            if path:
                if path not in self.app._upload_path_history:
                    self.app._upload_path_history.append(path)
                    logger.info('新增上传路径: %s', path)
                    self._save_and_refresh()
                else:
                    logger.debug('路径已存在: %s', path)
                dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        tk.Button(btn_frame, text="确定", width=8, command=on_ok).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_frame, text="取消", width=8, command=on_cancel).pack(side=tk.LEFT)
    
    def _edit_path(self, index=None):
        """编辑路径"""
        if index is None:
            index = self._selected_index
        
        if index is None or index < 0 or index >= len(self.app._upload_path_history):
            return
        
        old_path = self.app._upload_path_history[index]
        
        dialog = tk.Toplevel(self.parent)
        dialog.title("编辑上传路径")
        dialog.geometry("400x120")
        dialog.transient(self.parent)
        dialog.grab_set()
        
        frame = tk.Frame(dialog, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(frame, text="请输入百度网盘路径:", anchor=tk.W).pack(fill=tk.X, pady=(0, 8))
        entry_var = tk.StringVar(value=old_path)
        entry = tk.Entry(frame, textvariable=entry_var, width=40)
        entry.pack(fill=tk.X, pady=(0, 12))
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        btn_frame = tk.Frame(frame)
        btn_frame.pack(side=tk.RIGHT)
        
        def on_ok():
            new_path = entry_var.get().strip()
            if new_path and new_path != old_path:
                # 检查是否与其他路径重复
                if new_path in self.app._upload_path_history:
                    idx = self.app._upload_path_history.index(new_path)
                    if idx != index:
                        messagebox.showwarning("提示", "该路径已存在")
                        logger.debug('编辑路径失败，路径已存在: %s', new_path)
                        return
                self.app._upload_path_history[index] = new_path
                logger.info('编辑上传路径: %s -> %s', old_path, new_path)
                # 如果修改的是默认路径，更新默认路径
                if self.app._upload_path == old_path:
                    self.app._upload_path = new_path
                    logger.debug('默认路径已更新: %s', new_path)
                self._save_and_refresh()
            elif new_path == old_path:
                logger.debug('路径未修改')
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        tk.Button(btn_frame, text="确定", width=8, command=on_ok).pack(side=tk.LEFT, padx=(0, 8))
        tk.Button(btn_frame, text="取消", width=8, command=on_cancel).pack(side=tk.LEFT)
    
    def _remove_path(self):
        """删除路径"""
        if self._selected_index is not None:
            if len(self.app._upload_path_history) <= 1:
                messagebox.showinfo("提示", "至少需要保留一个路径")
                logger.debug('删除路径失败：至少需要保留一个路径')
                return
            
            # 移除选中的路径
            deleted_path = self.app._upload_path_history[self._selected_index]
            del self.app._upload_path_history[self._selected_index]
            logger.info('删除上传路径: %s', deleted_path)
            
            # 如果删除的是默认路径之前的索引，默认路径索引需要调整
            default_index = self.app._upload_path_history.index(self.app._upload_path) if self.app._upload_path in self.app._upload_path_history else 0
            
            # 更新默认路径
            if self.app._upload_path not in self.app._upload_path_history:
                old_default = self.app._upload_path
                self.app._upload_path = self.app._upload_path_history[0] if self.app._upload_path_history else "/测试"
                logger.debug('默认路径已更新: %s -> %s', old_default, self.app._upload_path)
            
            self._save_and_refresh()
            self._selected_index = None
    
    def _move_up(self):
        """上移路径"""
        if self._selected_index is not None and self._selected_index > 0:
            # 交换位置
            old_pos = self._selected_index
            new_pos = self._selected_index - 1
            self.app._upload_path_history[self._selected_index], self.app._upload_path_history[self._selected_index - 1] = \
                self.app._upload_path_history[self._selected_index - 1], self.app._upload_path_history[self._selected_index]
            logger.debug('移动上传路径: %s, 从 %d 到 %d', self.app._upload_path_history[new_pos], old_pos, new_pos)
            
            # 如果移动的是默认路径，更新默认路径
            if self._selected_index == 0:
                self.app._upload_path = self.app._upload_path_history[0]
                logger.debug('默认路径已更新: %s', self.app._upload_path)
            
            self._save_and_refresh()
            self._selected_index -= 1
    
    def _move_down(self):
        """下移路径"""
        if self._selected_index is not None and self._selected_index < len(self.app._upload_path_history) - 1:
            # 交换位置
            old_pos = self._selected_index
            new_pos = self._selected_index + 1
            self.app._upload_path_history[self._selected_index], self.app._upload_path_history[self._selected_index + 1] = \
                self.app._upload_path_history[self._selected_index + 1], self.app._upload_path_history[self._selected_index]
            logger.debug('移动上传路径: %s, 从 %d 到 %d', self.app._upload_path_history[new_pos], old_pos, new_pos)
            
            # 如果移动的是默认路径，更新默认路径
            if self._selected_index == 0:
                self.app._upload_path = self.app._upload_path_history[0]
                logger.debug('默认路径已更新: %s', self.app._upload_path)
            
            self._save_and_refresh()
            self._selected_index += 1
    
    def _set_default_path(self, index):
        """设置默认路径"""
        if 0 <= index < len(self.app._upload_path_history):
            old_default = self.app._upload_path
            self.app._upload_path = self.app._upload_path_history[index]
            logger.info('设置默认上传路径: %s', self.app._upload_path)
            self._save_and_refresh()
    
    def _save_and_refresh(self):
        """保存并刷新"""
        _save_upload_paths(self.app._upload_path, self.app._upload_path_history)
        self._refresh_paths()
        if hasattr(self.app, "_home_page") and self.app._home_page:
            self.app._home_page.refresh_upload_path_combo()
    
    def _update_button_states(self):
        """更新按钮状态"""
        has_selection = self._selected_index is not None
        self.btn_remove.config(state=tk.NORMAL if has_selection else tk.DISABLED)
        self.btn_edit.config(state=tk.NORMAL if has_selection else tk.DISABLED)
        
        if has_selection:
            self.btn_up.config(state=tk.NORMAL if self._selected_index > 0 else tk.DISABLED)
            self.btn_down.config(state=tk.NORMAL if self._selected_index < len(self.app._upload_path_history) - 1 else tk.DISABLED)
        else:
            self.btn_up.config(state=tk.DISABLED)
            self.btn_down.config(state=tk.DISABLED)
