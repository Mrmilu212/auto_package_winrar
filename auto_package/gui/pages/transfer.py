import tkinter as tk
import time
from pathlib import Path
from auto_package.utils.logging_config import get_logger

logger = get_logger()

try:
    import windnd
except ImportError:
    windnd = None


class Transfer:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self._transfer_tasks = []  # 存储传输任务信息
        self._job_tick_id = None
        self._build_ui()
        self._start_job_timer()
    
    def _build_ui(self):
        """构建传输窗口的UI"""
        pad = {"padx": 12, "pady": 8}
        
        # 传输窗口标题
        title_label = tk.Label(
            self.parent,
            text="传输管理",
            font=("Segoe UI", 12, "bold"),
            fg="#333"
        )
        title_label.pack(anchor=tk.W, **pad)
        
        # 传输列表
        transfer_list = tk.Label(
            self.parent,
            text="传输任务列表",
            fg="#666"
        )
        transfer_list.pack(anchor=tk.W, **pad)
        
        # 任务列表容器
        self._tasks_frame = tk.Frame(self.parent)
        self._tasks_frame.pack(fill=tk.BOTH, expand=True, **pad)
        
        # 上传拖放区域
        drop_label = tk.Label(
            self.parent,
            text="拖放文件夹上传",
            fg="#666"
        )
        drop_label.pack(anchor=tk.W, **pad)
        
        self._drop_frame = tk.Frame(self.parent, relief=tk.SUNKEN, bd=2, bg="#e8f5e9", height=80)
        self._drop_frame.pack(fill=tk.X, **pad)
        self._drop_frame.pack_propagate(False)
        
        self._drop_label = tk.Label(
            self._drop_frame,
            text="拖放文件夹到此处直接上传",
            bg="#e8f5e9",
            fg="#456",
            font=("Segoe UI", 10),
        )
        self._drop_label.pack(expand=True)
        
        # 传输状态
        self._status_label = tk.Label(
            self.parent,
            text="传输状态：就绪",
            fg="#333"
        )
        self._status_label.pack(anchor=tk.W, **pad)
        
        # 注册拖放
        self._hook_drop()
    
    def _fmt_time(self, seconds: float) -> str:
        """格式化时间"""
        seconds = max(0.0, float(seconds))
        m, s = divmod(int(seconds + 0.5), 60)
        h, m = divmod(m, 60)
        if h:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"
    
    def _start_job_timer(self):
        """启动任务计时器"""
        def tick():
            # 更新所有任务的时间
            for task in self._transfer_tasks:
                if 'start_time' in task and 'time_label' in task:
                    elapsed = time.monotonic() - task['start_time']
                    task['time_label'].config(text=self._fmt_time(elapsed))
            # 继续计时
            self._job_tick_id = self.parent.after(200, tick)
        
        tick()
    
    def add_transfer_task(self, file_name, cancel_callback=None):
        """添加传输任务"""
        logger.info('添加传输任务: %s', file_name)
        # 创建任务行
        task_frame = tk.Frame(self._tasks_frame)
        task_frame.pack(fill=tk.X, pady=2, side=tk.TOP)
        
        # 文件名
        file_label = tk.Label(
            task_frame,
            text=file_name,
            anchor=tk.W,
            width=40
        )
        file_label.pack(side=tk.LEFT, padx=5)
        
        # 已用时间
        time_label = tk.Label(
            task_frame,
            text="00:00",
            anchor=tk.CENTER,
            width=10
        )
        time_label.pack(side=tk.LEFT, padx=5)
        
        # 取消按钮
        cancel_btn = tk.Button(
            task_frame,
            text="取消",
            width=8,
            command=lambda: self._cancel_task(task_frame, cancel_callback)
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        # 添加任务信息（插入到列表开头）
        task_info = {
            'frame': task_frame,
            'file_label': file_label,
            'time_label': time_label,
            'cancel_btn': cancel_btn,
            'start_time': time.monotonic(),
            'cancel_callback': cancel_callback
        }
        self._transfer_tasks.insert(0, task_info)
        
        # 更新状态
        self._update_status()
    
    def _cancel_task(self, task_frame, cancel_callback=None):
        """取消传输任务"""
        # 移除任务
        for task in self._transfer_tasks:
            if task['frame'] == task_frame:
                file_name = task['file_label'].cget('text')
                logger.info('取消传输任务: %s', file_name)
                # 调用取消回调
                if task.get('cancel_callback'):
                    task['cancel_callback']()
                # 移除任务
                self._transfer_tasks.remove(task)
                task_frame.destroy()
                break
        
        # 更新状态
        self._update_status()
    
    def _update_status(self):
        """更新传输状态"""
        task_count = len(self._transfer_tasks)
        if task_count == 0:
            self._status_label.config(text="传输状态：就绪")
        else:
            self._status_label.config(text=f"传输状态：{task_count} 个任务进行中")
    
    def remove_transfer_task(self, file_name):
        """根据文件名移除传输任务"""
        logger.debug('移除传输任务: %s', file_name)
        # 查找并移除任务
        for task in self._transfer_tasks:
            if task['file_label'].cget('text') == file_name:
                task['frame'].destroy()
                self._transfer_tasks.remove(task)
                break
        
        # 更新状态
        self._update_status()
    
    def _hook_drop(self):
        """注册拖放功能"""
        if windnd is None:
            return
        try:
            kw = {"func": self._on_drop_upload, "force_unicode": True}
            windnd.hook_dropfiles(self._drop_frame, **kw)
            windnd.hook_dropfiles(self._drop_label, **kw)
        except Exception as e:
            print(f"拖放注册失败: {e}")
    
    def _on_drop_upload(self, paths):
        """处理拖放的文件夹上传"""
        logger.info('拖放上传文件夹')
        decoded = []
        for p in paths:
            if isinstance(p, bytes):
                try:
                    decoded.append(p.decode('utf-8', errors='replace'))
                except Exception:
                    decoded.append(str(p))
            else:
                decoded.append(str(p))
        
        items = [Path(p) for p in decoded]
        
        # 只处理文件夹
        folders = [p for p in items if p.is_dir()]
        if not folders:
            logger.debug('拖放的不是文件夹，忽略')
            return
        
        # 只处理第一个文件夹
        folder = folders[0]
        if len(folders) > 1:
            folder = folders[0]
        
        logger.info('拖放上传文件夹: %s', folder)
        # 启动上传
        self.app.start_upload(folder)
