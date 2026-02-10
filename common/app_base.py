"""
基础应用窗口类 (PySide6) - 所有小工具的通用 UI 框架

提供：
- 主窗口管理（Fusion 风格，跨平台一致外观）
- 配置区域（子类填充）
- 开始/停止 按钮
- 进度条 + 状态文字
- 日志面板（彩色）
- 后台线程任务管理（基于 QThread + Signal）

子类只需要重写:
    create_content(layout)  - 创建工具特定的配置 UI
    do_work()               - 在后台线程中执行任务逻辑
    validate()              - 点击开始前的校验（可选）
"""

import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QPushButton, QProgressBar, QTextEdit, QLabel,
    QMessageBox, QSizePolicy, QDialog, QScrollArea,
)
from PySide6.QtCore import Qt, QThread, Signal, Slot
from PySide6.QtGui import QFont, QColor, QTextCharFormat, QTextCursor, QPixmap


# ================================================================
#  后台工作线程
# ================================================================

class _WorkerThread(QThread):
    """后台任务线程，执行 do_work() 并在异常时发出信号"""
    error_signal = Signal(str)

    def __init__(self, func, parent=None):
        super().__init__(parent)
        self._func = func

    def run(self):
        try:
            self._func()
        except Exception as e:
            self.error_signal.emit(str(e))


# ================================================================
#  基础应用窗口
# ================================================================

class BaseApp(QMainWindow):
    """
    基础应用窗口 - 所有小工具继承此类

    使用方法::

        class MyTool(BaseApp):
            APP_NAME = "我的工具"
            APP_VERSION = "1.0"

            def create_content(self, layout):
                # layout 是 QVBoxLayout，在此添加配置控件
                ...

            def validate(self):
                # 返回 True 表示校验通过
                return True

            def do_work(self):
                # 后台任务逻辑（在子线程中运行）
                self.log("正在处理...")
                self.update_progress(50, "处理中")
                ...

        if __name__ == "__main__":
            MyTool().run()
    """

    # 线程安全信号 —— 子线程 emit，主线程 slot 处理 UI
    _log_signal = Signal(str, str)
    _progress_signal = Signal(float, str)

    # --- 子类可覆盖的配置 ---
    APP_NAME = "工具"
    APP_VERSION = "1.0"
    WINDOW_WIDTH = 880
    WINDOW_HEIGHT = 720

    # --- 推广信息配置（子类可覆盖）---
    # 设置为 None 则不显示推广栏
    PROMO_TEXT = None          # 如: "关注公众号获取最新版本"
    PROMO_IMAGES = None        # 如: [("公众号", "path/to/qr1.png"), ("微信", "path/to/qr2.png")]
    PROMO_IMAGE_SIZE = 120     # 二维码图片显示尺寸

    def __init__(self):
        # 确保只有一个 QApplication 实例
        self._app = QApplication.instance() or QApplication(sys.argv)
        self._app.setStyle("Fusion")

        super().__init__()
        self.setWindowTitle(f"{self.APP_NAME} v{self.APP_VERSION}")
        self.resize(self.WINDOW_WIDTH, self.WINDOW_HEIGHT)
        self.setMinimumSize(650, 500)

        # 任务状态
        self._running = False
        self._stop_requested = False
        self._worker = None

        # 连接信号到槽
        self._log_signal.connect(self._on_log)
        self._progress_signal.connect(self._on_progress)

        # 构建界面
        self._build_ui()
        self._center_window()

    # ================================================================
    #  UI 构建
    # ================================================================

    def _center_window(self):
        """窗口居中显示"""
        screen = self._app.primaryScreen().availableGeometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def _build_ui(self):
        """构建完整 UI 布局"""
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(10)

        # 1. 配置区域（子类填充）
        self.config_group = QGroupBox("配置")
        config_inner = QVBoxLayout(self.config_group)
        main_layout.addWidget(self.config_group)
        self.create_content(config_inner)

        # 2. 操作按钮
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)

        self.start_btn = QPushButton("▶ 开始处理")
        self.start_btn.setCursor(Qt.PointingHandCursor)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50; color: white;
                border: none; padding: 8px 28px; border-radius: 4px;
                font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #43A047; }
            QPushButton:pressed { background-color: #388E3C; }
            QPushButton:disabled { background-color: #C8E6C9; color: #999; }
        """)
        self.start_btn.clicked.connect(self._on_start)

        self.stop_btn = QPushButton("■ 停止")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #F44336; color: white;
                border: none; padding: 8px 28px; border-radius: 4px;
                font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #E53935; }
            QPushButton:pressed { background-color: #D32F2F; }
            QPushButton:disabled { background-color: #FFCDD2; color: #999; }
        """)
        self.stop_btn.clicked.connect(self._on_stop)

        btn_layout.addWidget(self.start_btn)
        btn_layout.addWidget(self.stop_btn)
        btn_layout.addStretch()
        main_layout.addLayout(btn_layout)

        # 3. 进度条 + 状态
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ddd; border-radius: 4px;
                text-align: center; height: 22px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50; border-radius: 3px;
            }
        """)
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #666; font-size: 12px;")
        main_layout.addWidget(self.status_label)

        # 4. 日志面板
        log_group = QGroupBox("日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(6, 6, 6, 6)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        mono_font = self._get_mono_font()
        self.log_text.setFont(mono_font)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #FAFAFA; border: 1px solid #E0E0E0;
                border-radius: 4px; padding: 4px;
            }
        """)
        log_layout.addWidget(self.log_text)
        main_layout.addWidget(log_group, stretch=1)

        # 5. 推广信息栏（如果子类配置了）
        if self.PROMO_TEXT or self.PROMO_IMAGES:
            self._build_promo_bar(main_layout)

    @staticmethod
    def _get_mono_font():
        """获取当前平台合适的等宽字体"""
        if sys.platform == 'darwin':
            return QFont("Menlo", 11)
        elif sys.platform == 'win32':
            return QFont("Consolas", 9)
        else:
            return QFont("Monospace", 10)

    def _build_promo_bar(self, parent_layout):
        """构建底部推广信息栏"""
        promo_frame = QWidget()
        promo_frame.setStyleSheet("""
            QWidget {
                background-color: #F5F5F5;
                border: 1px solid #E0E0E0;
                border-radius: 6px;
            }
        """)
        promo_layout = QHBoxLayout(promo_frame)
        promo_layout.setContentsMargins(12, 8, 12, 8)
        promo_layout.setSpacing(12)

        # 文字
        if self.PROMO_TEXT:
            text_label = QLabel(self.PROMO_TEXT)
            text_label.setStyleSheet(
                "color: #555; font-size: 12px; border: none; background: none;"
            )
            text_label.setWordWrap(True)
            promo_layout.addWidget(text_label, stretch=1)

        # 查看详情按钮
        if self.PROMO_IMAGES:
            detail_btn = QPushButton("查看详情")
            detail_btn.setCursor(Qt.PointingHandCursor)
            detail_btn.setStyleSheet("""
                QPushButton {
                    background-color: #1976D2; color: white;
                    border: none; padding: 6px 18px; border-radius: 4px;
                    font-size: 12px;
                }
                QPushButton:hover { background-color: #1565C0; }
            """)
            detail_btn.clicked.connect(self._show_promo_dialog)
            promo_layout.addWidget(detail_btn)

        parent_layout.addWidget(promo_frame)

    def _show_promo_dialog(self):
        """弹出推广详情弹窗（显示二维码图片）"""
        dialog = QDialog(self)
        dialog.setWindowTitle("联系我们")
        dialog.setMinimumWidth(400)

        dlg_layout = QVBoxLayout(dialog)
        dlg_layout.setSpacing(16)
        dlg_layout.setContentsMargins(20, 20, 20, 20)

        if self.PROMO_TEXT:
            text_lbl = QLabel(self.PROMO_TEXT)
            text_lbl.setStyleSheet("font-size: 14px; color: #333;")
            text_lbl.setWordWrap(True)
            text_lbl.setAlignment(Qt.AlignCenter)
            dlg_layout.addWidget(text_lbl)

        # 图片区域
        if self.PROMO_IMAGES:
            img_row = QHBoxLayout()
            img_row.setSpacing(24)
            img_row.addStretch()

            for label_text, img_path in self.PROMO_IMAGES:
                col = QVBoxLayout()
                col.setSpacing(6)

                # 图片
                img_label = QLabel()
                img_label.setAlignment(Qt.AlignCenter)
                pixmap = self._load_promo_image(img_path)
                if pixmap and not pixmap.isNull():
                    scaled = pixmap.scaled(
                        self.PROMO_IMAGE_SIZE, self.PROMO_IMAGE_SIZE,
                        Qt.KeepAspectRatio, Qt.SmoothTransformation
                    )
                    img_label.setPixmap(scaled)
                else:
                    img_label.setText(f"[{label_text}]")
                    img_label.setStyleSheet(
                        "color: #999; font-size: 11px; "
                        f"min-width: {self.PROMO_IMAGE_SIZE}px; "
                        f"min-height: {self.PROMO_IMAGE_SIZE}px; "
                        "border: 1px dashed #ccc; border-radius: 4px;"
                    )
                    img_label.setAlignment(Qt.AlignCenter)

                col.addWidget(img_label)

                # 标签
                name_lbl = QLabel(label_text)
                name_lbl.setAlignment(Qt.AlignCenter)
                name_lbl.setStyleSheet("font-size: 12px; color: #666;")
                col.addWidget(name_lbl)

                img_row.addLayout(col)

            img_row.addStretch()
            dlg_layout.addLayout(img_row)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.setCursor(Qt.PointingHandCursor)
        close_btn.setStyleSheet("""
            QPushButton {
                padding: 6px 30px; border-radius: 4px;
                border: 1px solid #ccc; font-size: 12px;
            }
            QPushButton:hover { background-color: #f0f0f0; }
        """)
        close_btn.clicked.connect(dialog.accept)
        dlg_layout.addWidget(close_btn, alignment=Qt.AlignCenter)

        dialog.exec()

    def _load_promo_image(self, img_path):
        """
        加载推广图片，支持相对路径。

        搜索顺序:
        1. 绝对路径直接使用
        2. 相对于 RESOURCE_BASE_FILE 所在目录（子类定义）
        3. 相对于 get_app_dir()
        4. 相对于 sys.argv[0] 所在目录
        """
        from common.utils import get_app_dir

        p = Path(img_path)
        if p.is_absolute() and p.exists():
            return QPixmap(str(p))

        # 候选基准目录
        candidates = []

        # 子类可以设置 RESOURCE_BASE_FILE = __file__，
        # 这样就以子类 .py 文件所在目录为基准
        base_file = getattr(self, 'RESOURCE_BASE_FILE', None)
        if base_file:
            candidates.append(Path(base_file).resolve().parent)

        candidates.append(get_app_dir())
        candidates.append(Path(sys.argv[0]).resolve().parent if sys.argv else Path.cwd())

        for base in candidates:
            full = base / img_path
            if full.exists():
                return QPixmap(str(full))

        return None

    # ================================================================
    #  子类接口（重写这些方法）
    # ================================================================

    def create_content(self, layout):
        """
        子类重写：在 layout (QVBoxLayout) 中添加配置控件

        :param layout: 配置区域的布局管理器
        """
        pass

    def validate(self):
        """子类重写：点击开始前的校验，返回 True 才会执行任务"""
        return True

    def do_work(self):
        """
        子类重写：在后台线程中执行实际任务

        可安全调用:
            self.log(message, level)
            self.update_progress(value, text)
            self.should_stop  (属性，检查用户是否请求停止)
        """
        pass

    # ================================================================
    #  公共 API（子类在 do_work 中调用）
    # ================================================================

    def log(self, message, level="info"):
        """
        线程安全的日志输出

        :param message: 日志内容
        :param level: info / success / warning / error
        """
        self._log_signal.emit(str(message), level)

    def update_progress(self, value, text=""):
        """
        线程安全的进度更新

        :param value: 0~100
        :param text: 可选的状态文字
        """
        self._progress_signal.emit(float(value), text)

    @property
    def should_stop(self):
        """检查用户是否请求停止（在 do_work 循环中使用）"""
        return self._stop_requested

    # ================================================================
    #  信号槽（主线程中执行 UI 更新）
    # ================================================================

    @Slot(str, str)
    def _on_log(self, message, level):
        """日志信号的槽函数"""
        colors = {
            'info': '#1976D2',
            'success': '#388E3C',
            'warning': '#F57C00',
            'error': '#D32F2F',
        }
        symbols = {
            'info': 'ℹ',
            'success': '✓',
            'warning': '⚠',
            'error': '✗',
        }

        color = colors.get(level, colors['info'])
        symbol = symbols.get(level, 'ℹ')

        fmt = QTextCharFormat()
        fmt.setForeground(QColor(color))

        cursor = self.log_text.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(f" {symbol} {message}\n", fmt)
        self.log_text.setTextCursor(cursor)
        self.log_text.ensureCursorVisible()

    @Slot(float, str)
    def _on_progress(self, value, text):
        """进度信号的槽函数"""
        self.progress_bar.setValue(int(min(value, 100)))
        if text:
            self.status_label.setText(text)

    # ================================================================
    #  内部逻辑
    # ================================================================

    def _on_start(self):
        """开始按钮点击"""
        if self._running:
            return
        if not self.validate():
            return

        # 清空日志
        self.log_text.clear()

        self._running = True
        self._stop_requested = False
        self.progress_bar.setValue(0)
        self.status_label.setText("处理中...")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)

        # 启动后台线程
        self._worker = _WorkerThread(self.do_work, parent=self)
        self._worker.finished.connect(self._on_done)
        self._worker.error_signal.connect(
            lambda msg: self.log(f"任务异常: {msg}", "error")
        )
        self._worker.start()

    def _on_stop(self):
        """停止按钮点击"""
        if self._running:
            self._stop_requested = True
            self.log("正在停止任务...", "warning")

    def _on_done(self):
        """任务结束后恢复 UI 状态"""
        self._running = False
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self._worker = None

    def run(self):
        """启动应用主循环"""
        self.show()
        sys.exit(self._app.exec())
