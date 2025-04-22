'''
python 3.10.11
3.27
倒计时程序功能升级,读上传的代码,完成新功能:
1.简单任务,字体更新用下拉别表,显示黑体\微软雅黑\宋体等10个字体吧,字号也用下拉别表进行更新设置.
2.欢迎词设置:缺省内容:欢迎您参加答辩,答辩时长XX分钟,XX就是倒计时时长,是否播放,用户可选.
3.倒计时目前是10秒,语音提示,提示词可定义,倒计时提示时间也可定制.
pyinstaller --onefile --noconsole --hidden-import=pyttsx3 --hidden-import=comtypes --hidden-import=psutil --icon=timer.ico count_down12.py
3.28
语音倒计时升级，基于目前的代码进行改进，：
1、与PPT运行的状态建立关联，PPT全屏播放时候，启动程序，根据预设时间进行倒计时。
2、时间到，强制PPT关闭PPT
3、python模式下没有生成EXE 文件时候，如何进行功能测试，帮忙设计测试方案。
4、设置界面优化，语音、黑屏逻辑调整
3.29 
显示当前时间
完善字体设置
PPT运行关联
倒计时窗口前置
4.1
倒计时字体错误修正，倒计时初始化5分钟
关联WPS启动全屏

4.2
背景颜色设置
背景色透明度控制
4.3
移除主窗体置顶，优化显示与操作逻辑
倒计时显示窗体下方label字体背景色设置
4.20
pyinstaller --onefile --icon=timer.ico --noconsole --hidden-import=pyttsx3 --hidden-import=pyttsx3.drivers.sapi5 --hidden-import=pyttsx3.voice --hidden-import=comtypes.client  --hidden-import=comtypes.gen --hidden-import=win32gui --hidden-import=win32process  --hidden-import=win32api --hidden-import=ctypes --hidden-import=numpy --hidden-import=psutil --add-data "comtypes\gen;comtypes\gen" --uac-admin count_down12.py
'''
# pyinstaller --onefile --icon=timer.ico --noconsole --hidden-import pyttsx3.drivers.sapi5 --hidden-import pyttsx3.voice count_down2.py WPS
import tkinter as tk
from tkinter import messagebox, simpledialog, colorchooser
import time
import json
import os
import pyttsx3
from tkinter import ttk
import win32gui
import win32process
import psutil
import win32api
import win32con
import os
import sys


def get_config_path():
    if getattr(sys, 'frozen', False):
        # 打包后的路径：EXE 同级目录
        return os.path.join(os.path.dirname(sys.executable), 'config.json')
    else:
        # 开发环境路径：代码所在目录
        return os.path.join(os.path.dirname(__file__), 'config.json')


# 强制设置 comtypes 生成路径
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
    os.environ["GEN_DIR"] = os.path.join(base_dir, "comtypes", "gen")
else:
    os.environ["GEN_DIR"] = os.path.join(
        os.path.dirname(__file__), "comtypes", "gen")


class CountdownSettings:
    def __init__(self):
        self.alert_red_min = 1    # 默认变红提醒时间（分钟）
        self.alert_voice_sec = 10  # 默认语音提醒时间（秒）
        self.black_screen_sec = 10  # 默认黑屏持续时间
        self.enable_welcome = True
        self.welcome_message = "欢迎您参加答辩，答辩时长5分钟"
        self.voice_message = "剩余XX秒"
        self.voice_alert_time = 10
        # 倒计时显示字体
        self.countdown_font_name = "微软雅黑"
        self.countdown_font_size = 48
        # 当前时间字体
        self.current_time_font_name = "Arial"
        self.current_time_font_size = 12
        # 时间到提示字体
        self.time_up_font_name = "黑体"
        self.time_up_font_size = 72
        self.always_black = False  # 是否永久黑屏
        self.enable_red_alert = True    # 新增：是否启用变红提醒
        self.enable_voice_alert = True  # 新增：是否启用语音提醒
        self.enable_black_screen = True  #
        self.show_current_time = False  # 新增：是否显示当前时间
        # ...其他原有属性
        self.background_color = "#FFFFFF"  # 新增背景颜色属性
        self.font_color = "#000000"       # 新增字体颜色属性
        self.alpha = 1.0  # 新增透明度属性（0.0-1.0）
        self.load_settings()
        try:
            self.engine = pyttsx3.init()
        except Exception as e:
            self.engine = None
            with open("error.log", "a") as f:
                f.write(f"[{time.ctime()}] 语音引擎初始化失败: {str(e)}\n")

    def save_settings(self):
        # 使用 UTF-8 编码写入文件
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(self.__dict__, f, ensure_ascii=False, indent=2)  # 禁用 ASCII 转义，格式化输出

    def load_settings(self):
        try:
            # 使用明确的 UTF-8 编码读取文件
            with open('config.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
                self.__dict__.update(data)
        except FileNotFoundError:
            # 文件不存在时初始化默认配置
            self.save_settings()
        except json.JSONDecodeError as e:
            # 捕获 JSON 解析错误
            print(f"配置文件损坏，已重置为默认配置。错误信息: {str(e)}")
            os.rename('config.json', 'config_corrupted.json')  # 备份损坏文件
            self.save_settings()
        except Exception as e:
            print(f"加载配置失败: {str(e)}")
            self.save_settings()


def is_ppt_fullscreen():
    """检测PPT或WPS是否处于全屏播放状态"""
    try:
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        process_name = process.name().lower()

        # 检测进程名称
        is_ppt = process_name == "powerpnt.exe"
        is_wps = process_name == "wpp.exe"  # WPS演示的进程名

        if is_ppt or is_wps:
            # 获取屏幕尺寸
            screen_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            screen_height = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)

            # 获取窗口尺寸
            rect = win32gui.GetWindowRect(hwnd)
            window_width = rect[2] - rect[0]
            window_height = rect[3] - rect[1]

            # 检测窗口类名（WPS特有）
            class_name = win32gui.GetClassName(hwnd)
            is_wps_fullscreen = (class_name == "KHwpsApp") and is_wps

            # 判断条件：尺寸匹配或WPS类名匹配
            return (is_wps_fullscreen or
                    (abs(window_width - screen_width) <= 1 and
                     abs(window_height - screen_height) <= 1))
        return False
    except Exception as e:
        print("检测全屏状态失败:", e)
        return False


class CountdownApp:
    def __init__(self, master):
        self.master = master
        master.title("语音提醒倒计时！")
        master.geometry("300x170+0+0")
        # 新增窗口置顶属性设置
        #master.attributes('-topmost', True)  # 关键代码
        # 正确初始化设置
        style = ttk.Style()
        style.configure('Accent.TButton', font=('微软雅黑', 10),
                        foreground='white', background='#0078d4')
        style.map('Accent.TButton', background=[('active', '#006cbd')])
        self.settings = CountdownSettings()  # 创建设置实例
        self.settings.load_settings()       # 加载设置

        self.create_widgets()
        self.create_menu()
        self.remaining = 0
        self.is_running = False
        self.engine = None
        self.black_window = None
        self.master.after(1000, self.check_ppt_status)  # 启动后持续检测PPT状态
        self.apply_background_color()  # 新增初始化颜色应用
        self.master.attributes('-alpha', self.settings.alpha)  # 应用保存的透明度
        if pyttsx3:
            try:
                self.engine = pyttsx3.init()
            except Exception as e:
                print("语音初始化失败:", e)

    def check_ppt_status(self):
        """周期性检测PPT是否全屏"""
        if is_ppt_fullscreen():
            self.start_countdown()
        else:
            self.master.after(1000, self.check_ppt_status)  # 每隔1秒检测一次

    def create_menu(self):
        menubar = tk.Menu(self.master)

        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        settings_menu.add_command(label="提醒方式设置", command=self.alert_settings)
        settings_menu.add_command(
            label="欢迎词设置", command=self.welcome_settings)  # 新增菜单项
        settings_menu.add_command(
            label="当前时间显示设置", command=self.current_time_settings)  # 新增菜单项
        settings_menu.add_command(label="显示字体设置", command=self.font_settings)
        settings_menu.add_command(
            label="背景颜色设置", command=self.bg_color_settings)  # 新增
        menubar.add_cascade(label="系统设置", menu=settings_menu)
        self.master.config(menu=menubar)
    # 新增透明度设置对话框

    def bg_color_settings(self):
        """整合颜色和透明度设置的对话框"""
        dialog = tk.Toplevel()
        dialog.title("颜色与透明度设置")
        dialog.geometry("400x300")

        # 颜色选择部分
        color_frame = ttk.LabelFrame(dialog, text="背景颜色设置", padding=10)
        color_frame.pack(pady=10, fill=tk.X, padx=10)

        # 当前颜色预览
        preview_frame = ttk.Frame(color_frame)
        preview_frame.pack(fill=tk.X)
        self.color_preview = tk.Label(
            preview_frame,
            text="当前颜色",
            width=15,
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        self.color_preview.pack(side=tk.LEFT, padx=5)

        # 颜色选择按钮
        ttk.Button(
            preview_frame,
            text="选择颜色",
            command=lambda: self.choose_color(dialog)
        ).pack(side=tk.RIGHT, padx=5)

        # 透明度调节部分
        alpha_frame = ttk.LabelFrame(dialog, text="透明度调节", padding=10)
        alpha_frame.pack(pady=10, fill=tk.X, padx=10)

        # 滑动条框架
        slider_frame = ttk.Frame(alpha_frame)
        slider_frame.pack(fill=tk.X)

        ttk.Label(slider_frame, text="透明度（0-100%）:").pack(side=tk.LEFT)
        alpha_var = tk.IntVar(value=int(self.settings.alpha*100))

        # 现代风格滑动条
        self.alpha_slider = ttk.Scale(
            slider_frame,
            from_=0,
            to=100,
            orient='horizontal',
            variable=alpha_var,
            command=lambda _: self.update_alpha_preview(alpha_var.get()/100)
        )
        self.alpha_slider.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # 输入框框架
        entry_frame = ttk.Frame(alpha_frame)
        entry_frame.pack(pady=5, fill=tk.X)

        ttk.Label(entry_frame, text="精确值（0.0-1.0）:").pack(side=tk.LEFT)
        self.alpha_entry = ttk.Entry(entry_frame, width=8)
        self.alpha_entry.insert(0, f"{self.settings.alpha:.2f}")
        self.alpha_entry.pack(side=tk.LEFT, padx=5)
        self.alpha_entry.bind("<Return>", lambda e: self.validate_alpha_input())

        # 操作按钮
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)

        ttk.Button(
            btn_frame,
            text="应用设置",
            command=lambda: self.save_all_settings(dialog, alpha_var.get()/100),
            #style='Accent.TButton'
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            btn_frame,
            text="取消",
            command=dialog.destroy
        ).pack(side=tk.RIGHT, padx=5)

        # 实时预览开关
        self.preview_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            dialog,
            text="实时预览",
            variable=self.preview_var,
            command=self.toggle_preview
        ).pack(anchor=tk.W, padx=10)


    def choose_color(self, parent):
        """颜色选择并实时更新预览"""
        color_code = colorchooser.askcolor(title="选择背景颜色", parent=parent)[1]
        if color_code:
            self.settings.background_color = color_code
            self.settings.font_color = self.get_inverted_color(color_code)
            if self.preview_var.get():
                self.color_preview.config(
                    bg=self.settings.background_color,
                    fg=self.settings.font_color
                )
                self.apply_background_color()


    def update_alpha_preview(self, value):
        """透明度滑动条实时更新"""
        if self.preview_var.get():
            self.master.attributes('-alpha', value)
        self.alpha_entry.delete(0, tk.END)
        self.alpha_entry.insert(0, f"{float(value):.2f}")


    def validate_alpha_input(self):
        """验证透明度输入"""
        try:
            value = float(self.alpha_entry.get())
            if 0.0 <= value <= 1.0:
                self.alpha_slider.set(value*100)
                if self.preview_var.get():
                    self.master.attributes('-alpha', value)
            else:
                messagebox.showerror("错误", "请输入0.0到1.0之间的数值")
        except ValueError:
            messagebox.showerror("错误", "请输入有效数字")


    def save_all_settings(self, dialog, alpha_value):
        """保存所有设置"""
        self.settings.alpha = alpha_value
        self.settings.save_settings()
        self.apply_background_color()
        self.master.attributes('-alpha', alpha_value)
        dialog.destroy()


    def toggle_preview(self):
        """切换实时预览状态"""
        if not self.preview_var.get():
            # 恢复原始设置
            self.master.attributes('-alpha', self.settings.alpha)
            self.apply_background_color()
    def update_alpha(self, value, entry_widget):
        """滑动条实时更新"""
        self.master.attributes('-alpha', value)
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, f"{value:.2f}")


    def validate_alpha(self, entry_widget, slider):
        """验证输入值有效性"""
        try:
            value = float(entry_widget.get())
            if 0.0 <= value <= 1.0:
                slider.set(int(value*100))
                self.update_alpha(value, entry_widget)
            else:
                messagebox.showerror("错误", "请输入0.0到1.0之间的数值")
        except ValueError:
            messagebox.showerror("错误", "请输入有效数字")


    def save_alpha(self, value):
        """保存透明度设置"""
        self.settings.alpha = value
        self.settings.save_settings()
        self.master.attributes('-alpha', value)
    def current_time_settings(self):
        dialog = tk.Toplevel()
        dialog.title("当前时间显示设置")

        enable_var = tk.BooleanVar(value=self.settings.show_current_time)
        tk.Checkbutton(dialog, text="显示当前时间",
                       variable=enable_var).pack(padx=20, pady=10)

        def save_settings():
            self.settings.show_current_time = enable_var.get()
            self.settings.save_settings()
            self.update_time_display()  # 立即更新显示状态
            dialog.destroy()

        tk.Button(dialog, text="保存", command=save_settings).pack(pady=10)

    # 修改字体设置对话框

    def font_settings(self):
        dialog = tk.Toplevel()
        dialog.title("高级字体设置")
        notebook = ttk.Notebook(dialog)  # 使用选项卡控件

        # 选项卡1：倒计时字体设置 -------------------------
        countdown_frame = ttk.Frame(notebook)
        self.create_font_tab(countdown_frame,
                             "倒计时显示字体:",
                             self.settings.countdown_font_name,
                             self.settings.countdown_font_size)
        notebook.add(countdown_frame, text="倒计时字体")

        # 选项卡2：当前时间字体设置 -----------------------
        current_time_frame = ttk.Frame(notebook)
        self.create_font_tab(current_time_frame,
                             "当前时间字体:",
                             self.settings.current_time_font_name,
                             self.settings.current_time_font_size)
        notebook.add(current_time_frame, text="时间显示字体")

        # 选项卡3：时间到提示字体设置 ---------------------
        time_up_frame = ttk.Frame(notebook)
        self.create_font_tab(time_up_frame,
                             "结束提示字体:",
                             self.settings.time_up_font_name,
                             self.settings.time_up_font_size)
        notebook.add(time_up_frame, text="结束提示字体")

        notebook.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)

        # 保存按钮修改为以下形式
        tk.Button(dialog, text="保存全部设置",
                  command=lambda: self.save_font_settings(
                      countdown_frame,
                      current_time_frame,
                      time_up_frame,
                      dialog=dialog  # 显式指定关键字参数
                  )).pack(pady=10)

    def create_font_tab(self, parent, label_text, cur_font, cur_size):
        """创建字体设置子面板的通用方法"""
        ttk.Label(parent, text=label_text).pack(pady=5)

        # 字体选择
        font_combo = ttk.Combobox(
            parent,
            values=["黑体", "微软雅黑", "宋体", "楷体", "仿宋",
                    "隶书", "幼圆", "方正舒体", "华文琥珀", "华文行楷"],
            state="readonly"
        )
        font_combo.set(cur_font)
        font_combo.pack(pady=5)

        # 字号选择
        size_combo = ttk.Combobox(
            parent,
            values=["12", "24", "36", "48", "60", "72", "96", "120"],
            state="readonly"
        )
        size_combo.set(str(cur_size))
        size_combo.pack(pady=5)

        # 存储控件引用
        parent.font_combo = font_combo
        parent.size_combo = size_combo

    def save_font_settings(self, *frames, dialog):
        """保存所有选项卡的设置"""
        try:
            # 倒计时字体设置
            self.settings.countdown_font_name = frames[0].font_combo.get()
            self.settings.countdown_font_size = int(frames[0].size_combo.get())

            # 当前时间字体设置
            self.settings.current_time_font_name = frames[1].font_combo.get()
            self.settings.current_time_font_size = int(
                frames[1].size_combo.get())

            # 结束提示字体设置
            self.settings.time_up_font_name = frames[2].font_combo.get()
            self.settings.time_up_font_size = int(frames[2].size_combo.get())

            self.settings.save_settings()
            self.update_all_fonts()  # 立即更新界面字体
            dialog.destroy()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的字号")
    # 更新所有字体显示

    def update_all_fonts(self):
        # 更新倒计时字体
        self.countdown_label.config(
            font=(self.settings.countdown_font_name,
                  self.settings.countdown_font_size)
        )

        # 更新时间显示字体
        self.current_time_label.config(
            font=(self.settings.current_time_font_name,
                  self.settings.current_time_font_size)
        )

        # 更新黑屏提示字体（在time_up方法中）
    def welcome_settings(self):
        dialog = tk.Toplevel()
        dialog.title("欢迎词设置")

        tk.Label(dialog, text="欢迎词内容:").grid(row=0, column=0, sticky='w')
        welcome_entry = tk.Text(dialog, width=40, height=4)
        welcome_entry.insert("end", self.settings.welcome_message)
        welcome_entry.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        enable_var = tk.BooleanVar(value=self.settings.enable_welcome)
        tk.Checkbutton(dialog, text="启用欢迎词", variable=enable_var).grid(
            row=2, column=0, sticky='w')

        def save_settings():
            self.settings.welcome_message = welcome_entry.get("1.0", "end-1c")
            self.settings.enable_welcome = enable_var.get()
            self.settings.save_settings()
            dialog.destroy()

        tk.Button(dialog, text="保存", command=save_settings).grid(
            row=3, columnspan=2, pady=10)

    def create_widgets(self):
        # 主框架
        main_frame = tk.Frame(self.master, bg=self.settings.background_color)
        main_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)

        # 时间标签（修正参数顺序）
        time_label = tk.Label(
            main_frame,
            text="倒计时时间（分:秒）:",
            font=("黑体", 10),
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        time_label.pack()

        # 输入框（修正颜色设置）
        self.time_entry = tk.Entry(
            main_frame,
            font=("黑体", 12),
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        self.time_entry.insert(0, "5:00")
        self.time_entry.pack(fill=tk.X)

        # 开始按钮（添加颜色设置）
        start_btn = tk.Button(
            main_frame,
            text="开始倒计时",
            font=("黑体", 12),
            command=self.start_countdown,
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        start_btn.pack(pady=15)

        # 倒计时显示标签（修正参数顺序）
        self.countdown_label = tk.Label(
            self.master,
            font=(self.settings.countdown_font_name,
                  self.settings.countdown_font_size),
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        self.countdown_label.pack_forget()

        # 当前时间标签（修正参数顺序）
        self.current_time_label = tk.Label(
            self.master,
            text="",
            font=(self.settings.current_time_font_name,
                  self.settings.current_time_font_size),
            bg=self.settings.background_color,
            fg=self.settings.font_color
        )
        if self.settings.show_current_time:
            self.current_time_label.pack(side=tk.BOTTOM, pady=5)

        # 底部标签（修正参数顺序）
        self.footer_label = tk.Label(
            self.master,
            text="改进建议: limeng@lit.edu.cn",
            fg="gray60",
            font=("", 9),
            bg=self.settings.background_color
        )
        # 启动时间更新
        self.update_current_time()
        # 底部提示（主界面）
        self.footer_label = tk.Label(bg=self.settings.background_color,
                                     fg="gray60",  # 保持原有特殊颜色self.master,
                                     text="改进建议: limeng@lit.edu.cn",
                                     font=("", 9))
        self.footer_label.pack(side=tk.BOTTOM, pady=5)

        self.master.protocol("WM_DELETE_WINDOW", self.on_close)

    def update_current_time(self):
        if self.settings.show_current_time:
            current_time = time.strftime("%Y-%m-%d %H:%M")
            self.current_time_label.config(text=current_time)
        self.master.after(1000, self.update_current_time)  # 每秒更新

    def update_time_display(self):
        if self.settings.show_current_time:
            self.current_time_label.pack(side=tk.BOTTOM, pady=5)
        else:
            self.current_time_label.pack_forget()

    def alert_settings(self):
        dialog = tk.Toplevel()
        dialog.title("提醒设置")
        dialog.geometry("400x170")  # 增加高度适应新布局

        # ======== 第一行：变红提醒设置 ========
        red_frame = tk.Frame(dialog)
        red_frame.grid(row=0, column=0, columnspan=3, sticky='w', pady=5)
        red_enable_var = tk.BooleanVar(value=self.settings.enable_red_alert)
        tk.Checkbutton(red_frame, text="启用",
                       variable=red_enable_var).pack(side='left')
        tk.Label(red_frame, text="变红提醒时间（分钟）:").pack(side='left', padx=5)
        red_entry = tk.Entry(red_frame, width=8)
        red_entry.insert(0, str(self.settings.alert_red_min))
        red_entry.pack(side='left')

        # ======== 第二行：语音提醒设置 ========
        voice_frame = tk.Frame(dialog)
        voice_frame.grid(row=1, column=0, columnspan=3, sticky='w', pady=5)
        voice_enable_var = tk.BooleanVar(
            value=self.settings.enable_voice_alert)
        tk.Checkbutton(voice_frame, text="启用",
                       variable=voice_enable_var).pack(side='left')

        # 语音时间设置
        tk.Label(voice_frame, text="提醒时间（秒）:").pack(side='left', padx=5)
        voice_time_entry = tk.Entry(voice_frame, width=8)
        voice_time_entry.insert(0, str(self.settings.alert_voice_sec))
        voice_time_entry.pack(side='left')

        # 语音内容设置
        tk.Label(voice_frame, text="提示内容:").pack(side='left', padx=5)
        voice_msg_entry = tk.Entry(voice_frame, width=15)
        voice_msg_entry.insert(0, self.settings.voice_message)
        voice_msg_entry.pack(side='left')

        # ======== 第三行：黑屏设置 ========
        black_frame = tk.Frame(dialog)
        black_frame.grid(row=2, column=0, columnspan=3, sticky='w', pady=5)
        black_enable_var = tk.BooleanVar(
            value=self.settings.enable_black_screen)
        tk.Checkbutton(black_frame, text="启用",
                       variable=black_enable_var).pack(side='left')

        # 黑屏时长
        tk.Label(black_frame, text="黑屏时长（秒）:").pack(side='left', padx=5)
        black_entry = tk.Entry(black_frame, width=8)
        black_entry.insert(0, str(self.settings.black_screen_sec))
        black_entry.pack(side='left')

        # 永久黑屏选项
        always_var = tk.BooleanVar(value=self.settings.always_black)
        tk.Checkbutton(black_frame, text="永久黑屏", variable=always_var,
                       command=lambda: black_entry.config(
                           state="disabled" if always_var.get() else "normal")
                       ).pack(side='left', padx=10)

        # ======== 保存按钮 ========
        def save_settings():
            try:
                # 变红提醒
                self.settings.enable_red_alert = red_enable_var.get()
                self.settings.alert_red_min = int(red_entry.get())

                # 语音提醒
                self.settings.enable_voice_alert = voice_enable_var.get()
                self.settings.alert_voice_sec = int(voice_time_entry.get())
                self.settings.voice_message = voice_msg_entry.get(
                ) if voice_msg_entry.get() else "还有XX秒结束"  # 默认内容

                # 黑屏设置
                self.settings.enable_black_screen = black_enable_var.get()
                self.settings.black_screen_sec = 0 if always_var.get() else int(
                    black_entry.get())  # 永久黑屏时忽略时长
                self.settings.always_black = always_var.get()

                self.settings.save_settings()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("输入错误", "请输入有效的数字")

        tk.Button(dialog, text="保存设置", command=save_settings,
                  width=15).grid(row=3, column=1, pady=10)

    def parse_time(self, time_str):
        """解析分:秒格式的时间"""
        try:
            if ':' not in time_str:
                raise ValueError("时间格式错误")
            m, s = time_str.split(':')
            minutes = int(m)
            seconds = int(s)
            if seconds >= 60:
                raise ValueError("秒数不能超过59")
            return minutes * 60 + seconds
        except ValueError as e:
            messagebox.showerror("输入错误", f"请使用分:秒格式（例如5:00）\n{str(e)}")
            return None

    def start_countdown(self):
        # 在开始倒计时时置顶窗口
        self.master.attributes('-topmost', True)  # 新增代码
        if self.settings.enable_welcome and self.engine:
            total_mins = self.remaining // 60
            welcome_msg = self.settings.welcome_message.replace(
                "XX", str(total_mins))
            self.engine.say(welcome_msg)
            self.engine.runAndWait()
        # 解析时间输入
        time_str = self.time_entry.get()
        total_seconds = self.parse_time(time_str)
        if total_seconds is None:
            return
        self.remaining = total_seconds

        # 设置窗口尺寸
        self.master.geometry("300x150")

        # 隐藏设置控件
        for widget in self.master.winfo_children():
            if widget not in [self.countdown_label]:
                widget.pack_forget()

        # 从设置中获取字体信息
        font_name = self.settings.countdown_font_name
        font_size = self.settings.countdown_font_size

        # 配置倒计时标签
        self.countdown_label.config(
            font=(font_name, font_size),
            fg="black"
        )

        # 底部提示（主界面）
        self.footer_label = tk.Label(bg=self.settings.background_color,
                                     fg="gray60",  # 保持原有特殊颜色self.master,
                                     text="改进建议: limeng@lit.edu.cn",
                                     font=("", 9))
        self.footer_label.pack(side=tk.BOTTOM, pady=5)
        self.countdown_label.pack(expand=True, fill=tk.BOTH)

        self.is_running = True
        self.update_countdown()
        # 保持时间显示状态
        if self.settings.show_current_time:
            self.current_time_label.pack(side=tk.BOTTOM, pady=5)

    def update_countdown(self):
        '''
        # 修改语音提示逻辑
        if self.remaining == self.settings.voice_alert_time and self.engine:
            msg = self.settings.voice_message.replace(
                "XX", str(self.remaining))
            self.engine.say(msg)
            self.engine.runAndWait()
        if not self.is_running:
            return
        '''
        mins, secs = divmod(self.remaining, 60)
        self.countdown_label.config(text=f"{mins:02d}:{secs:02d}")

        # 变红提醒逻辑
        if self.remaining <= self.settings.alert_red_min * 60:
            self.countdown_label.config(fg="red")

        # 语音提醒逻辑
        if (self.settings.enable_voice_alert and
            self.remaining == self.settings.alert_voice_sec and
                self.engine):
            msg = self.settings.voice_message.replace(
                "XX", str(self.remaining))
            self.engine.say(msg)
            self.engine.runAndWait()

        if self.remaining <= 0:
            self.time_up()
            return

        self.remaining -= 1
        self.master.after(1000, self.update_countdown)

    def close_ppt():
        """终止PowerPoint或WPS进程"""
        for proc in psutil.process_iter(['name']):
            name = proc.info['name'].lower()
            if name in ("powerpnt.exe", "wpp.exe"):  # 添加WPS进程
                try:
                    proc.kill()
                except psutil.NoSuchProcess:
                    pass
                except Exception as e:
                    print("关闭失败:", e)

    def time_up(self):
        self.is_running = False
        self.master.withdraw()

        # 仅在启用黑屏时执行
        if self.settings.enable_black_screen:
            # 创建全屏黑屏窗口
            self.black_window = tk.Toplevel(self.master)
            self.black_window.attributes("-fullscreen", True)
            self.black_window.attributes("-topmost", True)
            self.black_window.config(bg="black")

            # 添加文字提示
            tk.Label(self.black_window,
                     text="时间到！",
                     fg="white",
                     bg="black",
                     font=(self.settings.time_up_font_name,  # 使用专用字体设置
                           self.settings.time_up_font_size,
                           "bold")).pack(expand=True)

            # 绑定ESC键
            self.black_window.bind(
                "<Escape>", lambda e: self.close_black_screen())

            # 永久黑屏逻辑
            if not self.settings.always_black:
                self.black_window.after(self.settings.black_screen_sec * 1000,
                                        self.close_black_screen)
        close_ppt()

    def close_black_screen(self):
        if self.black_window:
            self.black_window.destroy()
        self.master.attributes('-topmost', False)  # 新增：恢复窗口层级
        self.master.destroy()

        def close_all():
            black_window.destroy()
            self.master.destroy()
        black_window.after(5000, close_all)
        black_window.protocol("WM_DELETE_WINDOW", close_all)  # 拦截关闭事件

    def on_close(self):
        self.is_running = False
        self.master.attributes('-topmost', False)  # 新增：恢复窗口层级
        self.master.destroy()

    def choose_color(self, parent):
        """独立颜色选择方法"""
        color_code = colorchooser.askcolor(title="选择背景颜色")[1]
        if color_code:
            self.settings.background_color = color_code
            self.settings.font_color = self.get_inverted_color(color_code)
            self.apply_background_color()
    def get_inverted_color(self, hex_color):
        hex_color = hex_color.lstrip('#')
        rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        inverted = tuple(255 - x for x in rgb)
        return "#%02x%02x%02x" % inverted

    def apply_background_color(self):
        """应用颜色到所有控件"""
        # 主窗口
        self.master.config(bg=self.settings.background_color)

        # 遍历所有子控件
        for widget in self.master.winfo_children():
            try:
                widget.config(
                    bg=self.settings.background_color,
                    fg=self.settings.font_color
                )
            except:
                continue

        # 特殊处理底部标签
        self.footer_label.config(fg="gray60")
        self.master.update()
    

if __name__ == "__main__":
    root = tk.Tk()
    app = CountdownApp(root)
    root.mainloop()
