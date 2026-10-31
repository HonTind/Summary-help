"""Shared legacy Tk appearance; no business logic or data access."""
import tkinter as tk
from tkinter import ttk


UI_BG = "#F5F5F7"              # 窗口浅色底
UI_SURFACE = "#FFFFFF"         # 卡片表面
UI_BORDER = "#E1E1E6"          # 卡片描边与分隔线
UI_TEXT = "#1D1D1F"            # 主文字
UI_TEXT_SECONDARY = "#6E6E73"  # 次要文字
UI_TEXT_TERTIARY = "#86868B"   # 辅助说明
UI_ACCENT = "#0071E3"          # 主操作强调色
UI_ACCENT_ACTIVE = "#1A7FEB"
UI_ACCENT_PRESSED = "#0058B0"
UI_ACCENT_DISABLED = "#E5E5EA"
UI_BUTTON_SECONDARY = "#F2F2F4"
UI_BUTTON_SECONDARY_ACTIVE = "#E8E8ED"
UI_BUTTON_SECONDARY_PRESSED = "#DEDEE3"   # 与 UI_BORDER 同色，视觉上与卡片融为一体
UI_DISABLED_TEXT = "#9A9AA0"
UI_FIELD_BG = "#FAFAFB"        # 文件列表 / 运行状态的浅色底
UI_OK = "#1D7A3E"              # 状态色沿用原有语义（绿=可识别）
UI_WARN = "#B25E00"
UI_ERROR = "#C0392B"

# 字号层级（Windows 125% 缩放：10pt ≈ 13px）
UI_FONT_FAMILY = "Microsoft YaHei UI"
UI_FONT_CODE = ("Consolas", 9)

# 窗口尺寸：沿用原有最小尺寸 540×480 与原有默认尺寸 680×580。
# 新布局自然高度高于 580，由内部滚动区域保证所有控件仍可达（不提高最小尺寸）；
# 若所在屏幕的逻辑可用区域装不下默认尺寸，启动时按可用区域收敛（见 _fit_initial_geometry）。
MIN_WINDOW_WIDTH = 540
MIN_WINDOW_HEIGHT = 480
DEFAULT_WINDOW_WIDTH = 680
DEFAULT_WINDOW_HEIGHT = 580
DEFAULT_WINDOW_SIZE = f"{DEFAULT_WINDOW_WIDTH}x{DEFAULT_WINDOW_HEIGHT}"
# 客户区窄于该宽度时隐藏卡片头部/导航的辅助说明，避免标题被挤压或截断
COMPACT_LAYOUT_WIDTH = 600

# 样式/字体是按 Tk 解释器（每个 Tk 根窗口一套）存储的，因此这里记录已经配置过的
# 根窗口：同一个进程内出现第二个 Tk 根（例如销毁后重建主窗口）时必须
# 重新配置，否则新窗口的卡片会退回系统默认灰底。
_UI_STYLE_ROOT = None
# 具名字体对象必须保留引用：tkinter 的 Font 析构会删除 Tcl 侧的具名字体。
_UI_FONTS = {}
_UI_FONTS_ROOT = None
_UI_FONT_SPEC = {
    "UITitleFont": (26, "bold"),
    "UIPageTitleFont": (17, "bold"),
    "UICardTitleFont": (12, "bold"),
    "UISubtitleFont": (11, "normal"),
    "UIBodyFont": (11, "normal"),
    "UISmallFont": (10, "normal"),
    "UIBadgeFont": (10, "bold"),
}
# 语义别名：控件通过 _UI_FONT_ROLE 取字体
_UI_FONT_ROLE = {
    "title": "UITitleFont",
    "page_title": "UIPageTitleFont",
    "card_title": "UICardTitleFont",
    "subtitle": "UISubtitleFont",
    "body": "UIBodyFont",
    "small": "UISmallFont",
    "badge": "UIBadgeFont",
}


class _TkPixel:
    """ttk 的 padding 需要「pixel 对象」；直接写裸整数会被 Tk 当成像素数。

    Tk 8.6 的 Tcl 侧没有可用的 `-pixel` 构造命令，因此提供一个鸭子类型对象，
    让 tkinter 在拼接命令时输出 `<name> <value>`（例如 "UIPad2 14"），
    与 tkinter 自带 `_tkinter.Tcl_Obj` 的 `stringify` 行为一致。
    """

    def __init__(self, value):
        self._name = f"UIPad{int(value)}"
        self.string = f"{self._name} {int(value)}"

    def __str__(self):
        return self.string


def _fw(*values):
    """把整数内边距包装成 pixel 对象（ttk 样式用）。"""
    if len(values) == 1:
        return _TkPixel(values[0])
    return tuple(_TkPixel(v) for v in values)


def _ui_font_family():
    """选择可用的界面字体（Microsoft YaHei UI 优先，回退 Microsoft YaHei）。"""
    try:
        from tkinter import font as tkfont
        families = set(tkfont.families())
        for name in ("Microsoft YaHei UI", "Microsoft YaHei"):
            if name in families:
                return name
    except Exception:
        pass
    return "Microsoft YaHei UI"


def _ui_font(role):
    """返回语义角色对应的字体对象（未初始化时按需创建）。

    不用 ttk.Label：vista 主题会给 ttk 标签画与卡片不同色的浅色底，
    因此页面文字统一由 tk.Label 承载，这里返回可同时用于控件与画布的字体对象。
    """
    global _UI_FONTS_ROOT
    name = _UI_FONT_ROLE[role]
    from tkinter import font as tkfont
    root = tk._default_root
    if root is None:
        size, weight = _UI_FONT_SPEC[name]
        return (_ui_font_family(), size, weight)
    # 具名字体同样是每个 Tk 解释器一套：根窗口变化时必须重建
    if _UI_FONTS_ROOT is not root:
        _UI_FONTS.clear()
        _UI_FONTS_ROOT = root
    if name not in _UI_FONTS:
        size, weight = _UI_FONT_SPEC[name]
        try:
            _UI_FONTS[name] = tkfont.Font(root=root, name=name, family=_ui_font_family(),
                                          size=size, weight=weight)
        except tk.TclError:
            _UI_FONTS[name] = tkfont.nametofont(name, root=root)
    return _UI_FONTS[name]


def _ui_fonts(root):
    """按当前界面字体族创建全部具名字体，保证引用存活。"""
    for role in _UI_FONT_ROLE:
        _ui_font(role)
    return dict(_UI_FONTS)


def _init_ui_style(root):
    """配置本页所需的具名字体与 ttk 样式；按根窗口幂等。

    样式与具名字体都挂在 Tk 解释器（根窗口）上，因此同一个根只配置一次；
    换成另一个根（新 Tk 实例）会重新配置，避免新窗口退回系统默认灰底。

    注意：ttk 样式的 font 必须是「具名字体」（直接写 "Microsoft YaHei UI 11"
    这类字符串会被当成字体名解析失败），padding 必须是「pixel 对象」（直接写
    裸整数会被 Tk 当成像素数并触发未知选项错误）——两者都经 helper 归一化。
    """
    global _UI_STYLE_ROOT
    if _UI_STYLE_ROOT is root:
        return
    _UI_STYLE_ROOT = root

    fonts = _ui_fonts(root)
    style = ttk.Style(root)

    # 默认主题保持系统观感，仅按钮/卡片使用各主题共有的可配色元素。
    # 同时覆盖基础 TFrame：任何未显式指定 style 的 Frame 也不出现系统灰底，
    # 避免浅色页面上出现与设计色不一致的灰块。
    style.configure("TFrame", background=UI_BG)
    style.configure("Card.TFrame", background=UI_SURFACE)
    style.configure("Page.TFrame", background=UI_BG)
    style.configure("Rule.TFrame", background=UI_BORDER)

    style.configure("Secondary.TButton", background=UI_BUTTON_SECONDARY,
                    foreground=UI_TEXT, font=fonts["UIBodyFont"], borderwidth=0,
                    focuscolor=UI_BUTTON_SECONDARY, padding=_fw(14, 6))
    style.map("Secondary.TButton",
              background=[("disabled", "#F6F6F8"),
                          ("pressed", UI_BUTTON_SECONDARY_PRESSED),
                          ("active", UI_BUTTON_SECONDARY_ACTIVE)],
              foreground=[("disabled", UI_DISABLED_TEXT),
                          ("pressed", UI_TEXT),
                          ("active", UI_TEXT)])

    style.configure("Quiet.TButton", background=UI_SURFACE, foreground=UI_ACCENT,
                    font=fonts["UIBodyFont"], borderwidth=0, focuscolor=UI_SURFACE,
                    padding=_fw(10, 4))
    style.map("Quiet.TButton",
              background=[("disabled", UI_SURFACE),
                          ("pressed", "#E8F1FE"),
                          ("active", "#F0F6FF")],
              foreground=[("disabled", UI_DISABLED_TEXT),
                          ("pressed", UI_ACCENT_PRESSED),
                          ("active", UI_ACCENT_PRESSED)])

    style.configure("Card.TRadiobutton", background=UI_SURFACE, foreground=UI_TEXT,
                    font=fonts["UIBodyFont"], padding=_fw(0, 2))
    style.map("Card.TRadiobutton",
              background=[("active", UI_SURFACE)],
              foreground=[("disabled", UI_DISABLED_TEXT)])


def _ui_label(parent, text, font_role, foreground, background, **kwargs):
    """页面文字统一使用 tk.Label。

    不用 ttk.Label：vista 主题会给标签画一块与卡片不同的浅色底，
    在白色卡片上形成可见色块。tk.Label 的背景/前景完全由我们指定。
    """
    return tk.Label(parent, text=text, font=_ui_font(font_role),
                    foreground=foreground, background=background, **kwargs)


def _ui_card(parent, padding=(16, 10)):
    """白色卡片容器：1px 描边 + 内部留白，替代 LabelFrame 作为分步容器。

    返回 (外框, 内层卡片)：pack/grid 作用于外框，权重配置作用于内层卡片。
    """
    border = tk.Frame(parent, background=UI_BORDER, borderwidth=0, highlightthickness=0)
    card = ttk.Frame(border, style="Card.TFrame", padding=padding)
    card.pack(fill="both", expand=True, padx=1, pady=1)
    card.columnconfigure(0, weight=1)
    return border, card


def _ui_rule(parent, background=UI_SURFACE):
    """1px 水平分隔线。"""
    rule = tk.Frame(parent, height=1, background=background, borderwidth=0,
                    highlightthickness=0)
    return rule


def _ui_badge(parent, number):
    """步骤序号圆形徽标。"""
    size = 24
    canvas = tk.Canvas(parent, width=size, height=size, bg=UI_SURFACE,
                       highlightthickness=0, borderwidth=0, takefocus=0)
    canvas.create_oval(1, 1, size - 1, size - 1, fill=UI_ACCENT, outline="")
    canvas.create_text(size // 2, size // 2 + 1, text=str(number), fill="#FFFFFF",
                       font=_ui_font("badge"))
    return canvas


def _ui_card_header(card, number, title, hint=""):
    """卡片头部：序号徽标 + 标题 + 右侧辅助说明（窄窗口下自动隐藏说明）。"""
    header = ttk.Frame(card, style="Card.TFrame")
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(1, weight=1)
    _ui_badge(header, number).grid(row=0, column=0, padx=(0, 10))
    _ui_label(header, title, "card_title", UI_TEXT, UI_SURFACE).grid(
        row=0, column=1, sticky="w")
    hint_label = None
    if hint:
        hint_label = _ui_label(header, hint, "small", UI_TEXT_TERTIARY, UI_SURFACE)
        hint_label.grid(row=0, column=2, sticky="e", padx=(10, 0))
    # 交给 _bind_responsive_layout：窗口变窄时优先隐藏这些非关键说明
    card._ui_hints = list(getattr(card, "_ui_hints", []))
    if hint_label is not None:
        card._ui_hints.append(hint_label)
    return header


def _ui_card_body(card, row=1, pady=(12, 0)):
    """卡片内容区（头部下方，占满整行）。"""
    body = ttk.Frame(card, style="Card.TFrame")
    body.grid(row=row, column=0, sticky="ew", pady=pady)
    body.columnconfigure(0, weight=1)
    return body


class _AccentButton(tk.Frame):
    """主操作按钮：填充强调色 + 白色文字，可用鼠标与键盘操作。

    系统的 vista 主题用 MSCTL 元素绘制 ttk.Button，样式里配置的背景色不会
    生效（只见边框、文字与底色同化），因此主按钮用 tk 控件自绘，
    对外保持与 ttk.Button 一致的 config/cget(text, state) 接口。

    键盘语义与 ttk.Button 一致：Tab 可聚焦（takefocus），Space / Return
    激活，禁用态不触发，聚焦时显示强调色焦点环。
    """

    def __init__(self, parent, text, command, background=UI_ACCENT,
                 active_background=UI_ACCENT_ACTIVE, **kwargs):
        super().__init__(parent, background=background, borderwidth=0,
                         takefocus=1, highlightthickness=3,
                         highlightbackground=background,
                         highlightcolor=UI_ACCENT if background != UI_ACCENT else "#FFFFFF",
                         **kwargs)
        self._command = command
        self._background = background
        self._active_background = active_background
        self._state = "normal"
        self._text_font = _ui_font("body")
        self._label = tk.Label(self, text=text, font=self._text_font,
                               background=background, foreground="#FFFFFF",
                               cursor="hand2", padx=24, pady=10)
        self._label.pack(fill="both", expand=True)
        self._label.bind("<Enter>", self._on_enter)
        self._label.bind("<Leave>", self._on_leave)
        self._label.bind("<Button-1>", self._on_press)
        self._label.bind("<ButtonRelease-1>", self._on_release)
        # 键盘：Tab 聚焦后可用 Space / Return 激活
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<space>", self._on_key_activate)
        self.bind("<Return>", self._on_key_activate)
        self.bind("<KP_Enter>", self._on_key_activate)
        self.bind("<Destroy>", self._on_destroy, add="+")
        self._pending_release = None

    # ---- 事件 ----
    def _on_enter(self, _event=None):
        if self._state == "normal":
            self._set_background(self._active_background)

    def _on_leave(self, _event=None):
        self._set_background(self._background)

    def _on_press(self, _event=None):
        if self._state == "normal":
            self._set_background(UI_ACCENT_PRESSED)

    def _on_release(self, event=None):
        if self._state != "normal":
            return
        self._set_background(self._active_background)
        inside = (0 <= event.x < self._label.winfo_width()
                  and 0 <= event.y < self._label.winfo_height())
        if inside and self._command is not None:
            self._command()

    def _on_focus_in(self, _event=None):
        """焦点可见反馈：强调色焦点环。"""
        if self._state == "normal":
            tk.Frame.configure(self, highlightbackground=UI_ACCENT)

    def _on_focus_out(self, _event=None):
        tk.Frame.configure(self, highlightbackground=self._background)

    def _on_key_activate(self, _event=None):
        """Space / Return 激活；禁用态不触发。"""
        if self._state != "normal":
            return "break"
        self._set_background(UI_ACCENT_PRESSED)
        self._pending_release = self.after(
            60, lambda: self._set_background(self._active_background))
        self.invoke()
        return "break"

    def _on_destroy(self, _event=None):
        """控件销毁时撤销待执行的按压反馈，避免 after 回调操作已销毁的窗口。"""
        token = getattr(self, "_pending_release", None)
        if token is not None:
            self._pending_release = None
            try:
                self.after_cancel(token)
            except (tk.TclError, ValueError):
                pass

    def _set_background(self, color):
        tk.Frame.configure(self, background=color)
        self._label.configure(background=color)
        try:
            focused = self.focus_get() is self
        except (KeyError, tk.TclError):
            focused = False
        if not focused:
            tk.Frame.configure(self, highlightbackground=color)

    # ---- 兼容 ttk.Button 的接口 ----
    def config(self, **kwargs):
        """支持 config(text=..., state=..., command=..., background=...) 等常用选项。"""
        if "text" in kwargs:
            self._label.configure(text=kwargs.pop("text"))
        if "command" in kwargs:
            self._command = kwargs.pop("command")
        if "state" in kwargs:
            self.set_state(kwargs.pop("state"))
        if "background" in kwargs:
            self._background = kwargs.pop("background")
            if self._state == "normal":
                self._set_background(self._background)
        if "activebackground" in kwargs:
            self._active_background = kwargs.pop("activebackground")
        if kwargs:
            tk.Frame.configure(self, **kwargs)

    configure = config

    def cget(self, key):
        if key == "text":
            return self._label.cget("text")
        if key == "state":
            return self._state
        if key == "background":
            return self._background
        return tk.Frame.cget(self, key)

    def __getitem__(self, key):
        """兼容 ttk.Button 的 button["state"] / button["text"] 取值方式。"""
        return self.cget(key)

    def keys(self):
        return list(tk.Frame.keys(self)) + ["text", "state"]

    def set_state(self, state):
        """normal / disabled：禁用时压暗配色并去掉可点击观感。"""
        self._state = str(state)
        if self._state == "normal":
            self._set_background(self._background)
            self._label.configure(foreground="#FFFFFF", cursor="hand2")
        else:
            self._set_background(UI_ACCENT_DISABLED)
            self._label.configure(foreground=UI_DISABLED_TEXT, cursor="arrow")

    def invoke(self):
        if self._state == "normal" and self._command is not None:
            self._command()


def _ui_primary_button(parent, text, command):
    """主操作按钮（强调色填充的自绘按钮）。"""
    return _AccentButton(parent, text, command)


def _ui_secondary_button(parent, text, command, width=None):
    """次要操作按钮（浅灰填充）。"""
    btn = ttk.Button(parent, text=text, command=command, style="Secondary.TButton",
                     cursor="hand2")
    if width:
        btn.configure(width=width)
    return btn


def _ui_quiet_button(parent, text, command):
    """轻量按钮（卡片表面 + 强调色文字），用于卡片内的低优先级入口。"""
    return ttk.Button(parent, text=text, command=command, style="Quiet.TButton",
                      cursor="hand2")


def _ui_radio(parent, text, variable, value):
    """输出位置等互斥选项（vista 主题的 TRadiobutton 同样不接受 -font）。"""
    return ttk.Radiobutton(parent, text=text, variable=variable, value=value,
                           style="Card.TRadiobutton")


def _ui_text_field(parent, **kwargs):
    """只读文本区：浅色底、无色边框，与卡片融为一体。"""
    return tk.Text(
        parent, relief="flat", borderwidth=0, highlightthickness=0,
        background=UI_FIELD_BG, font=UI_FONT_CODE, **kwargs
    )
