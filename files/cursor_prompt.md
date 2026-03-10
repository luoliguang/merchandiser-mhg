# 棉花果订单工作台 GUI 重构指令

## 目标
将现有 PyQt GUI 界面重构为工业风格的深色界面，参考以下设计规范，**不要 AI 感，不要圆角卡片堆叠风格**。

---

## 整体布局结构

```
┌─────────────────────────────────────────────┐
│  标题栏（窗口控制 + 应用名 + 语言/主题切换）     │
├──────────┬──────────────────────────────────┤
│          │  状态栏（总任务 / 成功 / 失败 / 进度条）│
│  侧边栏   ├──────────────────────────────────┤
│          │                                  │
│  · 订单查询│  文件配置区（输入/输出路径）          │
│  · 微信核对│                                  │
│  · 设置   │  操作按钮行                        │
│          │                                  │
│          │  日志区（左：运行日志 | 右：运行历史）  │
└──────────┴──────────────────────────────────┘
```

---

## 颜色规范

```python
# 全部使用 QSS 变量或直接硬编码以下色值
BG        = "#0f1117"   # 窗口背景
SURFACE   = "#181c27"   # 面板/侧边栏背景
SURFACE2  = "#1e2333"   # 悬停/次级背景
BORDER    = "#2a3045"   # 普通边框
BORDER2   = "#353d58"   # 强调边框
TEXT      = "#c8d0e8"   # 主文字
TEXT_DIM  = "#5a6480"   # 次要文字
TEXT_MUTE = "#3a4260"   # 占位/标签文字
ACCENT    = "#4e9eff"   # 主色（蓝）
ACCENT2   = "#2d7dd6"   # 主色深（按钮背景）
GREEN     = "#3ecf6e"   # 成功
RED       = "#e05555"   # 失败/错误
YELLOW    = "#e0a830"   # 警告
```

---

## 字体规范

```python
# 主字体：IBM Plex Sans SC（中文+界面文字）
# 等宽字体：IBM Plex Mono（数字、路径、日志、标签）
# 字号：
#   标签/说明文字  → 11~12px
#   按钮          → 12~13px
#   日志内容      → 11.5px
#   统计数字      → 18px bold
#   区块标题      → 10px 全大写 letter-spacing 0.08em
```

---

## 各区域详细规范

### 1. 标题栏
- 高度 38px，背景 SURFACE，底部 1px BORDER
- 左侧：macOS 风格三色圆点（纯装饰，11px）+ 应用名（等宽字体，TEXT_DIM）
- 右侧：Language 和 Theme 切换按钮（小型，边框按钮风格）

### 2. 侧边栏
- 宽度 180px，背景 SURFACE，右侧 1px BORDER
- 顶部 Logo 区：主标题 ACCENT 色等宽字体 + 副标题全大写小字
- 导航项：左侧 2px 竖线指示当前页（ACCENT 色），hover 背景 SURFACE2
- 区块标题（"功能"/"系统"）：10px 全大写，TEXT_MUTE 色

### 3. 状态栏
- 高度 44px，背景 SURFACE，底部 1px BORDER
- 三个统计项横排：总任务（TEXT）/ 成功（GREEN）/ 失败（RED）
- 每项：小标签 + 18px 数字，项之间用 1px BORDER 竖线分隔
- 右侧：进度条（4px 高，160px 宽）+ 百分比文字 + 重试按钮

### 4. 文件配置区
- 背景 SURFACE，边框 1px BORDER，圆角 6px，内边距 14px 16px
- 顶部区块标题："文件配置"（10px 全大写）
- 两行输入：标签（52px 宽等宽字体）+ 路径输入框（只读，等宽字体）+ 浏览按钮
- 底部复选框行：checkbox + 说明文字 + 右侧文件名标签（小徽章样式）
- 输入框：背景 BG，边框 BORDER2，focus 时边框变 ACCENT

### 5. 操作按钮行
- 主按钮（运行订单查询）：flex:1，背景 ACCENT2，边框 ACCENT，白色文字
- 次级按钮（打开输出目录）：背景 SURFACE2，边框 BORDER2，TEXT_DIM 文字
- 按钮高度统一，圆角 5px

### 6. 日志区（重点，占剩余全部高度）

**左侧：运行日志（flex:1）**
- 背景 SURFACE，边框 BORDER，圆角 6px
- 顶部 header：10px 全大写标题 + 右侧状态圆点（运行时绿色发光）
- 日志内容：等宽字体 11.5px，行高 1.7，可滚动
- 每行格式：`时间戳（TEXT_MUTE）` + `消息内容`
- 颜色规则：
  - 普通信息 → TEXT
  - 成功 ✓   → GREEN
  - 警告 ⚠   → YELLOW
  - 错误 ✗   → RED
- 滚动条：4px 细条，BORDER2 色

**右侧：运行历史（固定宽度 220px）**
- 相同面板样式
- 每条历史项：时间（TEXT_MUTE）+ 成功数（GREEN）/ 失败数（RED）/ 总数（TEXT_DIM）
- 底部固定：「复用所选记录」按钮（次级样式，全宽）

---

## PyQt 实现要点

```python
# 1. 窗口去掉默认标题栏
self.setWindowFlags(Qt.FramelessWindowHint)

# 2. 整体背景色
self.setStyleSheet("QMainWindow { background: #0f1117; }")

# 3. 布局结构
# QHBoxLayout (主体)
#   ├── SidebarWidget (固定宽 180px)
#   └── QVBoxLayout (内容区)
#         ├── StatusBarWidget (固定高 44px)
#         └── QVBoxLayout (main)
#               ├── FileConfigWidget
#               ├── ButtonRowWidget
#               └── QHBoxLayout (日志区，stretch=1)
#                     ├── LogPanel (stretch=1)
#                     └── HistoryPanel (固定宽 220px)

# 4. 日志颜色用 QTextEdit + HTML 插入
def append_log(self, msg, level='info'):
    colors = {'info': '#c8d0e8', 'ok': '#3ecf6e', 'warn': '#e0a830', 'err': '#e05555'}
    ts = datetime.now().strftime('%H:%M:%S')
    html = f'<span style="color:#3a4260">{ts}</span> <span style="color:{colors[level]}">{msg}</span>'
    self.log_text.append(html)

# 5. 进度条用 QProgressBar + 自定义 QSS
# chunk 背景色 ACCENT，track 背景色 BORDER，高度 4px

# 6. 侧边栏导航选中态
# 用 QPushButton + QSS :checked 或 property 切换
# border-left: 2px solid #4e9eff; background: rgba(78,158,255,0.07);
```

---

## 不要做的事
- ❌ 不要用 QTabWidget 默认样式做 Tab 切换
- ❌ 不要用 QGroupBox 默认边框
- ❌ 不要用系统默认滚动条样式
- ❌ 不要用渐变背景或发光阴影（除状态圆点外）
- ❌ 不要用圆角超过 6px 的卡片
- ❌ 不要用 Inter / Roboto / 微软雅黑 等通用字体

