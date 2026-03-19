# 棉花果订单工具使用文档

本项目是一套面向 `58mhg.com` 的订单处理工具，包含扫描订单、批量查询进度以及微信发货比对等流程。本文档整合模块说明与使用指引，便于统一维护与交接。

---

## 一、项目文件说明

- `scan_orders_main.py`：扫描订单 Excel（子文件夹内），生成 `orders_时间戳.xlsx`。
- `mhg_batch_query.py`：读取订单清单，批量查询订单状态并导出结果。
- `wechat_order_reconcile.py`：对接微信导出表，与订单结果进行发货比对并汇总。
- `set_password.py`：设置启动密码（生成 `auth_config.json`）。
- `build_exe.py`：将脚本打包为 exe。
- `auth_config.json`：启动密码哈希配置（需与脚本或 exe 同目录）。

---

## 二、模块说明

### 1. 界面模块说明

界面模块以命令行交互为主，负责参数输入与运行引导。主要特征如下：

- 运行时通过命令行提示用户输入文件路径、日期范围等配置，并支持拖拽 Excel 文件到终端自动带入路径。
- 支持默认值回填，便于批量运行时快速确认输入。
- 对输出结果进行颜色标注与状态提示，增强可读性。
- 界面配置参数会持久化保存（例如 `gui_config.json`），实现“上次输入自动回填”。

### 2. 数据模块说明

数据模块负责配置与运行数据的持久化管理，主要包含三类数据文件：

1) **运行记录与历史日志**：`run_history.csv` 记录每次任务执行的时间、任务类型、状态、耗时与参数，便于回溯与审计。  
2) **本地配置与安全信息**：`gui_config.json` 保存界面输入的最新参数（如路径、时间区间、开关等），用于下次启动时自动回填；`auth_config.json` 仅保存加密后的密码哈希与提示信息，避免明文存储。  
3) **接口样例与数据参考**：`referencePaper/*.json` 用于保存接口返回的结构样例（订单列表、工序进度等），作为数据解析与对照的参考资料。

---

## 三、运行环境

- Windows 10/11（推荐）
- Python 3.9+

安装依赖：

```bash
pip install openpyxl requests pandas pyinstaller
```

> 如果只运行扫描脚本，最少 `openpyxl` 即可。

---

## 四、首次使用（必须）

### 1）设置启动密码

在项目目录执行：

```bash
python set_password.py
```

按提示输入两次新密码。成功后会在当前目录生成（或覆盖）`auth_config.json`。

### 2）确认配置文件位置

请确保：

- `auth_config.json`
- 你要运行的 `py` 或 `exe`

位于同一目录。

---

## 五、步骤1：扫描订单并生成清单

运行：

```bash
python scan_orders_main.py
```

按提示输入：

1. 启动密码
2. 网盘路径（可直接回车使用当前目录）
3. 起始编号（可空）
4. 结束编号（可空）

程序会扫描目标路径下每个子文件夹中的 `.xlsm/.xlsx`，读取：

- 用户名：`F5`（或兼容 `G5`）
- 订单编号：`J2`

输出文件：

- `orders_YYYYMMDD_HHMMSS.xlsx`

输出列说明：

- `用户名`
- `订单编号(oId)`
- `基础编号`（如 `99556-1` 的基础编号为 `99556`）
- `状态`（已找到 / 缺失）

### 编号区间规则

- 若填写起止编号，会按基础编号进行过滤。
- 若同时填写了起止编号，程序会自动补齐区间缺失编号并标记 `缺失(未找到文件)`。

---

## 六、步骤2：批量查询订单进度

运行前先编辑 `mhg_batch_query.py` 顶部 `CONFIG`：

- `cookie`：从浏览器复制最新 Cookie（过期需更新）
- `input_excel`：上一步生成的 `orders_*.xlsx` 路径
- `output_excel`：输出结果文件名
- `request_delay`：请求间隔（建议保留 1 秒左右）
- `start_date` / `end_date`：查询日期范围（建议覆盖订单时间）

运行：

```bash
python mhg_batch_query.py
```

执行时可再次输入/确认输入与输出 Excel 路径（支持拖拽文件到终端）。

输出为订单状态结果表（含颜色标识）：

- 红：查无此单
- 蓝：待开始
- 黄：生产中
- 绿：已完成

---

## 七、步骤3：微信发货比对（可选）

可选运行微信发货比对，输出核对结果并写入汇总库。

运行示例：

```bash
python wechat_order_reconcile.py --log data/wechat_shipment_log.xlsx --output data/reconcile_result.xlsx --wechat data/群聊_单号群.xlsx --orders data/orders_result.xlsx
```

主要作用：

- 将微信导出表与订单查询结果做匹配
- 标记已发货、未发货、异常项
- 可将结果写入历史汇总库便于持续追踪

---

## 八、打包为 EXE（可选）

执行：

```bash
python build_exe.py
```

成功后在 `dist/` 下生成：

- `scan_orders.exe`（或带时间戳名称）

将 exe 与 `auth_config.json` 放在同一目录后，可在无 Python 环境电脑上双击运行。

---

## 九、常见问题

### 1）提示“未读取到密码配置”

原因：未找到同目录 `auth_config.json`。

处理：重新运行 `set_password.py`，并把生成文件放到脚本/exe同目录。

### 2）扫描不到订单

请检查：

- 目标路径下是否是“按订单分子文件夹”结构
- 子文件夹内是否有 `.xlsm/.xlsx`
- Excel 对应单元格（`J2`,`F5/G5`）是否有值
- 起止编号是否限制过严

### 3）批量查询返回异常或查不到

请检查：

- Cookie 是否过期（最常见）
- 日期范围是否覆盖目标订单
- `input_excel` 路径是否正确
- 网络是否可访问 `58mhg.com`

### 4）打包失败

可能旧 exe 正在运行/被占用。先关闭后重试，或使用脚本自动回退的时间戳 exe 名称。

---

## 十、推荐使用流程（简版）

1. `python set_password.py`
2. `python scan_orders_main.py`（生成 `orders_*.xlsx`）
3. 更新 `mhg_batch_query.py` 的 `CONFIG`（Cookie + input_excel）
4. `python mhg_batch_query.py`
5. （可选）`python wechat_order_reconcile.py`
6. 查看导出的结果 Excel
