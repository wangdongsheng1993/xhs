# 小红书蒲公英提号脚本说明

## 当前推荐用法

目前主流程推荐使用本地 Excel 版。

原因：
- 本地 Excel 版支持现有的品牌表、电商表批量处理。
- 本地 Excel 版更适合保留截图、链接、格式等结果。
- 飞书在线版可以读写文字、数字、链接，但不适合继续深挖“自动插图进单元格”。

当前统一入口是 [xhs_excel_runner.py](c:/code_20251212/AI/xhs/xhs_excel_runner.py)。

它会根据 sheet 类型自动调用：
- [xhs_extractor.py](c:/code_20251212/AI/xhs/xhs_extractor.py)：`小红书品牌-KOL`
- [xhs_ecommerce_extractor.py](c:/code_20251212/AI/xhs/xhs_ecommerce_extractor.py)：`小红书电商-KOL`

## 文件说明

- [xhs_excel_runner.py](c:/code_20251212/AI/xhs/xhs_excel_runner.py)
  统一入口，按 sheet 和行号批量调用本地 Excel 抓取脚本。
- [xhs_extractor.py](c:/code_20251212/AI/xhs/xhs_extractor.py)
  本地 Excel 品牌表处理脚本。
- [xhs_ecommerce_extractor.py](c:/code_20251212/AI/xhs/xhs_ecommerce_extractor.py)
  本地 Excel 电商表处理脚本。
- [xhs_extractor_feishu.py](c:/code_20251212/AI/xhs/xhs_extractor_feishu.py)
  飞书在线品牌表处理脚本。
- [xhs_ecommerce_extractor_feishu.py](c:/code_20251212/AI/xhs/xhs_ecommerce_extractor_feishu.py)
  飞书在线电商表处理脚本。

## 推荐运行方式

### 1. 跑品牌 sheet

```powershell
python xhs_excel_runner.py brand 395-421 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

也可以直接写中文别名：

```powershell
python xhs_excel_runner.py 品牌 395-421 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

### 2. 跑电商 sheet

```powershell
python xhs_excel_runner.py ecommerce 358-372 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

也可以直接写中文别名：

```powershell
python xhs_excel_runner.py 电商 358-372 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

### 3. 行号写法

支持这些写法：

- 连续区间：`395-421`
- 单行：`340`
- 多段混合：`338-340,395,401-405`

## 脚本行为说明

### 浏览器与登录

运行时会打开蒲公英浏览器页面，并复用本地登录态。

当前默认行为：
- 打开浏览器
- 自动等待 10 秒
- 然后继续执行

默认不再要求你手动按 Enter。

### 批量执行方式

现在 runner 已优化成“一次调用、一次浏览器会话、连续处理整段行号”。

这意味着：
- 不会再一行一行重复打开首页
- 不会每行都重新起一个子进程
- 不会因为逐行重新读原始 Excel，把前面已写入的结果覆盖掉

### 输出文件

建议始终把结果写到单独的结果文件，不直接覆盖原始表。

常用方式：
- 原始文件：`【内部深演智能】老板电器C5 提号表.xlsx`
- 结果文件：`【内部深演智能】老板电器C5 提号表_结果.xlsx`

## 常用环境变量

### 通用

- `XHS_EXCEL_PATH`
  覆盖输入 Excel 路径。
- `XHS_OUTPUT_PATH`
  覆盖输出 Excel 路径。
- `XHS_SHEET_NAME`
  覆盖目标 sheet 名。
- `XHS_DEBUG_ROW`
  指定要处理的 Excel 行号，支持单行、逗号、区间。
- `XHS_DEBUG_NAME`
  按达人名称筛选。
- `XHS_DEBUG_MAX_ROWS`
  限制最大处理条数。
- `XHS_DEBUG_VERBOSE`
  设为 `1` 时输出更详细的调试日志。

### 登录等待

- `XHS_LOGIN_WAIT_SECONDS`
  浏览器打开后自动等待多少秒再继续，默认 `10`。
- `XHS_REQUIRE_ENTER_CONFIRM`
  设为 `1` 时，恢复为“必须按 Enter 后继续”。

示例：

```powershell
$env:XHS_LOGIN_WAIT_SECONDS='5'
python xhs_excel_runner.py 品牌 395-399 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

## 品牌表处理逻辑说明

[xhs_extractor.py](c:/code_20251212/AI/xhs/xhs_extractor.py) 当前关键逻辑：

- `主页链接`
  不是输入列，而是输出列。
  脚本会先打开 `蒲公英链接`，再点击页面里的 `小红书号`，读取新打开主页的真实地址并写回。
- `ID`
  直接从蒲公英页的小红书号获取，不再从主页链接反推。
- `近30天预估阅读量`
  优先取页面/API摘要值。
  如果取不到或为 `0`，回退到“合作笔记前 4 条阅读量”的中位数。
- `近30天互动量`
  优先取页面/API摘要值。
  如果取不到或为 `0`，回退到“合作笔记前 4 条互动量”的中位数。
- 中位数口径
  排序后取中间实际值。
  偶数条时取靠前那个中间值，不取平均数。
- 占比列
  输出为百分比文本格式，保持与原表使用习惯一致。

## 电商表处理逻辑说明

[xhs_ecommerce_extractor.py](c:/code_20251212/AI/xhs/xhs_ecommerce_extractor.py) 当前关键逻辑：

- `主页链接`
  与品牌表一致，点击蒲公英页里的 `小红书号` 后读取真实主页地址写回。
- `参考案例`
  从 `合作笔记` 里优先挑大家电相关案例。
  优先级如下：
  1. 大家电相关广告笔记
  2. 大家电相关笔记
  3. 任意广告笔记
- `参考案例` 链接
  优先通过页面里的“复制小红书笔记链接”拿完整链接，因此会尽量保留 `xsec_token` 和 `xsec_source=pc_pgy` 参数。
- `KOL类型（图文/视频）`
  根据笔记案例里图文/视频数量判断。
- `平台价格`
  根据 `KOL类型（图文/视频）` 对应取 `图文笔记一口价` 或 `视频笔记一口价`。
- `近30天预估阅读量`
  取不到或为 `0` 时，回退到合作笔记前 4 条阅读量中位数。
- `近30天互动量`
  取不到或为 `0` 时，回退到合作笔记前 4 条互动量中位数。
- `达人厨房图`
  本地 Excel 版按最近内容启发式识别，有命中时可插图，未命中则写提示语。

## 飞书版现状

飞书版脚本：

- [xhs_extractor_feishu.py](c:/code_20251212/AI/xhs/xhs_extractor_feishu.py)
- [xhs_ecommerce_extractor_feishu.py](c:/code_20251212/AI/xhs/xhs_ecommerce_extractor_feishu.py)

当前适用场景：
- 在线读取飞书表格
- 在线写回文字、数值、链接
- 小范围验证单条或几行数据

当前限制：
- 不建议把飞书版作为主流程
- 飞书在线表格不适合继续研究“自动插图进单元格”
- `性别占比（截图）`、`粉丝年龄占比（截图）`、`粉丝地域（截图）` 这类截图列，飞书版只写提示文本，不在线插图

## 常见问题

### 1. 为什么日志看起来像卡住了？

常见原因：
- 浏览器正在等待蒲公英页面加载
- 正在复用登录态
- 当前达人页面响应较慢

现在默认不会卡在“按 Enter 继续”。

### 2. 为什么会反复打开首页？

旧版 runner 会逐行起脚本，导致每行都重新开首页。

现在已经修复，统一入口会一次处理整段行号。

### 3. 为什么结果会被覆盖？

旧版逐行模式会重复从原始 Excel 读数据，导致后面行覆盖前面行已写入结果。

现在已经修复，批量处理会持续累计写入结果文件。

### 4. 飞书里为什么没有截图？

因为飞书在线表格当前不适合作为截图自动插入载体。

如果需要截图类结果，请优先使用本地 Excel 版。

## 快速命令备忘

```powershell
python xhs_excel_runner.py 品牌 395-431 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

```powershell
python xhs_excel_runner.py 电商 358-372 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

## 口语化指令约定

平时直接用下面这种说法即可：

- `品牌 sheet 跑395-421`
- `品牌 sheet 跑395-431`
- `电商 sheet 跑358-372`
- `电商 sheet 跑326-329`

对应理解为：

- `品牌 sheet 跑395-421`
  等价于运行品牌表，并处理 `395-421` 行。
- `电商 sheet 跑358-372`
  等价于运行电商表，并处理 `358-372` 行。

如果需要落到命令行，分别对应：

```powershell
python xhs_excel_runner.py 品牌 395-421 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```

```powershell
python xhs_excel_runner.py 电商 358-372 --excel "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表.xlsx" --output "c:\code_20251212\AI\xhs\【内部深演智能】老板电器C5 提号表_结果.xlsx"
```
