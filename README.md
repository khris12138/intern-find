# 实习僧上海校招新增岗位筛选

这个目录用于自动筛选实习僧上海校招/全职岗位，目标是每天 11:00 找出“上次扫描以来新增”的岗位中适合社会学博士生申请的机会。

新逻辑不再依赖“24 小时内发布”。脚本每次启动都会抓取实习僧上海校招列表，优先使用列表页已经给出的岗位名、公司名、薪资、行业、公司规模和标签做快速筛选；用“岗位名 + 公司名”作为查重签名，与上次保存的扫描快照比对，只展示新增职位。

为提高速度，程序会先在列表页完成薪资/规模硬筛；只有“新增 + 通过硬筛”的岗位才打开详情页。最终输出前必须在详情页工作地点中确认包含 `上海`，因为列表页里的 `全国` 不一定真的包含上海。

当前硬筛条件：

- 已知薪资下限低于 `10000` 元/月的岗位剔除。
- 已知公司规模区间下限低于 `1000` 人的岗位剔除。
- `薪资面议` 或解析不到规模的岗位暂时保留，避免因网站字段缺失误杀。

## 文件说明

- `full_scan_once.py`
  - 主脚本。
  - 抓取实习僧列表页，并从列表页直接解析岗位名、公司、薪资、规模和标签。
  - 固定使用 `city=上海` 和 `type=school`。
  - 用“岗位名 + 公司名”与上次扫描快照比对，只展示新增岗位。
  - 默认剔除薪资下限低于 10k、公司规模下限低于 1000 人的岗位。
  - 默认只对新增且通过硬筛的岗位打开详情页，并确认工作地点包含上海。
  - 按“直接匹配”和“近似匹配”进行筛选。
  - 输出 CSV 和 JSON。

- `run_daily_scan.py`
  - 每日任务入口。
  - 调用 `full_scan_once.py` 后自动生成 Excel。
  - 默认输出到 `outputs/上海校招新增岗位筛选_YYYYMMDD_HHMMSS.xlsx`。
  - 同一天多次运行不会覆盖旧结果；如果文件已被 WPS 打开，新结果会写入新的文件名。
  - 会尝试更新 `outputs/上海校招新增岗位筛选_latest.xlsx` 作为最新副本；如果该副本被占用，不影响带时间戳的正式结果。

- `build_excel_report.py`
  - 把 `full_scan_once.py` 生成的 CSV 转成 `.xlsx`。
  - 只使用 Python 标准库，不依赖 `openpyxl`、`pandas` 或 `xlsxwriter`。

- `scripts/install_launchd.sh`
  - macOS 定时任务安装脚本。
  - 安装后每天 11:00 自动运行 `run_daily_scan.py --refresh`。
  - 日志写入 `logs/daily_scan.out.log` 和 `logs/daily_scan.err.log`。

- `prototype_scan.py`
  - 早期 3 页验证原型。
  - 主要用于保留实验过程，正式复用建议使用 `full_scan_once.py`。

- `shixiseng_school_cache.sqlite3`
  - 网页缓存。
  - 已抓过的列表页和详情页会短期保存在这里，默认保留 2 天，避免数据库无限变大。
  - `scan_state` 表会保存累计岗位签名快照，用于下一次增量比对。
  - 清理网页缓存不会清理 `scan_state`，所以不会影响每日基线判断。

- `outputs/`
  - 输出目录。
  - 每次运行 `full_scan_once.py` 会生成一个时间戳子目录。

## 每日自动运行

在本目录运行：

```bash
bash scripts/install_launchd.sh
```

安装后，macOS 会在每天 11:00 自动扫描并生成当天 Excel。每次运行都会使用新的时间戳文件名保留结果。

如需查看是否已安装：

```bash
launchctl list | grep com.khris.intern-find.daily
```

## 手动运行一次

在本目录运行：

```bash
python3 run_daily_scan.py --refresh
```

输出示例：

```text
outputs/20260606_190619/
  current_list_jobs.csv
  new_jobs.csv
  explicit_matches.csv
  approximate_matches.csv
  all_matches.csv
  matches.json
outputs/上海校招新增岗位筛选_20260606_110000.xlsx
```

含义：

- `current_list_jobs.csv`：本轮列表页看到的全部岗位，是当前查重基底的可读备份。
- `new_jobs.csv`：上次扫描以来新增的岗位；被列表页跳过或硬筛剔除的岗位也会保留原因。
- `explicit_matches.csv`：直接匹配。岗位描述中明确出现 `社会学`、`社会科学`、`人文社科`、`社科`、`用户`、`用户研究`、`用户洞察`、`用研`。
- `approximate_matches.csv`：近似匹配。没有直接出现上述词，但包含研究、调研、行业分析、政策/公益/社区/性别、AI 人文评估等可迁移能力信号。
- `all_matches.csv`：前两类合并。
- `matches.json`：同样结果的 JSON 版，方便交给其他程序或 AI 继续处理。

Excel 的岗位表只保留 6 列，并包含 `当前列表`、`直接匹配`、`近似匹配` 等 sheet：

- 岗位名称
- 公司
- 工作地点
- 薪资
- 详情链接
- 命中关键词

## 生成 Excel

扫描结束后，把时间戳目录替换成你的实际输出目录：

```bash
python3 build_excel_report.py outputs/20260606_190619 outputs/上海校招新增岗位筛选_20260606.xlsx
```

Excel 包含三个 sheet：

- `摘要`
- `直接匹配`
- `近似匹配`

## 使用缓存和快照

列表页默认每次重新抓取，确保能发现新增岗位；详情页只在新增岗位通过薪资/规模硬筛后打开，并优先读取缓存。

如果只是修改关键词、分数或分类逻辑，可以直接运行：

```bash
python3 full_scan_once.py --pages 500 --delay 0.15 --outdir outputs
```

脚本会优先读取 `shixiseng_school_cache.sqlite3` 中已有的详情页缓存。

如果今天只想把当前所有岗位登记为之后查重的基底，不想逐个打开详情页：

```bash
python3 run_daily_scan.py --reset-baseline
```

如果已经建立了今天基线，但还想把今天当前列表中的全部岗位按筛选规则跑一遍并输出 Excel，同时保留这个基线用于明天对比：

```bash
python3 run_daily_scan.py --scan-current
```

如果你希望手动指定更高页数上限：

```bash
python3 run_daily_scan.py --pages 800
```

如果想强制重新访问详情页，加：

```bash
python3 full_scan_once.py --pages 500 --delay 0.15 --outdir outputs --refresh
```

日常扫描会把本次列表中的岗位签名追加进累计基线：第一天已有的岗位不会在第二天重复出现，第二天新增并登记过的岗位也不会在第三天重复出现。

如果想重新建立增量基线，优先使用 `--reset-baseline`；它会用当前列表覆盖旧基线，并且不会把当前既有岗位当作新增职位展示。

## 修改筛选逻辑

打开 `full_scan_once.py`，重点看三处：

1. `EXPLICIT_TERMS`
   - 直接匹配关键词。
   - 例如：`社会学`、`社会科学`、`人文社科`、`用户`、`用户研究`、`用户洞察`。

2. `APPROX_RULES`
   - 近似匹配规则。
   - 格式是：

```python
("关键词", 分数, "能力类别")
```

3. `CORE_CATEGORIES`
   - 至少命中一个核心类别，近似匹配才会保留。
   - 用来防止泛词造成过多误收。

## 当前筛选 URL

脚本会自动生成和访问以下同等条件的分页 URL：

```text
https://resume.shixiseng.com/interns?page=1&type=school&keyword=&area=&months=&days=&degree=&official=&enterprise=&salary=-0&publishTime=&sortType=&city=%E4%B8%8A%E6%B5%B7&internExtend=
```
