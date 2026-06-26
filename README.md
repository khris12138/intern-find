# 实习僧上海社会学相关实习筛选

这个目录用于自动筛选实习僧上海实习岗位，目标是每天 11:00 找出“1天内发布、base 上海”的岗位中适合社会学博士生申请的机会。

## 文件说明

- `full_scan_once.py`
  - 主脚本。
  - 抓取实习僧列表页和岗位详情页。
  - 固定使用 `city=上海` 和 `publishTime=day`。
  - 按“直接匹配”和“近似匹配”进行筛选。
  - 输出 CSV 和 JSON。

- `run_daily_scan.py`
  - 每日任务入口。
  - 调用 `full_scan_once.py` 后自动生成 Excel。
  - 默认输出到 `outputs/上海社会学相关实习筛选_YYYYMMDD_HHMMSS.xlsx`。
  - 同一天多次运行不会覆盖旧结果；如果文件已被 WPS 打开，新结果会写入新的文件名。
  - 会尝试更新 `outputs/上海社会学相关实习筛选_latest.xlsx` 作为最新副本；如果该副本被占用，不影响带时间戳的正式结果。

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

- `shixiseng_cache.sqlite3`
  - 网页缓存。
  - 已抓过的列表页和详情页会保存在这里，后续改关键词规则时可以不重新访问网站。

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
  explicit_matches.csv
  approximate_matches.csv
  all_matches.csv
  matches.json
outputs/上海社会学相关实习筛选_20260606_110000.xlsx
```

含义：

- `explicit_matches.csv`：直接匹配。岗位描述中明确出现 `社会学`、`社会科学`、`人文社科`、`社科`、`用户`、`用户研究`、`用户洞察`、`用研`。
- `approximate_matches.csv`：近似匹配。没有直接出现上述词，但包含研究、调研、行业分析、政策/公益/社区/性别、AI 人文评估等可迁移能力信号。
- `all_matches.csv`：前两类合并。
- `matches.json`：同样结果的 JSON 版，方便交给其他程序或 AI 继续处理。

## 生成 Excel

扫描结束后，把时间戳目录替换成你的实际输出目录：

```bash
python3 build_excel_report.py outputs/20260606_190619 outputs/上海社会学相关实习筛选_20260606.xlsx
```

Excel 包含三个 sheet：

- `摘要`
- `直接匹配`
- `近似匹配`

## 使用缓存重算

如果只是修改关键词、分数或分类逻辑，不想重新抓网页，可以直接运行：

```bash
python3 full_scan_once.py --pages 50 --delay 0.15 --outdir outputs
```

脚本会优先读取 `shixiseng_cache.sqlite3`。

如果想强制重新访问网站，加：

```bash
python3 full_scan_once.py --pages 50 --delay 0.15 --outdir outputs --refresh
```

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
https://www.shixiseng.com/interns?page=1&type=intern&keyword=&area=&months=&days=&degree=&official=&enterprise=&salary=-0&publishTime=day&sortType=&city=%E4%B8%8A%E6%B5%B7&internExtend=
```
