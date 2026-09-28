# Shift Schedule Calendar

把排班 Excel 转成本地 `.ics` 日历文件。仓库只保存生成工具，不保存真实班表、姓名配置或生成后的日历文件，适合公开仓库。

## 使用方法

1. 安装依赖：

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

2. 生成 ICS：

```powershell
.\.venv\Scripts\python scripts\generate_schedule_ics.py --input "D:\path\schedule.xlsx" --name "你的姓名" --output "dist\schedule.ics"
```

如果 Excel 表头只有日期中的“日号”，并且不是当前月份，手动指定年月：

```powershell
.\.venv\Scripts\python scripts\generate_schedule_ics.py --input "D:\path\schedule.xlsx" --name "你的姓名" --year 2026 --month 9 --output "dist\schedule.ics"
```

3. 上传 `dist\schedule.ics` 到你自己的存储位置，再把上传后的 ICS 地址添加到日历订阅。

## 公开仓库注意事项

- 不要提交真实排班 Excel。
- 不要提交包含姓名的配置文件。
- 不要提交生成后的 `.ics`，除非你确认日历内容可以公开。
- `input/`、`dist/` 和 `*.ics` 已经被 `.gitignore` 忽略。

## 表格要求

- 默认读取第一个工作表。
- 前 8 行内需要有一行日期表头，至少包含 7 个日期列。
- 表头上方或表头行里最好有“姓名”或“名字”列；如果没有，脚本默认第 4 列为姓名列。
- 脚本会导出目标姓名所在行的班次代码，并尝试从下一行读取工时。
