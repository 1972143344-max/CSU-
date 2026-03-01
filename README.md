
一个基于 `Python + Tkinter + pandas` 的桌面小工具，用于从学校课表 Excel 中快速查询指定时间段的空教室。

这个项目的基本特点：
- 有实际应用场景（教学楼空教室查询）
- 有完整桌面 GUI 交互
- 有数据解析、规则匹配、结果导出流程

## 功能说明

- 导入课表 Excel（`.xlsx` / `.xls`）
- 按星期、节次、当前周查询空教室
- 支持周次范围与单双周规则匹配
- 支持导出当天各节次空教室为文本文件
- 支持查看整上午（1-4节）与整下午（5-8节）的可用教室交集

## 技术栈

- Python 3.10+
- Tkinter（GUI）
- pandas（Excel 数据处理）
- openpyxl / xlrd（Excel 引擎）
- PyInstaller（可执行文件打包）

## 项目结构

```text
.
├─ free_classroom.py        # 主程序（GUI + 查询逻辑）
├─ run.py                   # 启动入口
├─ free_classroom.spec      # PyInstaller 配置
├─ app.ico                  # 应用图标
└─ 空教室查询/               # 示例数据与历史导出（建议按需保留）
```

## 本地运行

1. 安装依赖

```bash
pip install -r requirements.txt
```

2. 运行程序

```bash
python run.py
```

## 打包为 EXE（可选）

```bash
pyinstaller free_classroom.spec
```

## 项目亮点

- 独立开发桌面工具，完成“课表解析-规则过滤-结果导出”全流程。
- 设计周次范围与单双周匹配逻辑，提升空教室查询准确性。
- 通过 GUI 降低使用门槛，使非技术用户可直接操作。
- 支持按节次与整时段交集查询，满足多种教学场景。

## 演示截图

<img width="758" height="717" alt="image" src="https://github.com/user-attachments/assets/0dafc31c-432f-4ecc-a03a-ce095499eeaf" />


## 注意事项

- 请勿上传包含隐私信息的真实课表数据（可做脱敏后再提交）。
- 若 `.xls` 文件读取失败，请确认已安装 `xlrd` 且版本兼容。
