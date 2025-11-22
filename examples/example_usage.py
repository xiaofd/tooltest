"""示例：使用 generate_word_report 生成 Word 报告（Python 3.8）。

运行前请确认：
- 系统为 Windows 且已安装 Microsoft Word。
- 已安装 ``pywin32``。可通过 ``pip install pywin32`` 获取。
- 将输出路径修改为可写目录。
"""
from __future__ import annotations

import os
from report_generator import generate_word_report


sections = [
    {
        "Title": "项目概览",
        "Paragraphs": [
            "本报告由 Python 自动生成，用于展示样例格式。",
            "可根据需要替换为正式内容。",
        ],
        "Bullets": ["需求梳理完成", "方案评审通过", "关键里程碑已确认"],
        "Figures": [
            {
                "Path": os.path.join(os.path.dirname(__file__), "demo_figure.png"),
                "Caption": "示例图片（请替换为实际图片路径）",
                "RowIndex": 1,
            }
        ],
    },
    {
        "Title": "数据汇总",
        "Tables": [
            {
                "Header": ["指标", "取值"],
                "Rows": [
                    ["吞吐量", "24 req/s"],
                    ["响应时间", "120 ms"],
                    ["错误率", "0.01%"],
                ],
            }
        ],
    },
]

options = {
    "Template": "C:\\temp\\report_template.dotx",  # 可选模板路径，不需要可置空
    "Author": "自动化脚本",
    "Company": "示例团队",
    "FooterText": "保密 - 内部使用",
    "AddPageNums": True,
    "Margins": {"Top": 54, "Bottom": 54, "Left": 72, "Right": 72},
    "LineSpacing": 1.2,
    "Placeholders": {
        "project_name": "智慧工厂试点",
        "dynamic_table": {
            "Header": ["部门", "负责人"],
            "Rows": [["研发部", "李雷"], ["交付部", "韩梅"]],
        },
    },
}

output_path = os.path.join(os.getcwd(), "demo_report.doc")
report_title = "自动化示例报告"

if __name__ == "__main__":
    generate_word_report(output_path, report_title, sections, options)
    print(f"报告已生成：{output_path}")
