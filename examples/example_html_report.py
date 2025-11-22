"""示例：生成 HTML 报表并验证中文显示。"""
from pathlib import Path
import base64
import sys

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from report_generator import generate_html_report


# 一个 1x1 的蓝色 PNG 图像（base64 编码）
_BLUE_PIXEL = (
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9y5eKXsAAAAASUVORK5CYII="
)


def _prepare_sample_image(tmp_dir: Path) -> Path:
    image_path = tmp_dir / "blue_pixel.png"
    image_path.write_bytes(base64.b64decode(_BLUE_PIXEL))
    return image_path


def main() -> None:
    output_path = Path(__file__).with_name("example_report.html")
    tmp_dir = Path(__file__).with_name("example_html_report_assets")
    tmp_dir.mkdir(exist_ok=True)

    image_path = _prepare_sample_image(tmp_dir)

    sections = [
        {
            "Title": "简介",
            "Paragraphs": ["这是一个用于验证中文显示的 HTML 报表示例。"],
            "Bullets": ["段落、列表、表格均应正常显示。", "图片支持绝对路径或内嵌 base64。"],
        },
        {
            "Title": "数据表",
            "Tables": [
                {
                    "Header": ["项目", "数值"],
                    "Rows": [["温度", "22℃"], ["湿度", "60%"]],
                }
            ],
        },
        {
            "Title": "示例图片",
            "Figures": [
                {"Path": str(image_path), "Caption": "蓝色像素", "Embed": True},
            ],
        },
        {
            "Title": "占位符演示",
            "Paragraphs": ["{{CustomNote}}"],
        },
    ]

    options = {
        "Author": "示例用户",
        "Company": "演示公司",
        "EmbedImages": True,
        "Placeholders": {"CustomNote": "这里是通过占位符插入的自定义内容。"},
    }

    generate_html_report(str(output_path), "HTML 报表示例", sections, options)
    print(f"HTML 报表已生成: {output_path.resolve()}")


if __name__ == "__main__":
    main()
