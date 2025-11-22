# MATLAB Word 报告生成模板

该仓库提供一组兼容 MATLAB 7 及以上版本的基础脚本，通过 Windows 上的 ActiveX/COM 自动化生成专业的 Microsoft Word 报告。此外，新增了 Python 3.8 版本的实现，同样依赖 Windows + Word + COM 自动化，额外依赖仅为 `pywin32`。

## 功能概览
- 通过 `actxserver('Word.Application')` 控制 Word，保证在 MATLAB 7 时代的语法兼容性。
- 支持封面页、章节标题、正文段落、项目符号列表和简单表格。
- 可选页脚文字、页码、作者与公司元数据，支持基于模板 `.dot/.dotx` 开始文档。

## 快速开始
1. 确保在 Windows 环境并安装了 Microsoft Word。
2. 将仓库中的 `report_generator.m` 添加到 MATLAB 路径。
3. 根据 `examples/example_usage.m` 构造章节、占位符与选项，运行脚本即可生成 `.doc` 报告。

```matlab
% 基础片段示例（字段均可直接复制运行）
sections(1).Title = '项目概览';
sections(1).Paragraphs = {'项目背景', '目标'};
sections(1).Bullets = {'需求梳理完成', '方案评审通过'};

options.Template = 'C:\\temp\\report_template.dotx';
options.Author = '自动化脚本';
options.Company = '示例团队';
options.FooterText = '仅供内部审阅';
options.AddPageNums = false; % 需要页码时设为 true

% 在模板中放置 {{project_name}} / {{dynamic_table}} / {{dynamic_figures}} 即可替换
options.Placeholders.project_name = '智慧工厂试点';
options.Placeholders.dynamic_table.Header = {'部门', '负责人'};
options.Placeholders.dynamic_table.Rows = {'研发部', '李雷'; '交付部', '韩梅'};
options.Placeholders.dynamic_figures = struct('Path', {'C:\\temp\\logo.png'}, ...
    'Caption', {'动态插入的图例'}, 'RowIndex', {1});

outputPath = 'C:\\temp\\demo_report.doc';
generateWordReport(outputPath, '自动化示例报告', sections, options);
```

使用模板时请确保 `.dot/.dotx` 路径可访问；若不需要页脚或页码可分别清空 `FooterText` 或将 `AddPageNums` 设为 `false`。在 Word 模板中放置形如 `{{占位符名}}` 的标记，即可由 `Placeholders` 结构替换为动态文本、表格或图片。

## 兼容性注意
- 代码仅使用 `invoke`/`get`/`set`，避免类定义等现代特性，以保持 MATLAB 7 可运行。
- 若使用模板文件，请确保路径有效且 Word 能够访问对应格式。

## 目录说明
- `report_generator.m`：MATLAB 核心函数，负责生成 Word 报告。
- `generateHtmlReport.m`：MATLAB 版 HTML 报告生成器，输出带行内样式的 IE11 兼容 HTML。
- `examples/example_usage.m`：MATLAB 示例脚本，可直接运行体验效果。
- `examples/example_html_usage.m`：MATLAB HTML 示例脚本。
- `report_generator.py`：Python 3.8 版本，使用 `win32com` 生成报告。
- `examples/example_usage.py`：Python 示例脚本（需要安装 `pywin32`）。
