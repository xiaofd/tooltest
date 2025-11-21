# MATLAB Word 报告生成模板

该仓库提供一组兼容 MATLAB 7 及以上版本的基础脚本，通过 Windows 上的 ActiveX/COM 自动化生成专业的 Microsoft Word 报告。

## 功能概览
- 通过 `actxserver('Word.Application')` 控制 Word，保证在 MATLAB 7 时代的语法兼容性。
- 支持封面页、章节标题、正文段落、项目符号列表和简单表格。
- 可选页脚文字、页码、作者与公司元数据，支持基于模板 `.dot/.dotx` 开始文档。

## 快速开始
1. 确保在 Windows 环境并安装了 Microsoft Word。
2. 将仓库中的 `report_generator.m` 添加到 MATLAB 路径。
3. 根据 `examples/example_usage.m` 构造章节与选项，运行脚本即可生成 `.doc` 报告。

```matlab
% 片段示例
sections(1).Title = 'Overview';
sections(1).Paragraphs = {'Project background', 'Goals'};
options.Author = 'Auto Script';
outputPath = 'C:\\temp\\demo_report.doc';
generateWordReport(outputPath, 'Demo Report', sections, options);
```

## 兼容性注意
- 代码仅使用 `invoke`/`get`/`set`，避免类定义等现代特性，以保持 MATLAB 7 可运行。
- 若使用模板文件，请确保路径有效且 Word 能够访问对应格式。

## 目录说明
- `report_generator.m`：核心函数，负责生成报告。
- `examples/example_usage.m`：示例脚本，可直接运行体验效果。
