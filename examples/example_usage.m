% 示例：使用 generateWordReport 构建示例报告。
%
% 运行前请确认 Word 已安装，并根据实际情况调整输出与模板路径。

% 配置章节内容（段落、列表、图片、表格）
sections(1).Title = '项目概览';
sections(1).Paragraphs = {'本报告由MATLAB自动生成，用于展示样例格式。', ...
    '可根据需要替换为正式内容。'};
sections(1).Bullets = {'需求梳理完成', '方案评审通过', '关键里程碑已确认'};
tempFig = figure('Visible', 'off');
plot(1:10, rand(1, 10));
title('动态生成的图形');
sections(1).Figures = struct('Path', {tempFig}, ...
    'Caption', {'示例图片（自动导出）'}, 'RowIndex', {1});

sections(2).Title = '数据汇总';
sections(2).Tables(1).Header = {'指标', '取值'};
sections(2).Tables(1).Rows = { ...
    '吞吐量', '24 req/s'; ...
    '响应时间', '120 ms'; ...
    '错误率', '0.01%'};

% 占位符示例：可在 Word 模板中放置 {{placeholder_name}} 并在此替换
options.Placeholders.project_name = '智慧工厂试点';
options.Placeholders.dynamic_table.Header = {'部门', '负责人'};
options.Placeholders.dynamic_table.Rows = {'研发部', '李雷'; '交付部', '韩梅'};
options.Placeholders.dynamic_figures = struct('FigureHandle', {tempFig}, ...
    'Caption', {'动态插入的图例'}, 'RowIndex', {1});

% 可选配置：模板、页边距、行距、页眉页脚等
options.Template = 'C:\\temp\\report_template.dotx';
options.Author = '自动化脚本';
options.Company = '示例团队';
options.FooterText = '保密 - 内部使用';
options.AddPageNums = true; % 设为 false 可关闭页码
options.Margins.Top = 54;   % 单位磅，示例为 0.75 英寸
options.Margins.Bottom = 54;
options.Margins.Left = 72;  % 1 英寸
options.Margins.Right = 72;
options.LineSpacing = 1.2;  % 多倍行距

% 生成报告（请将输出路径修改为可写目录）
outputPath = fullfile(pwd, 'demo_report.doc');
reportTitle = '自动化示例报告';

generateWordReport(outputPath, reportTitle, sections, options);
fprintf('报告已生成：%s\n', outputPath);

% 清理示例中创建的临时 figure
if exist('tempFig', 'var') && isgraphics(tempFig)
    close(tempFig);
end
