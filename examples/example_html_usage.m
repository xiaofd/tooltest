% 示例：使用 generateHtmlReport 生成 HTML 报告。
%
% 运行前请确认目标目录可写。生成的 HTML 在 IE11 及现代浏览器中均可阅读。

% 配置章节内容（段落、列表、图片、表格）
sections(1).Title = '项目概览';
sections(1).Paragraphs = {'本报告由 MATLAB 自动生成，用于展示 HTML 输出格式。', ...
    '可根据需要替换为正式内容或复制到模板中。'};
sections(1).Bullets = {'需求梳理完成', '方案评审通过', '关键里程碑已确认'};

% 使用示例 figure 作为图片（会自动导出为 PNG）
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

% 占位符示例：在段落中放置 {{project_name}} / {{dynamic_table}}
sections(2).Paragraphs = {'本节数据来自 {{project_name}}，下方为动态表格占位符：', '{{dynamic_table}}'};

% 需要替换的占位符内容
options.Placeholders.project_name = '智慧工厂试点';
options.Placeholders.dynamic_table.Header = {'部门', '负责人'};
options.Placeholders.dynamic_table.Rows = {'研发部', '李雷'; '交付部', '韩梅'};
options.Placeholders.dynamic_figures = struct('FigureHandle', {tempFig}, ...
    'Caption', {'动态插入的图例'}, 'RowIndex', {1});

% 字体与标题大小可按需调整
options.BodyFontName = 'SimSun, Arial';
options.BodyFontSize = 14;
options.HeadingFontSize = 20;

% 生成报告（请将输出路径修改为可写目录）
outputPath = fullfile(pwd, 'demo_report.html');
reportTitle = 'HTML 报告示例';

generateHtmlReport(outputPath, reportTitle, sections, options);
fprintf('HTML 报告已生成：%s\n', outputPath);

% 清理示例中创建的临时 figure
if exist('tempFig', 'var') && isgraphics(tempFig)
    close(tempFig);
end
