%EXAMPLE_USAGE Build a sample report using generateWordReport.
%
% Update the output path before running. Word must be installed on Windows.

% Configure sections
sections(1).Title = '项目概览';
sections(1).Paragraphs = {'本报告由MATLAB自动生成，用于展示样例格式。', ...
    '可根据需要替换为正式内容。'};
sections(1).Bullets = {'需求梳理完成', '方案评审通过', '关键里程碑已确认'};

sections(2).Title = '数据汇总';
sections(2).Tables(1).Header = {'指标', '取值'};
sections(2).Tables(1).Rows = { ...
    '吞吐量', '24 req/s'; ...
    '响应时间', '120 ms'; ...
    '错误率', '0.01%'};

% Optional configuration
options.Author = '自动化脚本';
options.Company = '示例团队';
options.FooterText = '保密 - 内部使用';
options.AddPageNums = true;

% Generate report (update outputPath to a valid location on your machine)
outputPath = fullfile(pwd, 'demo_report.doc');
reportTitle = '自动化示例报告';

generateWordReport(outputPath, reportTitle, sections, options);
fprintf('报告已生成：%s\n', outputPath);
