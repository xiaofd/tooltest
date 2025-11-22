function generateHtmlReport(outputPath, reportTitle, sections, options)
%GENERATEHTMLREPORT 生成 HTML 报告（兼容 MATLAB 7 语法）。
%   generateHtmlReport(outputPath, reportTitle, sections, options) 会根据
%   与 generateWordReport 相同的数据结构生成带行内样式的 HTML 文件，
%   支持段落、列表、表格、图片与占位符替换，确保 IE11 中文显示正常。
%
%   参数
%   ----
%   outputPath : char
%       目标 HTML 文件完整路径（包含 .html）。
%   reportTitle : char
%       报告主标题。
%   sections : struct array
%       与 generateWordReport 相同的章节结构，包含 Title、Paragraphs、
%       Bullets、Tables、Figures 字段。
%   options : struct (optional)
%       可包含：
%         - BodyFontName   : 字体族，默认 'SimSun, Arial'。
%         - BodyFontSize   : 正文字号（px），默认 14。
%         - HeadingFontSize: 标题字号（px），默认 20。
%         - Placeholders   : 名称到文本/表格/图片结构的映射。
%
%   示例详见 examples/example_html_usage.m。

if nargin < 4
    options = struct;
end
if nargin < 3
    sections = [];
end

if ~isfield(options, 'BodyFontName')
    options.BodyFontName = 'SimSun, Arial';
end
if ~isfield(options, 'BodyFontSize')
    options.BodyFontSize = 14;
end
if ~isfield(options, 'HeadingFontSize')
    options.HeadingFontSize = 20;
end

baseDir = fileparts(outputPath);
if isempty(baseDir)
    baseDir = pwd;
end

bodyStyle = ['font-family: ' options.BodyFontName '; font-size: ' ...
    num2str(options.BodyFontSize) 'px; line-height: 1.6; word-break: break-all;'];
headingStyle = ['font-family: ' options.BodyFontName '; font-size: ' ...
    num2str(options.HeadingFontSize) 'px; word-break: break-all;'];
listStyle = ['font-family: ' options.BodyFontName '; font-size: ' ...
    num2str(options.BodyFontSize) 'px; word-break: break-all; margin: 0 0 12px 20px;'];
paragraphStyle = ['font-family: ' options.BodyFontName '; font-size: ' ...
    num2str(options.BodyFontSize) 'px; margin: 0 0 10px 0; word-break: break-all;'];

tableStyle = ['border-collapse: collapse; width: 100%; font-family: ' ...
    options.BodyFontName '; font-size: ' num2str(options.BodyFontSize) ...
    'px; word-break: break-all; margin: 0 0 12px 0;'];
tdStyle = 'border: 1px solid #999; padding: 6px; vertical-align: top;';
figureStyle = ['margin: 0 0 12px 0; font-family: ' options.BodyFontName ...
    '; font-size: ' num2str(options.BodyFontSize) 'px; word-break: break-all;'];
imageStyle = 'max-width: 100%; border: 1px solid #ccc;';

htmlParts = {};
idx = 1;
htmlParts{idx} = '<!DOCTYPE html>'; idx = idx + 1;
htmlParts{idx} = '<html>'; idx = idx + 1;
htmlParts{idx} = '<head>'; idx = idx + 1;
htmlParts{idx} = '<meta http-equiv="X-UA-Compatible" content="IE=edge">'; idx = idx + 1;
htmlParts{idx} = '<meta charset="UTF-8">'; idx = idx + 1;
htmlParts{idx} = ['<title>' escapeHtml(reportTitle) '</title>']; idx = idx + 1;
htmlParts{idx} = '</head>'; idx = idx + 1;
htmlParts{idx} = ['<body style="margin: 20px; ' bodyStyle '">']; idx = idx + 1;
htmlParts{idx} = ['<h1 style="margin: 0 0 14px 0; ' headingStyle '">' ...
    escapeHtml(reportTitle) '</h1>']; idx = idx + 1;

figureCounter = 1;
if ~isempty(sections)
    for s = 1:numel(sections)
        section = sections(s);
        if isfield(section, 'Title') && ~isempty(section.Title)
            htmlParts{idx} = ['<h2 style="margin: 18px 0 10px 0; ' headingStyle ...
                '">' escapeHtml(section.Title) '</h2>']; idx = idx + 1;
        end

        if isfield(section, 'Paragraphs') && ~isempty(section.Paragraphs)
            for p = 1:numel(section.Paragraphs)
                htmlParts{idx} = ['<p style="' paragraphStyle '">' ...
                    escapeHtml(section.Paragraphs{p}) '</p>']; idx = idx + 1;
            end
        end

        if isfield(section, 'Bullets') && ~isempty(section.Bullets)
            htmlParts{idx} = ['<ul style="padding-left: 20px; ' listStyle '">']; idx = idx + 1;
            for b = 1:numel(section.Bullets)
                htmlParts{idx} = ['<li style="margin: 0 0 6px 0;">' ...
                    escapeHtml(section.Bullets{b}) '</li>']; idx = idx + 1;
            end
            htmlParts{idx} = '</ul>'; idx = idx + 1;
        end

        if isfield(section, 'Tables') && ~isempty(section.Tables)
            for t = 1:numel(section.Tables)
                tableHtml = buildTableHtml(section.Tables(t), tableStyle, tdStyle);
                htmlParts{idx} = tableHtml; idx = idx + 1;
            end
        end

        if isfield(section, 'Figures') && ~isempty(section.Figures)
            [normalizedFigures, figureCounter] = normalizeFiguresForHtml(section.Figures, baseDir, figureCounter);
            figureHtml = buildFiguresHtml(normalizedFigures, figureStyle, imageStyle);
            if ~isempty(figureHtml)
                htmlParts{idx} = figureHtml; idx = idx + 1;
            end
        end
    end
end

htmlParts{idx} = '</body>'; idx = idx + 1;
htmlParts{idx} = '</html>';

html = '';
for i = 1:numel(htmlParts)
    html = [html htmlParts{i} sprintf('\n')];
end

html = replacePlaceholdersInHtml(html, options, tableStyle, tdStyle, figureStyle, imageStyle, baseDir, figureCounter);

fid = fopen(outputPath, 'w');
if fid == -1
    error('无法写入输出文件：%s', outputPath);
end
fwrite(fid, html, 'char');
fclose(fid);

end

%-------------------------------------------------------------------------%
function html = buildTableHtml(tbl, tableStyle, tdStyle)
html = '';
if ~isfield(tbl, 'Rows') || isempty(tbl.Rows)
    return;
end

rows = size(tbl.Rows, 1);
cols = size(tbl.Rows, 2);

htmlParts = {};
idx = 1;
htmlParts{idx} = ['<table style="' tableStyle '">']; idx = idx + 1;

if isfield(tbl, 'Header') && ~isempty(tbl.Header)
    htmlParts{idx} = '<tr>'; idx = idx + 1;
    for c = 1:cols
        htmlParts{idx} = ['<th style="' tdStyle ' background-color: #f5f5f5; font-weight: bold;">' ...
            escapeHtml(tbl.Header{c}) '</th>']; idx = idx + 1;
    end
    htmlParts{idx} = '</tr>'; idx = idx + 1;
end

for r = 1:size(tbl.Rows, 1)
    htmlParts{idx} = '<tr>'; idx = idx + 1;
    for c = 1:cols
        htmlParts{idx} = ['<td style="' tdStyle '">' escapeHtml(tbl.Rows{r, c}) '</td>']; idx = idx + 1;
    end
    htmlParts{idx} = '</tr>'; idx = idx + 1;
end

htmlParts{idx} = '</table>';

html = '';
for i = 1:numel(htmlParts)
    html = [html htmlParts{i} sprintf('\n')];
end
end

%-------------------------------------------------------------------------%
function html = buildFiguresHtml(figures, figureStyle, imageStyle)
html = '';
if isempty(figures)
    return;
end

htmlParts = {};
idx = 1;

rowIndices = ones(1, numel(figures));
for k = 1:numel(figures)
    if isfield(figures(k), 'RowIndex') && ~isempty(figures(k).RowIndex)
        rowIndices(k) = figures(k).RowIndex;
    end
end
uniqueRows = unique(rowIndices);

for r = 1:numel(uniqueRows)
    currentFigures = figures(rowIndices == uniqueRows(r));
    htmlParts{idx} = ['<div style="' figureStyle '">']; idx = idx + 1;
    for c = 1:numel(currentFigures)
        captionText = '';
        if isfield(currentFigures(c), 'Caption') && ~isempty(currentFigures(c).Caption)
            captionText = currentFigures(c).Caption;
        end
        htmlParts{idx} = ['<div style="display: inline-block; margin-right: 12px;">' ...
            '<img src="' escapeHtml(currentFigures(c).Path) '" style="' imageStyle '" alt="figure" />'];
        idx = idx + 1;
        if ~isempty(captionText)
            htmlParts{idx} = ['<div style="margin-top: 4px; word-break: break-all;">' ...
                escapeHtml(captionText) '</div>']; idx = idx + 1;
        end
        htmlParts{idx} = '</div>'; idx = idx + 1;
    end
    htmlParts{idx} = '</div>'; idx = idx + 1;
end

html = '';
for i = 1:numel(htmlParts)
    html = [html htmlParts{i} sprintf('\n')];
end
end

%-------------------------------------------------------------------------%
function html = replacePlaceholdersInHtml(html, options, tableStyle, tdStyle, figureStyle, imageStyle, baseDir, figureCounter)
if ~isstruct(options) || ~isfield(options, 'Placeholders') || isempty(options.Placeholders)
    return;
end

placeholderNames = fieldnames(options.Placeholders);

for idx = 1:numel(placeholderNames)
    name = placeholderNames{idx};
    token = ['{{' name '}}'];
    payload = options.Placeholders.(name);

    replacement = '';
    if ischar(payload)
        replacement = escapeHtml(payload);
    elseif isstruct(payload)
        if isfield(payload, 'Rows')
            replacement = buildTableHtml(payload, tableStyle, tdStyle);
        elseif isfield(payload, 'Path') || isfield(payload, 'FigureHandle') || numel(payload) > 1
            [normalizedFigures, figureCounter] = normalizeFiguresForHtml(payload, baseDir, figureCounter);
            replacement = buildFiguresHtml(normalizedFigures, figureStyle, imageStyle);
        end
    end

    html = strrep(html, token, replacement);
end
end

%-------------------------------------------------------------------------%
function [figures, figureCounter] = normalizeFiguresForHtml(figures, baseDir, figureCounter)
if isempty(figures)
    return;
end

for idx = 1:numel(figures)
    [figures(idx), figureCounter] = ensureFigurePathForHtml(figures(idx), baseDir, figureCounter);
end
end

%-------------------------------------------------------------------------%
function [figStruct, figureCounter] = ensureFigurePathForHtml(figStruct, baseDir, figureCounter)
if ~isfield(figStruct, 'Path')
    figStruct.Path = '';
end

candidateHandle = [];
if ~isempty(figStruct.Path) && ~ischar(figStruct.Path)
    candidateHandle = figStruct.Path;
elseif isfield(figStruct, 'FigureHandle') && ~isempty(figStruct.FigureHandle)
    candidateHandle = figStruct.FigureHandle;
end

if ~isempty(candidateHandle) && isgraphics(candidateHandle)
    figHandle = candidateHandle;
    if ~strcmpi(get(figHandle, 'Type'), 'figure')
        figHandle = ancestor(figHandle, 'figure');
    end

    if ~isempty(figHandle) && isgraphics(figHandle)
        fileName = ['figure_' num2str(figureCounter) '.png'];
        figureCounter = figureCounter + 1;
        imgPath = fullfile(baseDir, fileName);
        print(figHandle, '-dpng', imgPath);
        figStruct.Path = imgPath;
    end
end
end

%-------------------------------------------------------------------------%
function text = escapeHtml(text)
if isempty(text)
    text = '';
    return;
end

text = strrep(text, '&', '&amp;');
text = strrep(text, '<', '&lt;');
text = strrep(text, '>', '&gt;');
text = strrep(text, '"', '&quot;');
text = strrep(text, '''', '&#39;');
end
