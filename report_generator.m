function generateWordReport(outputPath, reportTitle, sections, options)
%GENERATEWORDREPORT 使用 COM 自动化生成专业的 Word 报告。
%   generateWordReport(outputPath, reportTitle, sections) 会通过 ActiveX 打开
%   一个隐藏的 Microsoft Word 会话，并按给定内容构建报告。为兼容 MATLAB 7
%   及更高版本，代码仅使用 actxserver 和基础的 struct/cell 数组。
%
%   参数
%   ----
%   outputPath : char
%       报告保存的完整路径（包含 .doc 或 .docx）。
%   reportTitle : char
%       封面页展示的主标题。
%   sections : struct array
%       每个元素可包含以下字段：
%         - Title (char)       : 章节标题文本。
%         - Paragraphs (cell)  : 段落字符串的单元格数组。
%         - Bullets (cell)     : 项目符号列表（可选）。
%         - Tables (struct)    : 可选的表格结构数组，包含 Header (cell)、Rows (cell 矩阵)。
%         - Figures (struct)   : 可选的图片结构数组，包含 Path (char)、Caption (char)、RowIndex (int)。
%                               Path 可为文件路径，也可直接传入 figure/axes/graphics handle。
%                               也可使用 FigureHandle 字段显式传入句柄，函数会临时导出图片。
%   options : struct (optional)
%       可选设置，用于细化文档外观：
%         - Template    : .dot/.dotx 模板文件路径。
%         - Author      : 文档属性中的作者名。
%         - Company     : 文档属性中的公司名。
%         - FooterText  : 每页底部的页脚文本。
%         - AddPageNums : 是否插入页码的逻辑开关（默认 true）。
%         - HeadingFont : 标题字体的 Name/Size 结构（默认 Arial/16）。
%         - BodyFont    : 正文字体的 Name/Size 结构（默认 Arial/11）。
%         - LineSpacing : 多倍行距数值（如 1.15）。
%         - SpaceBefore : 段前间距，单位为磅（默认 6）。
%         - SpaceAfter  : 段后间距，单位为磅（默认 6）。
%         - Margins     : 页边距 Top/Bottom/Left/Right（72 = 1 英寸）。
%         - TableStyle  : Word 内置表格样式名（默认 'Table Grid'）。
%         - Placeholders: 占位符名称到文本/结构内容的映射结构。
%
%   示例
%   ----
%   sections(1).Title = 'Overview';
%   sections(1).Paragraphs = {'Project background', 'Goals'};
%   sections(1).Bullets = {'Key milestone A', 'Key milestone B'};
%   sections(2).Title = 'Results';
%   sections(2).Tables(1).Header = {'Metric', 'Value'};
%   sections(2).Tables(1).Rows = {'Throughput', '24 req/s'; 'Latency', '120 ms'};
%   generateWordReport('C:\\temp\\demo.doc', 'Demo Report', sections);
%
%   注意事项
%   -------
%   * 需要在 Windows 上安装 Microsoft Word。
%   * 为保持 MATLAB 7 兼容性，使用 invoke/get/set 调用，而非新语法的属性访问。

if nargin < 4 || isempty(options)
    options = struct;
end

if ~isfield(options, 'AddPageNums')
    options.AddPageNums = true;
end

if ~isfield(options, 'HeadingFont') || ~isfield(options.HeadingFont, 'Name')
    options.HeadingFont.Name = 'Arial';
end
if ~isfield(options.HeadingFont, 'Size')
    options.HeadingFont.Size = 16;
end
if ~isfield(options, 'BodyFont') || ~isfield(options.BodyFont, 'Name')
    options.BodyFont.Name = 'Arial';
end
if ~isfield(options.BodyFont, 'Size')
    options.BodyFont.Size = 11;
end
if ~isfield(options, 'LineSpacing')
    options.LineSpacing = 1.15; % 多倍行距（1 = 单倍）
end
if ~isfield(options, 'SpaceBefore')
    options.SpaceBefore = 6;
end
if ~isfield(options, 'SpaceAfter')
    options.SpaceAfter = 6;
end
if ~isfield(options, 'Margins')
    options.Margins = struct;
end
if ~isfield(options.Margins, 'Top')
    options.Margins.Top = 72; % 1 英寸
end
if ~isfield(options.Margins, 'Bottom')
    options.Margins.Bottom = 72;
end
if ~isfield(options.Margins, 'Left')
    options.Margins.Left = 72;
end
if ~isfield(options.Margins, 'Right')
    options.Margins.Right = 72;
end
if ~isfield(options, 'TableStyle')
    options.TableStyle = 'Table Grid';
end

word = actxserver('Word.Application');
set(word, 'Visible', 0);  % 自动化过程中保持隐藏

try
    doc = createDocument(word, options);
    setDocumentProperties(doc, options);
    selection = get(word, 'Selection');
    configurePageSetup(word, options);

    addCoverPage(selection, reportTitle, options);
    addBodyContent(selection, sections, options);
    addFooterAndNumbers(doc, options);
    replacePlaceholders(doc, options);

    invoke(doc, 'SaveAs', outputPath);
catch err
    if exist('word', 'var') == 1
        invoke(word, 'Quit');
        delete(word);
    end
    rethrow(err);
end

invoke(doc, 'Close');
invoke(word, 'Quit');
delete(word);
end

%-------------------------------------------------------------------------%
function doc = createDocument(word, options)
%CREATEDOCUMENT 从模板或空白文档创建文档。
if isfield(options, 'Template') && ~isempty(options.Template) && exist(options.Template, 'file')
    docs = get(word, 'Documents');
    doc = invoke(docs, 'Open', options.Template);
else
    docs = get(word, 'Documents');
    doc = invoke(docs, 'Add');
end
end

%-------------------------------------------------------------------------%
function setDocumentProperties(doc, options)
%SETDOCUMENTPROPERTIES 根据传入内容配置作者/公司属性。
if isfield(options, 'Author') && ~isempty(options.Author)
    invoke(doc, 'SetProperty', 'Author', options.Author);
end
if isfield(options, 'Company') && ~isempty(options.Company)
    invoke(doc, 'SetProperty', 'Company', options.Company);
end
end

%-------------------------------------------------------------------------%
function configurePageSetup(word, options)
%CONFIGUREPAGESETUP 为当前文档应用页边距设置。
doc = get(word, 'ActiveDocument');
pageSetup = get(doc, 'PageSetup');
set(pageSetup, 'TopMargin', options.Margins.Top);
set(pageSetup, 'BottomMargin', options.Margins.Bottom);
set(pageSetup, 'LeftMargin', options.Margins.Left);
set(pageSetup, 'RightMargin', options.Margins.Right);
end

%-------------------------------------------------------------------------%
function applyParagraphFormatting(selection, options, alignment)
%APPLYPARAGRAPHFORMATTING 设置选区的对齐、段间距和行距。
paraFormat = get(selection, 'ParagraphFormat');
if nargin >= 3
    set(paraFormat, 'Alignment', alignment);
end
    set(paraFormat, 'LineSpacingRule', 5); % wdLineSpaceMultiple（多倍行距）
set(paraFormat, 'LineSpacing', options.LineSpacing * 12);
set(paraFormat, 'SpaceBefore', options.SpaceBefore);
set(paraFormat, 'SpaceAfter', options.SpaceAfter);
end

%-------------------------------------------------------------------------%
function addCoverPage(selection, reportTitle, options)
%ADDCOVERPAGE 创建居中的封面页，包含标题与可选页脚。
    invoke(selection, 'WholeStory');
    invoke(selection, 'Delete');

    applyParagraphFormatting(selection, options, 1); % wdAlignParagraphCenter（居中）
    set(selection.Font, 'Name', options.HeadingFont.Name);
    set(selection.Font, 'Size', options.HeadingFont.Size);
    set(selection.Font, 'Bold', 1);
    invoke(selection, 'TypeText', reportTitle);
    invoke(selection, 'TypeParagraph');

    set(selection.Font, 'Bold', 0);
    set(selection.Font, 'Name', options.BodyFont.Name);
    set(selection.Font, 'Size', options.BodyFont.Size);
    if isfield(options, 'Author') && ~isempty(options.Author)
        invoke(selection, 'TypeText', ['Author: ' options.Author]);
        invoke(selection, 'TypeParagraph');
    end
    if isfield(options, 'Company') && ~isempty(options.Company)
        invoke(selection, 'TypeText', ['Company: ' options.Company]);
        invoke(selection, 'TypeParagraph');
    end
    invoke(selection, 'InsertBreak', 3); % wdPageBreak（分页符）
end

%-------------------------------------------------------------------------%
function addBodyContent(selection, sections, options)
%ADDBODYCONTENT 迭代章节结构并生成正文内容。
    if nargin < 2 || isempty(sections)
        return;
    end

    for k = 1:numel(sections)
        section = sections(k);

        if isfield(section, 'Title') && ~isempty(section.Title)
            applyParagraphFormatting(selection, options, 0); % wdAlignParagraphLeft（左对齐）
            set(selection.Font, 'Name', options.HeadingFont.Name);
            set(selection.Font, 'Size', options.HeadingFont.Size);
            set(selection.Font, 'Bold', 1);
            invoke(selection, 'TypeText', section.Title);
            invoke(selection, 'TypeParagraph');
        end

        set(selection.Font, 'Bold', 0);
        set(selection.Font, 'Name', options.BodyFont.Name);
        set(selection.Font, 'Size', options.BodyFont.Size);
        applyParagraphFormatting(selection, options, 0);

        if isfield(section, 'Paragraphs') && ~isempty(section.Paragraphs)
            for p = 1:numel(section.Paragraphs)
                invoke(selection, 'TypeText', section.Paragraphs{p});
                invoke(selection, 'TypeParagraph');
            end
        end

        if isfield(section, 'Bullets') && ~isempty(section.Bullets)
            for b = 1:numel(section.Bullets)
            invoke(selection, 'TypeText', ['• ' section.Bullets{b}]);
            invoke(selection, 'TypeParagraph');
        end
        invoke(selection, 'TypeParagraph');
        end

        if isfield(section, 'Tables') && ~isempty(section.Tables)
            for t = 1:numel(section.Tables)
                addTable(selection, section.Tables(t), options);
                invoke(selection, 'TypeParagraph');
            end
        end

        if isfield(section, 'Figures') && ~isempty(section.Figures)
            addFigures(selection, section.Figures);
        end

    invoke(selection, 'InsertBreak', 7); % wdSectionBreakContinuous（连续分节符）
end
end

%-------------------------------------------------------------------------%
function addTable(selection, tbl, options)
%ADDTABLE 插入带可选表头的简单表格。
if ~isfield(tbl, 'Rows') || isempty(tbl.Rows)
    return;
end

rows = size(tbl.Rows, 1);
cols = size(tbl.Rows, 2);
if isfield(tbl, 'Header') && ~isempty(tbl.Header)
    rows = rows + 1;
end

range = get(selection, 'Range');
tables = get(selection, 'Tables');
wordTable = invoke(tables, 'Add', range, rows, cols);
set(wordTable, 'Style', options.TableStyle);
set(wordTable.Range.Font, 'Name', options.BodyFont.Name);
set(wordTable.Range.Font, 'Size', options.BodyFont.Size);

rowIndex = 1;
if isfield(tbl, 'Header') && ~isempty(tbl.Header)
    for c = 1:cols
        cellObj = invoke(wordTable, 'Cell', rowIndex, c);
        invoke(cellObj, 'Range', 'Text', tbl.Header{c});
        set(cellObj.Range.Font, 'Bold', 1);
    end
    rowIndex = rowIndex + 1;
end

for r = 1:size(tbl.Rows, 1)
    for c = 1:cols
        cellObj = invoke(wordTable, 'Cell', rowIndex, c);
        invoke(cellObj, 'Range', 'Text', tbl.Rows{r, c});
    end
    rowIndex = rowIndex + 1;
end

invoke(wordTable, 'AutoFitBehavior', 2); % wdAutoFitContent（按内容自适应）

rangeAfter = wordTable.Range;
set(rangeAfter, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）
invoke(rangeAfter, 'Select');
end

%-------------------------------------------------------------------------%
function addFigures(selection, figures)
%ADDFIGURES 按行分组插入图片并添加标题。
if nargin < 2 || isempty(figures)
    return;
end

    [figures, tempFiles] = normalizeFigures(figures);
    try
        rowIndices = ones(1, numel(figures));
        for idx = 1:numel(figures)
            if isfield(figures(idx), 'RowIndex') && ~isempty(figures(idx).RowIndex)
                rowIndices(idx) = figures(idx).RowIndex;
            end
        end

        uniqueRows = unique(rowIndices);
        for r = 1:numel(uniqueRows)
            currentFigures = figures(rowIndices == uniqueRows(r));
            addFigureRow(selection, currentFigures);
        end
    catch err
        deleteTempFiles(tempFiles);
        rethrow(err);
    end

    deleteTempFiles(tempFiles);
end

%-------------------------------------------------------------------------%
function addFigureRow(selection, figureRow)
%ADDFIGUREROW 将一组图片排布在单行表格中。
if isempty(figureRow)
    return;
end

range = get(selection, 'Range');
tables = get(selection, 'Tables');
wordTable = invoke(tables, 'Add', range, 1, numel(figureRow));
set(wordTable.Borders, 'Enable', 0);

for c = 1:numel(figureRow)
    cellObj = invoke(wordTable, 'Cell', 1, c);
    cellRange = get(cellObj, 'Range');
    inlineShapes = get(cellRange, 'InlineShapes');
    invoke(inlineShapes, 'AddPicture', figureRow(c).Path, 0, 1);

    set(cellRange, 'Collapse', 0); % wdCollapseEnd（折叠到末尾）
    invoke(cellRange, 'Select');

    captionText = '';
    if isfield(figureRow(c), 'Caption') && ~isempty(figureRow(c).Caption)
        captionText = [' ' figureRow(c).Caption];
    end
    invoke(selection, 'InsertCaption', 'Figure', captionText);
end

rangeAfter = wordTable.Range;
set(rangeAfter, 'Collapse', 0); % wdCollapseEnd（折叠到末尾）
invoke(rangeAfter, 'Select');
invoke(selection, 'TypeParagraph');
end

%-------------------------------------------------------------------------%
function addFooterAndNumbers(doc, options)
%ADDFOOTERANDNUMBERS 为各节添加页脚文本与页码。
    sections = get(doc, 'Sections');
    count = get(sections, 'Count');

for s = 1:count
    section = invoke(sections, 'Item', s);
    footers = get(section, 'Footers');
    primaryFooter = invoke(footers, 'Item', 1); % wdHeaderFooterPrimary（主页脚）
    range = get(primaryFooter, 'Range');

    if isfield(options, 'FooterText') && ~isempty(options.FooterText)
        invoke(range, 'Text', options.FooterText);
    end

    pageNumbers = get(primaryFooter, 'PageNumbers');
    if options.AddPageNums
        invoke(pageNumbers, 'Add', 1); % wdAlignPageNumberCenter（页码居中）
    end
end
end

%-------------------------------------------------------------------------%
function replacePlaceholders(doc, options)
%REPLACEPLACEHOLDERS 在生成文档后替换占位符文本或插入对象。
if ~isfield(options, 'Placeholders') || isempty(options.Placeholders)
    return;
end

placeholderNames = fieldnames(options.Placeholders);

for idx = 1:numel(placeholderNames)
    name = placeholderNames{idx};
    token = ['{{' name '}}'];
    payload = options.Placeholders.(name);

    searchRange = get(doc, 'Content');
    findObj = get(searchRange, 'Find');
    set(findObj, 'Forward', 1);
    set(findObj, 'Format', 0);

    while invoke(findObj, 'Execute', token, 0, 0, 0, 0, 0, 1, 1, 0, '', 0)
        if ischar(payload)
            set(searchRange, 'Text', payload);
        elseif isstruct(payload)
            if isfield(payload, 'Rows')
                addTableAtRange(searchRange, payload, options);
            elseif isfield(payload, 'Path') || isfield(payload, 'FigureHandle') || numel(payload) > 1
                addFiguresAtRange(searchRange, payload);
            else
                set(searchRange, 'Text', '');
            end
        else
            set(searchRange, 'Text', '');
        end

        startPos = get(searchRange, 'End');
        docContent = get(doc, 'Content');
        searchRange = invoke(doc, 'Range', startPos, get(docContent, 'End'));
        findObj = get(searchRange, 'Find');
        set(findObj, 'Forward', 1);
        set(findObj, 'Format', 0);
    end
end
end

%-------------------------------------------------------------------------%
function addTableAtRange(range, tbl, options)
%ADDTABLEATRANGE 在给定范围插入表格。
set(range, 'Text', '');
set(range, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）

if ~isfield(tbl, 'Rows') || isempty(tbl.Rows)
    return;
end

rows = size(tbl.Rows, 1);
cols = size(tbl.Rows, 2);
if isfield(tbl, 'Header') && ~isempty(tbl.Header)
    rows = rows + 1;
end

tables = get(range, 'Tables');
wordTable = invoke(tables, 'Add', range, rows, cols);
set(wordTable, 'Style', options.TableStyle);
set(wordTable.Range.Font, 'Name', options.BodyFont.Name);
set(wordTable.Range.Font, 'Size', options.BodyFont.Size);

rowIndex = 1;
if isfield(tbl, 'Header') && ~isempty(tbl.Header)
    for c = 1:cols
        cellObj = invoke(wordTable, 'Cell', rowIndex, c);
        invoke(cellObj, 'Range', 'Text', tbl.Header{c});
        set(cellObj.Range.Font, 'Bold', 1);
    end
    rowIndex = rowIndex + 1;
end

for r = 1:size(tbl.Rows, 1)
    for c = 1:cols
        cellObj = invoke(wordTable, 'Cell', rowIndex, c);
        invoke(cellObj, 'Range', 'Text', tbl.Rows{r, c});
    end
    rowIndex = rowIndex + 1;
end

invoke(wordTable, 'AutoFitBehavior', 2); % wdAutoFitContent（按内容自适应）

rangeAfter = wordTable.Range;
set(rangeAfter, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）
invoke(rangeAfter, 'Select');
end

%-------------------------------------------------------------------------%
function addFiguresAtRange(range, figures)
%ADDFIGURESATRANGE 在占位符范围插入一行或多行图片。
if nargin < 2 || isempty(figures)
    set(range, 'Text', '');
    return;
end

    [figures, tempFiles] = normalizeFigures(figures);
    try
        set(range, 'Text', '');
        set(range, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）

        rowIndices = ones(1, numel(figures));
        for idx = 1:numel(figures)
            if isfield(figures(idx), 'RowIndex') && ~isempty(figures(idx).RowIndex)
                rowIndices(idx) = figures(idx).RowIndex;
            end
        end

        uniqueRows = unique(rowIndices);
        for r = 1:numel(uniqueRows)
            currentFigures = figures(rowIndices == uniqueRows(r));

            tables = get(range, 'Tables');
            wordTable = invoke(tables, 'Add', range, 1, numel(currentFigures));
            set(wordTable.Borders, 'Enable', 0);

            for c = 1:numel(currentFigures)
                cellObj = invoke(wordTable, 'Cell', 1, c);
                cellRange = get(cellObj, 'Range');
                inlineShapes = get(cellRange, 'InlineShapes');
                invoke(inlineShapes, 'AddPicture', currentFigures(c).Path, 0, 1);

                set(cellRange, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）
                captionText = '';
                if isfield(currentFigures(c), 'Caption') && ~isempty(currentFigures(c).Caption)
                    captionText = [' ' currentFigures(c).Caption];
                end
                selection = get(get(range, 'Document').Application, 'Selection');
                invoke(cellRange, 'Select');
                invoke(selection, 'InsertCaption', 'Figure', captionText);
            end

            range = wordTable.Range;
            set(range, 'Collapse', 0); % wdCollapseEnd（折叠至末尾）
        end
    catch err
        deleteTempFiles(tempFiles);
        rethrow(err);
    end

    deleteTempFiles(tempFiles);
end

%-------------------------------------------------------------------------%
function [figures, tempFiles] = normalizeFigures(figures)
%NORMALIZEFIGURES 确保所有图片项均映射到可读的文件路径。
tempFiles = {};
if isempty(figures)
    return;
end

for idx = 1:numel(figures)
    [figures(idx), tempFiles] = ensureFigurePath(figures(idx), tempFiles);
end
end

%-------------------------------------------------------------------------%
function [figStruct, tempFiles] = ensureFigurePath(figStruct, tempFiles)
%ENSUREFIGUREPATH 将句柄导出为临时图片文件。
candidateHandle = [];

if isfield(figStruct, 'Path') && ~isempty(figStruct.Path) && ~ischar(figStruct.Path)
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
        tempPath = [tempname '.png'];
        print(figHandle, '-dpng', tempPath);
        figStruct.Path = tempPath;
        tempFiles{end + 1} = tempPath;
    end
end
end

%-------------------------------------------------------------------------%
function deleteTempFiles(tempFiles)
%DELETETEMPFILES 删除导出的临时图片。
for idx = 1:numel(tempFiles)
    if exist(tempFiles{idx}, 'file')
        delete(tempFiles{idx});
    end
end
end
