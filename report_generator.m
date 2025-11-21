function generateWordReport(outputPath, reportTitle, sections, options)
%GENERATEWORDREPORT Create a professional Word report via COM automation.
%   generateWordReport(outputPath, reportTitle, sections) opens a hidden
%   Microsoft Word session using ActiveX and builds a report from the
%   supplied content. The function is written to be compatible with MATLAB 7
%   and later (uses only actxserver and basic structs/cell arrays).
%
%   Parameters
%   ----------
%   outputPath : char
%       Full path (including .doc or .docx) where the report is saved.
%   reportTitle : char
%       Main title of the report displayed on the cover page.
%   sections : struct array
%       Each element may contain the fields:
%         - Title (char)       : section heading text.
%         - Paragraphs (cell)  : cell array of paragraph strings.
%         - Bullets (cell)     : bullet list items (optional).
%         - Tables (struct)    : optional struct array with fields
%                                Header (cell) and Rows (cell matrix).
%   options : struct (optional)
%       Optional settings to refine the document appearance:
%         - Template    : path to a .dot/.dotx template file.
%         - Author      : author name for document properties.
%         - Company     : company name for document properties.
%         - FooterText  : footer string printed on each page.
%         - AddPageNums : logical flag to insert page numbers (default true).
%         - HeadingFont : struct with Name/Size for headings (defaults Arial/16).
%         - BodyFont    : struct with Name/Size for body text (Arial/11).
%         - LineSpacing : multiple line spacing value (e.g., 1.15).
%         - SpaceBefore : paragraph space before in points (default 6).
%         - SpaceAfter  : paragraph space after in points (default 6).
%         - Margins     : struct with Top/Bottom/Left/Right in points (72 = 1");
%         - TableStyle  : built-in Word table style name (default 'Table Grid').
%
%   Example
%   -------
%   sections(1).Title = 'Overview';
%   sections(1).Paragraphs = {'Project background', 'Goals'};
%   sections(1).Bullets = {'Key milestone A', 'Key milestone B'};
%   sections(2).Title = 'Results';
%   sections(2).Tables(1).Header = {'Metric', 'Value'};
%   sections(2).Tables(1).Rows = {'Throughput', '24 req/s'; 'Latency', '120 ms'};
%   generateWordReport('C:\\temp\\demo.doc', 'Demo Report', sections);
%
%   Notes
%   -----
%   * This function requires Microsoft Word to be installed on Windows.
%   * To keep MATLAB 7 compatibility, the code uses invoke/get/set instead
%     of newer property accessor syntaxes.

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
    options.LineSpacing = 1.15; % multiple spacing (1 = single)
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
    options.Margins.Top = 72; % 1 inch
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
set(word, 'Visible', 0);  % keep hidden for automation

try
    doc = createDocument(word, options);
    setDocumentProperties(doc, options);
    selection = get(word, 'Selection');
    configurePageSetup(word, options);

    addCoverPage(selection, reportTitle, options);
    addBodyContent(selection, sections, options);
    addFooterAndNumbers(doc, options);

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
%CREATE DOCUMENT Either start from template or blank document.
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
%SETDOCUMENTPROPERTIES Configure author/company metadata when supplied.
if isfield(options, 'Author') && ~isempty(options.Author)
    invoke(doc, 'SetProperty', 'Author', options.Author);
end
if isfield(options, 'Company') && ~isempty(options.Company)
    invoke(doc, 'SetProperty', 'Company', options.Company);
end
end

%-------------------------------------------------------------------------%
function configurePageSetup(word, options)
%CONFIGUREPAGESETUP Apply margin settings to the active document.
doc = get(word, 'ActiveDocument');
pageSetup = get(doc, 'PageSetup');
set(pageSetup, 'TopMargin', options.Margins.Top);
set(pageSetup, 'BottomMargin', options.Margins.Bottom);
set(pageSetup, 'LeftMargin', options.Margins.Left);
set(pageSetup, 'RightMargin', options.Margins.Right);
end

%-------------------------------------------------------------------------%
function applyParagraphFormatting(selection, options, alignment)
%APPLYPARAGRAPHFORMATTING Set alignment, spacing, and line spacing on the selection.
paraFormat = get(selection, 'ParagraphFormat');
if nargin >= 3
    set(paraFormat, 'Alignment', alignment);
end
set(paraFormat, 'LineSpacingRule', 5); % wdLineSpaceMultiple
set(paraFormat, 'LineSpacing', options.LineSpacing * 12);
set(paraFormat, 'SpaceBefore', options.SpaceBefore);
set(paraFormat, 'SpaceAfter', options.SpaceAfter);
end

%-------------------------------------------------------------------------%
function addCoverPage(selection, reportTitle, options)
%ADDCOVERPAGE Simple centered cover page with title and optional footer.
    invoke(selection, 'WholeStory');
    invoke(selection, 'Delete');

    applyParagraphFormatting(selection, options, 1); % wdAlignParagraphCenter
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
    invoke(selection, 'InsertBreak', 3); % wdPageBreak
end

%-------------------------------------------------------------------------%
function addBodyContent(selection, sections, options)
%ADDBODYCONTENT Iterate over the section structs and emit content.
    if nargin < 2 || isempty(sections)
        return;
    end

    for k = 1:numel(sections)
        section = sections(k);

        if isfield(section, 'Title') && ~isempty(section.Title)
            applyParagraphFormatting(selection, options, 0); % wdAlignParagraphLeft
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

    invoke(selection, 'InsertBreak', 7); % wdSectionBreakContinuous
end
end

%-------------------------------------------------------------------------%
function addTable(selection, tbl, options)
%ADDTABLE Insert a simple table with optional header row.
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

invoke(wordTable, 'AutoFitBehavior', 2); % wdAutoFitContent

rangeAfter = wordTable.Range;
set(rangeAfter, 'Collapse', 0); % wdCollapseEnd
invoke(rangeAfter, 'Select');
end

%-------------------------------------------------------------------------%
function addFooterAndNumbers(doc, options)
%ADDFOOTERANDNUMBERS Footer text and page numbers for each section.
sections = get(doc, 'Sections');
count = get(sections, 'Count');

for s = 1:count
    section = invoke(sections, 'Item', s);
    footers = get(section, 'Footers');
    primaryFooter = invoke(footers, 'Item', 1); % wdHeaderFooterPrimary
    range = get(primaryFooter, 'Range');

    if isfield(options, 'FooterText') && ~isempty(options.FooterText)
        invoke(range, 'Text', options.FooterText);
    end

    pageNumbers = get(primaryFooter, 'PageNumbers');
    if options.AddPageNums
        invoke(pageNumbers, 'Add', 1); % wdAlignPageNumberCenter
    end
end
end
