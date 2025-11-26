function handles = launchReportGui(options)
%LAUNCHREPORTGUI 启动报告浏览 GUI（兼容 MATLAB 7 语法）。
%   handles = launchReportGui(options) 创建一个包含左侧文件列表、模式
%   切换列表以及右侧选项卡的界面。界面可使用 generateHtmlReport 生成 HTML
%   预览，或在图像/axes 选项卡中展示图片与 figure 快照。所有句柄图形
%   操作均使用 set/get，避免 MATLAB 7 以后新增的类特性。
%
%   options 可包含：
%     - InitialFiles     : 预置文件列表（cell 数组）。
%     - Sections         : 章节结构数组，传递给 generateHtmlReport。
%     - ReportOptions    : HTML 报告的选项结构。
%     - ReportTitle      : 报告标题，默认 'GUI Report Preview'。
%     - HtmlOutputPath   : 生成的临时 HTML 路径，默认 tempdir 下文件。
%     - HtmlGenerator    : 生成 HTML 的函数句柄（签名同 generateHtmlReport）。
%     - FigureLoader     : 可选函数句柄，接受 (axesHandle, filePath) 自定义绘图。
%     - StatusMessage    : 初始化状态栏文本。
%
%   返回值 handles 包含 figure、listbox、tab 等句柄，便于脚本进一步控制。

if nargin < 1
    options = struct;
end

if ~isfield(options, 'InitialFiles')
    options.InitialFiles = {};
end
if ~isfield(options, 'Sections')
    options.Sections = [];
end
if ~isfield(options, 'ReportOptions')
    options.ReportOptions = struct;
end
if ~isfield(options, 'ReportTitle')
    options.ReportTitle = 'GUI Report Preview';
end
if ~isfield(options, 'HtmlOutputPath')
    options.HtmlOutputPath = fullfile(tempdir, 'report_preview.html');
end
if ~isfield(options, 'HtmlGenerator')
    options.HtmlGenerator = @generateHtmlReport;
end
if ~isfield(options, 'FigureLoader')
    options.FigureLoader = [];
end
if ~isfield(options, 'StatusMessage')
    options.StatusMessage = '请选择左侧文件或使用“浏览”按钮加载示例。';
end

handles = struct;
handles.Options = options;
handles.FileListPaths = options.InitialFiles;
handles.ModeListStrings = {'HTML 报告', '图像/figure 预览'};

handles.Figure = figure('Name', 'Report Launcher', 'NumberTitle', 'off', ...
    'MenuBar', 'none', 'Toolbar', 'none', 'Units', 'normalized', ...
    'Position', [0.1 0.1 0.8 0.8]);

% 左侧面板：文件列表与模式切换
handles.LeftPanel = uipanel('Parent', handles.Figure, 'Units', 'normalized', ...
    'Position', [0.02 0.05 0.28 0.9], 'Title', '来源与模式');

handles.BrowseButton = uicontrol('Parent', handles.LeftPanel, 'Style', 'pushbutton', ...
    'Units', 'normalized', 'Position', [0.05 0.92 0.4 0.06], ...
    'String', '浏览文件...', 'Callback', @onBrowseFile);

handles.RefreshButton = uicontrol('Parent', handles.LeftPanel, 'Style', 'pushbutton', ...
    'Units', 'normalized', 'Position', [0.55 0.92 0.4 0.06], ...
    'String', '刷新', 'Callback', @onRefreshList);

handles.FileList = uicontrol('Parent', handles.LeftPanel, 'Style', 'listbox', ...
    'Units', 'normalized', 'Position', [0.05 0.45 0.9 0.45], ...
    'String', buildDisplayNames(handles.FileListPaths), 'UserData', handles.FileListPaths, ...
    'Callback', @onFileSelected, 'BackgroundColor', [1 1 1]);

handles.ModeList = uicontrol('Parent', handles.LeftPanel, 'Style', 'listbox', ...
    'Units', 'normalized', 'Position', [0.05 0.1 0.9 0.3], ...
    'String', handles.ModeListStrings, 'Value', 1, 'Callback', @onModeChanged, ...
    'BackgroundColor', [1 1 1]);

% 右侧面板：选项卡
handles.TabGroup = uitabgroup('Parent', handles.Figure, 'Units', 'normalized', ...
    'Position', [0.32 0.05 0.66 0.9]);
handles.HtmlTab = uitab('Parent', handles.TabGroup, 'Title', 'HTML 预览');
handles.FigureTab = uitab('Parent', handles.TabGroup, 'Title', '图像/figure');

% HTML 视图（uihtml/java/edit 兼容策略）
handles.HtmlPanel = uipanel('Parent', handles.HtmlTab, 'Units', 'normalized', ...
    'Position', [0 0 1 1], 'BorderType', 'none');
handles.HtmlViewer = createHtmlViewer(handles.HtmlPanel);

% 图像/figure 轴
handles.FigurePanel = uipanel('Parent', handles.FigureTab, 'Units', 'normalized', ...
    'Position', [0 0 1 1], 'BorderType', 'none');
handles.Axes = axes('Parent', handles.FigurePanel, 'Units', 'normalized', ...
    'Position', [0.08 0.08 0.84 0.84]);
set(handles.Axes, 'Visible', 'off');

handles.StatusBar = uicontrol('Parent', handles.Figure, 'Style', 'text', ...
    'Units', 'normalized', 'Position', [0.02 0.01 0.96 0.03], ...
    'HorizontalAlignment', 'left', 'String', options.StatusMessage);

    function onBrowseFile(~, ~)
        [fileName, filePath] = uigetfile({'*.*', '所有文件'}, '选择文件');
        if isequal(fileName, 0)
            return;
        end
        fullPath = fullfile(filePath, fileName);
        addFileToList(fullPath);
        setStatus(['已添加: ' fullPath]);
    end

    function onRefreshList(~, ~)
        set(handles.FileList, 'String', buildDisplayNames(handles.FileListPaths));
        set(handles.FileList, 'UserData', handles.FileListPaths);
        setStatus('列表已刷新。');
    end

    function onFileSelected(hObj, ~)
        contents = get(hObj, 'UserData');
        idx = get(hObj, 'Value');
        if isempty(contents) || idx < 1 || idx > numel(contents)
            return;
        end
        selectedPath = contents{idx};
        dispatchPreview(selectedPath);
    end

    function onModeChanged(~, ~)
        % 切换选项卡并尝试重新加载当前文件
        modeIdx = get(handles.ModeList, 'Value');
        if modeIdx == 1
            set(handles.TabGroup, 'SelectedTab', handles.HtmlTab);
        else
            set(handles.TabGroup, 'SelectedTab', handles.FigureTab);
        end
        onFileSelected(handles.FileList, []);
    end

    function dispatchPreview(filePath)
        modeIdx = get(handles.ModeList, 'Value');
        if modeIdx == 1
            loadHtmlReport(filePath);
        else
            loadFigurePreview(filePath);
        end
    end

    function addFileToList(fullPath)
        if ~iscell(handles.FileListPaths)
            handles.FileListPaths = {};
        end
        handles.FileListPaths{end + 1} = fullPath;
        set(handles.FileList, 'String', buildDisplayNames(handles.FileListPaths));
        set(handles.FileList, 'UserData', handles.FileListPaths);
        set(handles.FileList, 'Value', numel(handles.FileListPaths));
    end

    function loadHtmlReport(filePath)
        if nargin < 1 || isempty(filePath)
            filePath = handles.Options.HtmlOutputPath;
        end

        if ~isempty(handles.Options.HtmlGenerator) && (isempty(filePath) || ~isHtmlFile(filePath))
            try
                handles.Options.HtmlGenerator(handles.Options.HtmlOutputPath, ...
                    handles.Options.ReportTitle, handles.Options.Sections, handles.Options.ReportOptions);
                filePath = handles.Options.HtmlOutputPath;
                setStatus(['已生成 HTML 报告: ' filePath]);
            catch err
                setStatus(['生成 HTML 报告失败: ' err.message]);
                return;
            end
        end

        if exist(filePath, 'file') ~= 2
            setStatus(['找不到文件: ' filePath]);
            return;
        end

        renderHtmlInViewer(handles.HtmlViewer, filePath);
        set(handles.TabGroup, 'SelectedTab', handles.HtmlTab);
        setStatus(['已加载 HTML: ' filePath]);
    end

    function loadFigurePreview(filePath)
        cla(handles.Axes, 'reset');
        set(handles.Axes, 'Visible', 'off');
        if nargin < 1 || isempty(filePath) || exist(filePath, 'file') ~= 2
            setStatus('未找到可预览的图像/figure 文件。');
            return;
        end

        if ~isempty(handles.Options.FigureLoader)
            try
                feval(handles.Options.FigureLoader, handles.Axes, filePath);
                set(handles.Axes, 'Visible', 'on');
                set(handles.TabGroup, 'SelectedTab', handles.FigureTab);
                setStatus(['已使用自定义函数加载: ' filePath]);
                return;
            catch err
                setStatus(['自定义加载失败: ' err.message]);
            end
        end

        [~, ~, ext] = fileparts(filePath);
        ext = lower(ext);
        if strcmp(ext, '.fig')
            try
                figHandle = openfig(filePath, 'invisible');
                figAxes = findall(figHandle, 'Type', 'axes');
                if ~isempty(figAxes)
                    copyobj(figAxes, handles.FigurePanel);
                    setStatus(['已加载 FIG 文件: ' filePath]);
                end
                close(figHandle);
                set(handles.TabGroup, 'SelectedTab', handles.FigureTab);
                return;
            catch err
                setStatus(['读取 FIG 失败: ' err.message]);
            end
        end

        try
            imgData = imread(filePath);
            if exist('imshow', 'file') == 2
                imshow(imgData, 'Parent', handles.Axes);
            else
                image(imgData, 'Parent', handles.Axes);
                axis(handles.Axes, 'image');
            end
            set(handles.Axes, 'Visible', 'on');
            set(handles.TabGroup, 'SelectedTab', handles.FigureTab);
            setStatus(['已展示图像: ' filePath]);
        catch err
            setStatus(['读取图像失败: ' err.message]);
        end
    end

    function setStatus(msg)
        set(handles.StatusBar, 'String', msg);
        drawnow; %#ok<DRAWNOW>
    end
end

%-------------------------------------------------------------------------%
function names = buildDisplayNames(paths)
% 内部工具：根据路径生成 listbox 显示名称。
if isempty(paths)
    names = {'(无文件，请点击“浏览文件...” )'};
    return;
end
names = cell(1, numel(paths));
for i = 1:numel(paths)
    [~, name, ext] = fileparts(paths{i});
    names{i} = [name ext];
end
end

%-------------------------------------------------------------------------%
function viewer = createHtmlViewer(parent)
% 兼容 uihtml / Java / edit 的多级 HTML 视图。
viewer = struct('Mode', 'edit', 'Control', [], 'Container', []);

if exist('uihtml', 'file') == 2
    try
        viewer.Control = uihtml('Parent', parent, 'HTMLSource', '<html><body>就绪</body></html>', ...
            'Units', 'normalized', 'Position', [0 0 1 1]);
        viewer.Mode = 'uihtml';
        return;
    catch
        % 忽略，尝试 Java 方案
    end
end

if usejava('swing')
    try
        htmlPane = javax.swing.JEditorPane();
        htmlPane.setContentType('text/html');
        scroll = javax.swing.JScrollPane(htmlPane);
        [container, hContainer] = javacomponent(scroll, [0 0 1 1], parent); %#ok<JAVFMAT>
        set(hContainer, 'Units', 'normalized', 'Position', [0 0 1 1]);
        viewer.Control = htmlPane;
        viewer.Container = hContainer;
        viewer.Mode = 'java';
        return;
    catch
        % 继续使用 edit 方案
    end
end

viewer.Control = uicontrol('Parent', parent, 'Style', 'edit', 'Max', 10, 'Min', 0, ...
    'Units', 'normalized', 'Position', [0 0 1 1], 'HorizontalAlignment', 'left', ...
    'BackgroundColor', [1 1 1], 'String', 'HTML 预览将显示在这里。');
viewer.Mode = 'edit';
end

%-------------------------------------------------------------------------%
function renderHtmlInViewer(viewer, filePath)
% 根据 viewer 模式渲染 HTML 文件。
try
    switch viewer.Mode
        case 'uihtml'
            set(viewer.Control, 'HTMLSource', filePath);
        case 'java'
            viewer.Control.setPage(java.io.File(filePath).toURI().toURL());
        otherwise
            htmlText = fileread(filePath);
            set(viewer.Control, 'String', htmlText);
    end
catch err
    if strcmp(viewer.Mode, 'java')
        set(viewer.Control, 'Text', ['加载 HTML 失败: ' err.message]);
    else
        set(viewer.Control, 'String', ['加载 HTML 失败: ' err.message]);
    end
end
end

%-------------------------------------------------------------------------%
function flag = isHtmlFile(filePath)
[~, ~, ext] = fileparts(filePath);
ext = lower(ext);
flag = strcmp(ext, '.html') || strcmp(ext, '.htm');
end
