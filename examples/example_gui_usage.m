%EXAMPLE_GUI_USAGE 使用 launchReportGui 演示 HTML 与图像预览。
% 该脚本仅依赖 MATLAB 7 兼容语法，可直接运行生成示例 HTML 与图像
% 文件，然后调用 launchReportGui 进行浏览。

% 构造示例章节
sections = struct;
sections(1).Title = '示例章节';
sections(1).Paragraphs = {'此示例展示 generateHtmlReport 与 GUI 预览结合使用。', ...
    '左侧选择文件，右侧可切换 HTML 与图像预览。'};
sections(1).Bullets = {'兼容 MATLAB 7 的句柄图形', '支持 HTML 生成与回显'};

% HTML 生成配置
htmlOptions.BodyFontName = 'SimSun, Arial';
htmlOptions.BodyFontSize = 14;
htmlOptions.HeadingFontSize = 20;

htmlPath = fullfile(tempdir, 'gui_demo_report.html');
generateHtmlReport(htmlPath, 'GUI 示例报告', sections, htmlOptions);

% 构造一张示例图像
imgPath = fullfile(tempdir, 'gui_demo_image.png');
if exist('peaks', 'file') == 2
    imgData = peaks(256);
    fig = figure('Visible', 'off');
    imagesc(imgData); colormap(jet); axis off; axis image;
    saveas(fig, imgPath);
    close(fig);
else
    imgData = repmat(uint8(linspace(0, 255, 256)), 256, 1);
    imwrite(imgData, imgPath);
end

% GUI 选项与自定义绘图逻辑
options.InitialFiles = {htmlPath, imgPath};
options.Sections = sections;
options.ReportOptions = htmlOptions;
options.ReportTitle = 'GUI 示例报告';
options.HtmlOutputPath = htmlPath;
options.FigureLoader = @localFigureLoader;
options.StatusMessage = '预生成的 HTML 与图像已准备好。';

launchReportGui(options);

%--------------------------------------------------------------------------
function localFigureLoader(ax, filePath)
%LOCALFIGURELOADER 简单示例：优先尝试读取图片；若失败则尝试加载 MAT。 
    try
        img = imread(filePath);
        if exist('imshow', 'file') == 2
            imshow(img, 'Parent', ax);
        else
            image(img, 'Parent', ax);
            axis(ax, 'image');
        end
    catch
        % 如果不是图片，尝试加载矩阵
        data = load(filePath);
        fNames = fieldnames(data);
        if ~isempty(fNames)
            imagesc(ax, data.(fNames{1}));
            axis(ax, 'image');
        end
    end
    set(ax, 'Visible', 'on');
end
