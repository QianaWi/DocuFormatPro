📝 自动化 Word 排版重塑工具 - 完整产品需求文档 (PRD)
【项目背景与核心目标】
请为我开发一个独立的 Windows 桌面端应用程序。

核心功能： 用户批量导入 Word 文档（.doc / .docx），程序在后台静默调用本机 Word 引擎，对文档进行深度的格式化清洗、样式重置和自动化排版。

技术栈要求（极度重要）： 1. 前端 UI：使用 C# WPF (XAML) 构建，要求采用现代化的极简风格（Modern UI / Material Design），拒绝传统灰底控件。
2. 后端逻辑：使用 C# 结合 Microsoft.Office.Interop.Word 库进行底层 COM 操作。
3. 异步处理：必须使用 async/await 或 Task 等多线程技术处理 Word 文件的排版，确保主 UI 线程绝对流畅，绝不允许出现界面假死（未响应）状态。

【前端 UI 全局页面布局】
应用采用单窗口布局，分为以下三个核心区域：

1. 顶部状态与规则区 (Header & Rules Overview)

Logo & Title: 左上角显示应用名称 "DocuFormat Pro"。

规则展示卡片: 在右上角用微型标签（Badges）展示启用的排版规则，例如："页边距标准化"、"中英文字体分离"、"表格重塑"。

2. 核心交互区 (Main Content Area)

拖拽上传区 (Dropzone):

占据界面的主要空间，带有虚线边框和明显的上传图标。

文案："拖拽 Word 文档到此处，或点击选择文件"。

需支持文件拖拽事件拦截，仅允许拖入 Word 格式文件。

待处理文件列表 (File Queue List):

用户导入文件后，在拖拽区下方或切换视图显示文件列表（DataGrid 或 ListView）。

显示信息：文件名、文件大小、处理状态（排队中 / 处理中 / 完成 / 失败）。

操作列：每个文件右侧提供“移除”按钮。

3. 底部控制与进度区 (Bottom Control Panel)

全局进度条 (Global Progress Bar): 贯穿底部上方，平滑显示总体进度。

状态控制台 (Status Label): 动态显示当前正在执行的底层操作（例如："正在处理 第 2 个文件：项目汇报.docx..."）。

操作按钮:

主要按钮 (Primary Button)："开始排版"。大尺寸醒目按钮，开始后变为禁用状态并显示 Loading。

次要按钮 (Secondary Button)："清空列表" 或 "取消任务"。

【后台排版核心逻辑框架（需预留接口）】
请先在后台代码中搭建好 Word COM 调用的基础服务类（WordProcessingService.cs）。
要求：

控制 Word 在后台静默运行 (Visible = false)。

关闭屏幕刷新和警告弹窗以提升性能 (ScreenUpdating = false, DisplayAlerts = false)。

预留出排版处理的方法桩（Methods Stub），例如 SetPageMargins(), ResetStyles(), FormatAllTables()，我会在后续将具体的排版算法填入其中。

务必处理好 COM 对象的内存释放（Marshal.ReleaseComObject），防止后台残留 Word 进程。

【行动指令】
请根据以上需求，为我初始化整个 C# WPF 解决方案（Solution），包括完整的项目结构、XAML 界面设计代码，以及带有异步任务框架和 Word COM 调用的后台 C# 代码骨架。请确保代码可直接编译运行。