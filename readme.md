### 说明
可以直接运行主目录下的MarkdownEditor.jar文件。也可以打开工程编译运行。源文件和依赖文件在src、lib和主目录下。注意jacob.1.18-x64.dll也是依赖文件之一，需放在工程主目录下和可运行jar文件所在目录下。

### 实现功能
- 分为3个panel，左栏是标题导引，中间编辑框显示markdown文件内容，右边的panel显示html的预览。
- 更新预览和标题导引内容需要点击preview按钮。不支持实时预览。
- 点击左栏的标题，中间的编辑框会跳转到相应的header。
- 可处理1mb的markdown文件。注意1mb文件预览时间可能需要2~3秒。examples/example_1mb.md是大小超过1mb的文件。
- 可识别html tags。examples/example_with_html.md包含了html标签。
- 可以导入css文件定制。examples/github.css提供了github风格的css。
- 可以导出为html。点击file-save as html。
- 可以导出为docx。点击file-save as docx。
- 支持md文件的打开、保存、关闭等功能。

### 运行截图：
![](running_screenshot.png)