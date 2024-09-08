1. 安装 PyInstaller
首先，确保你已经安装了 PyInstaller。如果没有安装，可以使用以下命令进行安装：
pip install pyinstaller
2. 编写代码
确保你的代码已经保存为一个 .py 文件，例如 excel_processor.py。
3. 使用 PyInstaller 打包
在命令行中，导航到包含 excel_processor.py 文件的目录，然后运行以下命令：
pyinstaller --onefile --windowed excel_processor.py
参数说明：
--onefile：将所有文件打包成一个单独的 .exe 文件。
--windowed：不显示命令行窗口（适用于 GUI 应用程序）。
4. 查找生成的 .exe 文件
运行上述命令后，PyInstaller 会在当前目录下生成一个 dist 文件夹，其中包含生成的 .exe 文件。
5. 运行 .exe 文件
进入 dist 文件夹，找到生成的 .exe 文件，双击运行即可。
完整示例
假设你的代码文件名为 excel_processor.py，以下是完整的步骤：
安装 PyInstaller：
pip install pyinstaller
打包代码：
pyinstaller --onefile --windowed excel_processor.py
查找生成的 .exe 文件：
进入生成的 dist 文件夹，找到 excel_processor.exe 文件。
运行 .exe 文件：
双击 excel_processor.exe 文件即可运行你的应用程序。
注意事项
如果你的代码依赖于外部库或资源文件，确保这些文件在打包时也被正确包含。
如果打包后的 .exe 文件运行时出现错误，可以尝试使用 --debug 选项进行调试：
pyinstaller --onefile --windowed --debug excel_processor.py
通过这些步骤，你可以将 Python 脚本封装成一个独立的可执行文件，方便在没有 Python 环境的机器上运行。
