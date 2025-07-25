# PDF转Word工具使用说明

## 工具介绍

这是一个简单易用的PDF转Word工具，可以将PDF文件转换为Word文档（.docx格式），并保留原PDF中的文本、图片、表格和格式。

## 功能特点

- 简洁的图形用户界面
- 支持选择单个PDF文件进行转换
- 可自定义输出目录
- 实时显示转换进度
- 提供详细的转换结果报告
- 转换完成后可直接打开输出目录

## 安装说明

### 系统要求

- Windows 11 操作系统
- Python 3.6 或更高版本

### 安装步骤

1. 确保您的电脑已安装Python（如您所述，您已安装Python）
2. 下载并解压本工具的压缩包
3. 打开命令提示符（CMD）或PowerShell，进入解压后的目录
4. 运行以下命令安装所需依赖：

```
pip install -r requirements.txt
```

## 使用方法

1. 在命令提示符（CMD）或PowerShell中，进入工具目录并运行：

```
python pdf_to_word_gui.py
```

2. 在打开的图形界面中：
   - 点击"浏览"按钮选择要转换的PDF文件
   - 确认或修改输出目录
   - 点击"开始转换"按钮
   - 等待转换完成，查看结果报告
   - 转换完成后，可选择是否打开输出目录

## 注意事项

- 转换大型PDF文件可能需要较长时间，请耐心等待
- 转换过程中请勿关闭程序
- 如遇到复杂格式的PDF，转换后的Word文档可能需要手动调整部分格式

## 常见问题

1. **问题**：安装依赖时出现错误
   **解决方案**：确保您使用的是最新版本的pip，可以通过运行 `pip install --upgrade pip` 来更新

2. **问题**：程序无法启动
   **解决方案**：确保已正确安装所有依赖，并使用正确的Python版本

3. **问题**：转换后的文档格式有偏差
   **解决方案**：pdf2docx库尽力保留原格式，但复杂排版可能需要手动调整

## 技术支持

如有任何问题或建议，请联系开发者获取支持。
