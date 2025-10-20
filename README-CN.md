README.md：

## 项目概述

这个README文件是为一个将Markdown文件转换为PowerPoint演示文稿的Python应用程序编写的。该工具支持多种Markdown元素，包括标题、项目符号、表格和粗体文本格式。

## 主要功能

- 将Markdown文本或文件转换为PowerPoint（.pptx）格式
- 支持使用`---`分隔幻灯片
- 标题级别映射到幻灯片标题和内容标题
- 支持带缩进的项目符号
- 表格解析和渲染
- 粗体文本格式化（`**粗体文本**`）
- 基于Streamlit的Web用户界面
- 自动文件命名以避免覆盖

## 安装说明

1. 克隆代码仓库
2. 安装必要的依赖项：`python-pptx`和`streamlit`

## 使用方法

### 命令行使用
直接运行[mdtopptx.py](file://c:\workspace\pycodespace\abc\mdtopptx.py)脚本，默认会查找`.\input\sample.md`文件并将演示文稿输出到`./output`目录。

### Web界面使用
运行[webui.py](file://c:\workspace\pycodespace\abc\webui.py)启动Streamlit网页界面，提供两种转换方式：
1. **文本输入**：直接在文本区域粘贴Markdown内容
2. **文件上传**：上传.md或.markdown文件

## 支持的Markdown语法

README详细说明了支持的Markdown元素：
- 幻灯片分隔符（---）
- 不同级别的标题（#、##、###等）
- 项目符号列表和缩进
- 表格（GitHub风格的Markdown表格）
- 文本格式（粗体文本）
- 普通段落

## 项目结构

说明了项目的文件组织结构，包括核心文件和目录。

## 其他信息

还包括了贡献指南、许可证信息和联系方式。

这个README为用户提供了清晰的使用说明，帮助他们快速上手使用这个Markdown到PowerPoint的转换工具。