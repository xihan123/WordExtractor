# WordExtractor

![版本](https://img.shields.io/github/v/release/xihan123/WordExtractor)
![语言](https://img.shields.io/badge/语言-Python-brightgreen.svg)
![许可证](https://img.shields.io/badge/许可证-MIT-orange.svg)
![下载量](https://img.shields.io/github/downloads/xihan123/WordExtractor/total)

WordExtractor 是一个用于从 Word 文档 (`.docx`) 中批量提取数据并将其导出到 Excel 文件的工具。它通过用户定义的提取规则来定位和提取所需信息。

## 目录

- [功能特性](#功能特性)
- [如何使用](#如何使用)
  - [运行程序](#运行程序)
  - [主界面概览](#主界面概览)
  - [管理提取规则](#管理提取规则)
  - [选择文件并开始提取](#选择文件并开始提取)
  - [查看结果](#查看结果)
- [编写提取规则](#编写提取规则)
  - [规则基本构成](#规则基本构成)
  - [提取模式 (ExtractionMode)](#提取模式-extractionmode)
    - [1. 正则表达式 (REGEX)](#1-正则表达式-regex)
    - [2. 位置 (POSITION)](#2-位置-position)
    - [3. 书签 (BOOKMARK)](#3-书签-bookmark)
    - [4. 表格单元格 (TABLE_CELL)](#4-表格单元格-table_cell)
    - [5. 表格列 (TABLE_COLUMN)](#5-表格列-table_column)
    - [6. 表格行 (TABLE_ROW)](#6-表格行-table_row)
    - [7. 完整表格 (TABLE_FULL)](#7-完整表格-table_full)
  - [规则属性](#规则属性)
  - [保存和加载规则](#保存和加载规则)
- [从源码构建 (可选)](#从源码构建-可选)

## 功能特性

- **批量处理**: 支持一次性处理多个 Word 文档。
- **灵活的提取规则**: 支持多种提取方式，包括正则表达式、位置、书签和多种表格提取。
- **规则管理**: 方便地创建、编辑、删除、排序、导入和导出提取规则。
- **数据导出**: 将提取的数据导出为 Excel (`.xlsx`) 文件。
- **用户友好界面**: 提供图形用户界面，易于操作。
- **规则测试**: 支持对单个规则进行测试，方便调试。

## 如何使用

### 运行程序

如果您有可执行文件 (`WordExtractor.exe`)，直接双击运行即可。

### 主界面概览

程序主界面通常包含以下几个主要区域：

1. **规则管理区**: 用于添加、编辑、删除和管理提取规则。
2. **文件列表区**: 用于添加待处理的 Word 文档。
3. **输出设置区**: 设置输出 Excel 文件的路径和选项。
4. **控制区**: 开始、停止提取任务。
5. **状态/日志区**: 显示任务进度和结果信息。

### 管理提取规则

1. **添加规则**:
    - 点击“添加规则”按钮。
    - 在弹出的对话框中，填写规则名称、输出表头名、选择提取模式并配置相应的参数。
    - 点击“确定”保存规则。
2. **编辑规则**:
    - 在规则列表中选中要编辑的规则。
    - 点击“编辑规则”按钮或双击规则。
    - 修改后点击“确定”保存。
3. **删除规则**:
    - 选中规则，点击“删除规则”按钮。
4. **调整规则顺序**:
    - 选中规则，使用“上移”或“下移”按钮调整其在列表中的顺序。规则按列表顺序应用于文档。
5. **启用/禁用规则**:
    - 可以通过勾选或特定按钮来启用或禁用某条规则，被禁用的规则在提取过程中会被跳过。
6. **导入/导出规则**:
    - 使用“导入规则”和“导出规则”功能可以加载或保存规则配置（通常为 `.json` 文件），方便复用。

### 选择文件并开始提取

1. **添加文件**: 点击“添加文件”或类似按钮，选择一个或多个 `.docx` 文件。也可以支持添加文件夹。
2. **设置输出**: 指定提取结果要保存到的 Excel 文件名和路径。可以选择是否追加到现有文件、是否包含文件信息等。
3. **开始提取**: 点击“开始处理”或类似按钮启动提取过程。
4. **监控进度**: 界面会显示当前处理进度、已完成任务数、失败任务数等。

### 查看结果

提取完成后，打开指定的 Excel 文件即可查看提取到的数据。每一行通常对应一个源 Word 文档，每一列对应一条提取规则定义的字段。

## 编写提取规则

提取规则是 WordExtractor 的核心，它告诉程序如何从文档中找到并提取您需要的信息。

### 规则基本构成

每条规则通常包含以下基本信息：

- **字段名 (Field Name)**: 规则的描述性名称，方便用户识别。
- **表头名 (Header Name)**: 提取出的数据在最终 Excel 文件中对应的列标题。
- **提取模式 (Extraction Mode)**: 定义了提取数据所使用的方法。
- **配置 (Config)**: 根据所选提取模式，提供具体参数。
- **是否启用 (Enabled)**: 标记此规则是否在提取过程中生效。

### 提取模式 (ExtractionMode)

以下是支持的提取模式及其配置：

#### 1. 正则表达式 (REGEX)

通过正则表达式匹配并提取文本。

- **配置参数**:
  - `pattern` (字符串): 用于匹配文本的正则表达式。
  - `group` (整数, 可选, 默认 0): 指定返回哪个捕获组。0 表示整个匹配项，1 表示第一个捕获组，以此类推。
- **示例**:
  - 提取合同编号 "HT-2023-001":
    - `pattern`: `(HT-\d{4}-\d{3})`
    - `group`: `1` (或 `0` 如果不需要括号内的特定部分)

#### 2. 位置 (POSITION)

根据文本在文档中的起止字符位置提取。**注意**：这种方式对文档结构变化非常敏感。

- **配置参数**:
  - `start_index` (整数, 可选, 默认 0): 提取开始的字符索引（从0开始）。
  - `end_index` (整数, 可选): 提取结束的字符索引（不包含此索引处的字符）。如果省略，则提取到文本末尾。
- **示例**:
  - 提取文档前10个字符:
    - `start_index`: `0`
    - `end_index`: `10`

#### 3. 书签 (BOOKMARK)

提取 Word 文档中预定义书签所标记的内容。

- **配置参数**:
  - `bookmark_name` (字符串): Word 文档中书签的名称。
- **示例**:
  - 提取名为 "ApplicantName" 的书签内容:
    - `bookmark_name`: `ApplicantName`

#### 4. 表格单元格 (TABLE_CELL)

提取指定表格中特定单元格的内容。

- **配置参数**:
  - `table_index` (整数, 可选, 默认 0): 表格在文档中的索引（从0开始，即第一个表格为0）。
  - `row_index` (整数, 可选, 默认 0): 单元格所在行的索引（从0开始）。
  - `column_index` (整数, 可选, 默认 0): 单元格所在列的索引（从0开始）。
- **示例**:
  - 提取文档中第一个表格的第2行第3列的单元格内容:
    - `table_index`: `0`
    - `row_index`: `1`
    - `column_index`: `2`

#### 5. 表格列 (TABLE_COLUMN)

提取指定表格中特定列的所有单元格内容。结果通常是一个列表或以特定分隔符连接的字符串。

- **配置参数**:
  - `table_index` (整数, 可选, 默认 0): 表格索引。
  - `column_index` (整数, 可选, 默认 0): 列索引。
  - `has_header` (布尔值, 可选, 默认 `True`): 表格是否有标题行。如果为 `True`，则提取时不包含第一行（假定为表头）。
- **示例**:
  - 提取第一个表格的第二列数据（不包括表头）:
    - `table_index`: `0`
    - `column_index`: `1`
    - `has_header`: `True`

#### 6. 表格行 (TABLE_ROW)

提取指定表格中特定行的所有单元格内容。结果通常是一个列表或以特定分隔符连接的字符串。

- **配置参数**:
  - `table_index` (整数, 可选, 默认 0): 表格索引。
  - `row_index` (整数, 可选, 默认 0): 行索引。
- **示例**:
  - 提取第一个表格的第三行数据:
    - `table_index`: `0`
    - `row_index`: `2`

#### 7. 完整表格 (TABLE_FULL)

提取指定表格的全部内容。结果可能是多维列表或特定格式的文本。

- **配置参数**:
  - `table_index` (整数, 可选, 默认 0): 表格索引。
  - `has_header` (布尔值, 可选, 默认 `True`): 表格是否有标题行。影响数据的解析方式。
- **示例**:
  - 提取第一个表格的全部内容（包含表头）:
    - `table_index`: `0`
    - `has_header`: `False` (如果希望表头也作为数据一部分) 或 `True` (如果希望分别处理表头)

### 规则属性

- **启用/禁用**: 每条规则都可以被设置为启用或禁用。禁用的规则在批量处理时会被跳过。
- **顺序**: 规则列表中的顺序很重要，因为它们会按照这个顺序依次应用于文档。某些情况下，一个规则的输出可能影响后续规则的上下文（尽管在此工具中，规则通常是独立应用的）。

### 保存和加载规则

- **保存规则**: 您可以将当前配置的一组规则保存到文件中（通常是 `.json` 格式）。这允许您为不同类型的文档或提取任务创建不同的规则集。
- **加载规则**: 您可以从之前保存的文件中加载规则集，而不必每次都重新创建。

## 从源码构建 (可选)

如果您希望从源代码运行或构建此应用：

1. **克隆仓库**:

    ```bash
    git clone https://github.com/xihan123/WordExtractor
    cd WordExtractor
    ```

2. **设置 Python 环境**:
    建议使用虚拟环境。确保您安装了 Python (版本如 `requirements.txt` 或构建脚本中指定，例如 Python 3.13)。

    ```bash
    python -m venv venv
    # Windows
    venv\Scripts\activate
    # macOS/Linux
    # source venv/bin/activate
    ```

3. **安装依赖**:

    ```bash
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    ```

    如果需要构建可执行文件，还需安装 PyInstaller:

    ```bash
    pip install pyinstaller
    ```

4. **运行应用**:

    ```bash
    python main.py
    ```

5. **使用 PyInstaller 构建**:
    (参考 `.github/workflows/build.yml` 中的构建命令)

    ```bash
    pyinstaller --noconfirm --onefile --windowed --clean --uac-admin --icon=resources/app_icon.ico --name "WordExtractor" --add-data "resources;resources" main.py
    ```

    可执行文件将生成在 `dist` 目录下。

## 项目状态

![GitHub stars](https://img.shields.io/github/stars/xihan123/WordExtractor?style=social)
![GitHub forks](https://img.shields.io/github/forks/xihan123/WordExtractor?style=social)
![GitHub issues](https://img.shields.io/github/issues/xihan123/WordExtractor)
![GitHub pull requests](https://img.shields.io/github/issues-pr/xihan123/WordExtractor)

---

如果您觉得这个工具有用，请考虑给项目点个 ⭐️！
