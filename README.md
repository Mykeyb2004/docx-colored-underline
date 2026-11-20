# docx-colored-underline

为 Word（`.docx`）文档中的所有彩色文本批量添加**双下划线**（或指定样式下划线），并另存为新文件。

特别适合高中/初中文综知识点、考点提纲、讲义的黑白打印准备与重点标记转换。

## 功能特性

- **零外部依赖**：仅基于 Python 标准库（`zipfile` + `xml.dom.minidom`），无需额外安装第三方包。
- **支持 `uv run` 即开即用**：内置 PEP 723 单文件脚本元数据，单行命令秒级执行。
- **智能样式解析**：
  - 自动识别直接设置的 RGB 字体颜色（如红 `#FF0000`、蓝 `#0000FF`、橙 `#ED7D31` 等）；
  - 支持 Word 主题颜色（如 `accent1`、`accent2` 等）；
  - 支持解析字符样式（`rStyle`）与段落样式（`pStyle`）中继承的颜色。
- **严格遵循 OOXML 规范**：按照 ECMA-376 规范维护 `<w:rPr>` 子元素排序，100% 保持原有排版、图片、表格及结构。
- **支持一键转纯黑**：提供 `--to-black` 参数，加上双下划线的同时将彩色文字转为黑色，便于黑白打印或复印。

---

## 快速使用

### 1. 使用 `uv run`（推荐）

在任何终端中直接执行：

```bash
# 基本用法：默认在原文件同级目录生成 <原文件名>_双下划线.docx
uv run underline_colored_text.py "你的文档.docx"

# 预览模式：仅统计彩色文本与颜色分布，不生成文件
uv run underline_colored_text.py "你的文档.docx" --dry-run

# 加双下划线的同时把彩色文本转为纯黑色（推荐打印使用）
uv run underline_colored_text.py "你的文档.docx" --to-black

# 指定自定义输出路径
uv run underline_colored_text.py "你的文档.docx" -o "输出文件.docx"
```

也可以通过项目命令运行：
```bash
uv run docx-colored-underline "你的文档.docx"
```

### 2. 传统 Python 执行

```bash
python3 underline_colored_text.py "你的文档.docx"
```

---

## 命令行参数一览

| 参数 | 说明 |
| :--- | :--- |
| `input` | 待处理的 `.docx` 文件路径（必填） |
| `-o, --output` | 自定义输出路径（默认：原目录下的 `<原文件名>_双下划线.docx`） |
| `-u, --underline` | 下划线样式（默认 `double`，可选 `single`, `thick`, `wave`, `dotted`, `dash` 等） |
| `--to-black` | 加下划线的同时将彩色文字转为 `#000000`（黑白打印场景推荐） |
| `--dry-run` | 仅检查并统计彩文字段，不写出任何文件 |
| `--include-headers-footers` | 连同页眉与页脚中的彩色文字一起处理 |
| `--include-notes` | 连同脚注、尾注与批注中的彩色文字一起处理 |
| `--include-whitespace` | 也为空白字符/空格加下划线（默认自动忽略纯空白） |
| `--ignore-colors` | 忽略的颜色列表（默认 `000000,AUTO`，例如可追加 `FFFFFF` 跳过白色） |

---

## 示例

```bash
# 示例 1: 查看知识点讲义中有哪些颜色
uv run underline_colored_text.py "高中政治必修3《政治与法治》核心知识点.docx" --dry-run

# 示例 2: 处理并另存为双下划线黑白打印稿
uv run underline_colored_text.py "高中政治必修3《政治与法治》核心知识点.docx" --to-black
```
