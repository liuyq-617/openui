# OpenUI — 支持 Word (`.doc` / `.docx`) 修订模式（Track Changes）的增强版 Open WebUI

本仓库包含为 Open WebUI (`:9120`) 扩展的 **Word (`.doc` / `.docx`) 原生修订模式（Track Changes / 审阅模式）读写引擎、内置工具、全局增强过滤器与完整后端代码**。

## 核心改动与目录结构

- [`doc_revision_engine.py`](./doc_revision_engine.py) / [`backend/open_webui/utils/doc_revision_engine.py`](./backend/open_webui/utils/doc_revision_engine.py)：
  - **原文档修订模式改写 (`revise_document_file`)**：将上传的 `.doc` 或 `.docx` 合同按字符级差异（`difflib.SequenceMatcher`）注入 OpenXML `<w:del>`、`<w:ins>`、`<w:trackRevisions/>` 及 `<w:revisionView/>` 标签，100% 保留原文档的字体、段落格式、表格及页眉页脚，并同时导出 `.doc`（Word 97-2003 二进制格式）与 `.docx` 文件。
  - **从零编写修订模式文档 (`create_document_file`)**：支持使用 `[-删除原文-]{+修改后内容+}` 或 `~~删除原文~~++修改后内容++` 行内修订语法或替换对照表直接生成带 Word 修订痕迹的 `.doc` 与 `.docx` 文件。
  - **完整文档解析 (`extract_full_text_from_doc_path`)**：按顺序提取 `.doc` / `.docx` 全部段落与表格内容。
- [`backend/open_webui/tools/builtin.py`](./backend/open_webui/tools/builtin.py) & [`backend/open_webui/utils/tools.py`](./backend/open_webui/utils/tools.py)：
  - 内置注册 `revise_doc_with_track_changes`、`write_doc_file`、`read_doc_file` 三个工具。
- [`plugins/tools/doc_revision_writer.py`](./plugins/tools/doc_revision_writer.py)：
  - Open WebUI Workspace Tool 插件版本。
- [`plugins/functions/doc_revision_auto_filter.py`](./plugins/functions/doc_revision_auto_filter.py)：
  - Open WebUI 全局 Filter 函数（自动挂载工具、开启附件 `context: full` 全文模式、并在出口保证或自动兜底生成 `.doc` / `.docx` 下载链接）。
- [`Modelfile`](./Modelfile)：
  - 法务审阅专用模型（`contract-pro:latest`）配置。
- [`Dockerfile`](./Dockerfile)：
  - 集成 LibreOffice、中文字体及 `.doc` 修订模式引擎的定制镜像构建文件。
