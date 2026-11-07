"""
title: Word (.doc/.docx) 修订模式读写工具
author: Antigravity
version: 1.0.0
description: 支持读取、从零编写、以及按 Word 修订模式（Track Changes / 审阅模式）修改生成 .doc 和 .docx 文件
"""

import json
import sys
from typing import Optional

if "/app/backend/data" not in sys.path:
    sys.path.insert(0, "/app/backend/data")

import doc_revision_engine as eng


class Tools:
    def __init__(self):
        pass

    async def revise_doc_with_track_changes(
        self,
        replacements: str,
        author: str = "AI法务审阅",
        output_filename: str = "",
        file_id: str = "",
        __user__: Optional[dict] = None,
        __files__: Optional[list] = None,
        __messages__: Optional[list] = None,
        __chat_id__: Optional[str] = None,
        __event_emitter__=None,
    ) -> str:
        """
        对已上传的 .doc 或 .docx 合同/文档执行 Word 修订模式（Track Changes / 审阅模式）修改。
        完全保留原文档的排版、段落格式、表格、页眉页脚与字体，将所有修改以 Word 原生修订标记（删除线+新增高亮）写入，
        同时生成并返回 .doc（Word 97-2003 二进制格式）和 .docx 两个文件的下载链接。

        :param replacements: 必填。需要替换的『原文 -> 修改后』映射。支持 JSON 对象字符串（例如 '{"专家咨询服务费": "业务咨询服务费", "推荐专家": "推荐人员", "专家": "咨询人员"}'）、JSON 数组或 Markdown 替换表格。
        :param author: 修订模式显示的审阅人名称，默认为 "AI法务审阅"。
        :param output_filename: 可选。输出文件名（不含扩展名），默认在原文件名后加 "_修订版"。
        :param file_id: 可选。目标文件 ID，留空会自动查找当前对话或最近上传的 .doc/.docx 文件。
        :return: 包含修订统计及 .doc / .docx 文件下载链接的 JSON 结果（请务必将返回的 download_markdown 链接原样展示给用户）。
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": "正在以 Word 修订模式（Track Changes）处理文档并生成 .doc / .docx 文件...",
                        "done": False,
                    },
                }
            )

        user_id = (__user__ or {}).get("id")
        target = eng.resolve_uploaded_doc_file(
            file_id=file_id,
            files_meta=__files__,
            messages=__messages__,
            chat_id=__chat_id__,
            user_id=user_id,
        )
        if not target:
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {"description": "未找到可修订的 .doc/.docx 原文件", "done": True},
                    }
                )
            return json.dumps(
                {
                    "status": "error",
                    "error": "未找到已上传的 .doc 或 .docx 原文件。如果是新建文档，请改用 write_doc_file 工具。",
                },
                ensure_ascii=False,
            )

        try:
            res = eng.revise_document_file(
                source_path=target["path"],
                original_filename=target["filename"],
                raw_replacements=replacements,
                author=author or "AI法务审阅",
                output_filename_stem=output_filename or None,
                user_id=user_id,
            )
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": f"已完成 {res['total_changes']} 处修订标记，已生成 .doc 与 .docx 文件",
                            "done": True,
                        },
                    }
                )
            res["instruction_for_assistant"] = (
                "文件已成功生成！你必须在回复正文中原样粘贴以下 download_markdown 中的 Markdown 下载链接，禁止省略或修改 URL："
            )
            return json.dumps(res, ensure_ascii=False)
        except Exception as e:
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {"description": f"生成修订文档出错: {e}", "done": True},
                    }
                )
            return json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False)

    async def write_doc_file(
        self,
        title: str,
        content: str,
        filename: str = "",
        author: str = "AI法务审阅",
        replacements: str = "",
        __user__: Optional[dict] = None,
        __event_emitter__=None,
    ) -> str:
        """
        从零编写并生成新的 .doc（Word 97-2003 格式）和 .docx 文件，支持普通编写模式与 Word 修订模式（Track Changes / 审阅模式）。
        如需在编写的文档中体现修订/审阅痕迹，可在 content 正文中直接使用行内修订语法：
        - 替换：`[-被删除的原文-]{+修改后的新内容+}` 或 `~~被删除的原文~~++修改后的新内容++`
        - 仅删除：`[-被删除的原文-]`
        - 仅新增：`{+新增的内容+}`
        同时也支持通过 replacements 参数传入 JSON 替换映射（如 '{"旧词": "新词"}'）自动在生成的文档中标记修订。

        :param title: 文档标题（如合同名称、方案名称）。
        :param content: 完整的文档正文内容（支持 Markdown 标题 #/##、加粗 **文本**、表格 | A | B | 以及行内修订语法 [-旧-]{+新+}）。
        :param filename: 可选。保存的文件名（不含 .doc/.docx 后缀），默认使用 title。
        :param author: 修订模式显示的审阅人/作者名称，默认为 "AI法务审阅"。
        :param replacements: 可选。额外需要在正文中批量标记为修订模式的 JSON 映射字符串（如 '{"原词": "新词"}'）。
        :return: 包含生成的 .doc 和 .docx 下载链接的 JSON 结果（请务必将返回的 download_markdown 链接原样展示给用户）。
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": "正在编写并生成支持修订模式的 .doc / .docx Word 文档...",
                        "done": False,
                    },
                }
            )

        user_id = (__user__ or {}).get("id")
        try:
            res = eng.create_document_file(
                title=title,
                content=content,
                filename=filename or title or "文档",
                author=author or "AI法务审阅",
                raw_replacements=replacements or None,
                user_id=user_id,
            )
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": f"成功生成 {res['doc_file']['filename']} 与 {res['docx_file']['filename']}",
                            "done": True,
                        },
                    }
                )
            res["instruction_for_assistant"] = (
                "文件已成功生成！你必须在回复正文中原样粘贴以下 download_markdown 中的 Markdown 下载链接，禁止省略或修改 URL："
            )
            return json.dumps(res, ensure_ascii=False)
        except Exception as e:
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {"description": f"编写 Word 文档出错: {e}", "done": True},
                    }
                )
            return json.dumps({"status": "error", "error": str(e)}, ensure_ascii=False)

    async def read_doc_file(
        self,
        file_id: str = "",
        __user__: Optional[dict] = None,
        __files__: Optional[list] = None,
        __messages__: Optional[list] = None,
        __chat_id__: Optional[str] = None,
    ) -> str:
        """
        读取已上传的 .doc 或 .docx 文件的完整文本内容（包括全部段落和所有表格）。

        :param file_id: 可选。文件 ID，留空时自动读取当前对话或最新上传的 .doc/.docx 文件。
        :return: 包含文档全部段落和表格的完整文本。
        """
        user_id = (__user__ or {}).get("id")
        target = eng.resolve_uploaded_doc_file(
            file_id=file_id,
            files_meta=__files__,
            messages=__messages__,
            chat_id=__chat_id__,
            user_id=user_id,
        )
        if not target:
            return "未找到已上传的 .doc 或 .docx 文件。"
        try:
            text = eng.extract_full_text_from_doc_path(target["path"])
            return f"文件名：{target['filename']} (file_id: {target['id']})\n\n{text}"
        except Exception as e:
            return f"读取文件失败: {e}"
