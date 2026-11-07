"""
title: Word (.doc/.docx) 修订模式全局增强过滤器
author: Antigravity
version: 1.0.0
description: 自动为所有对话挂载 Word 修订模式读写工具、启用完整文档上下文，并在回复出口确保始终提供 .doc 与 .docx 修订版下载链接
"""

import json
import re
import sys
from typing import Optional

if "/app/backend/data" not in sys.path:
    sys.path.insert(0, "/app/backend/data")

import doc_revision_engine as eng

SYSTEM_INJECTION = (
    "\n\n【Word (.doc / .docx) 修订模式读写能力说明】\n"
    "你已原生接入 Word 文档生成与修订模式（Track Changes / 审阅模式）工具：\n"
    "1. 当用户要求对已上传的 .doc/.docx 合同或文档按『修订模式 / 审阅模式 / Track Changes』进行修改、术语替换并输出文件时，"
    "请直接调用 `revise_doc_with_track_changes` 工具，传入需要替换的 JSON 映射（如 `'{\"原文\": \"修改后\"}'`）。\n"
    "2. 当用户要求编写、起草或导出新的 .doc / .docx 文件（包括用修订模式编写新文档）时，请直接调用 `write_doc_file` 工具。"
    "在 `write_doc_file` 的 `content` 中可使用 `[-删除的原内容-]{+新增或修改的内容+}` 语法直接生成 Word 原生修订痕迹。\n"
    "3. **严禁**向用户回答“我无法直接生成 .doc 文件”或“无法生成带修订模式的 .doc 文件”。调用工具后，必须在回复中**原样输出**工具返回的 `download_markdown` 下载链接。"
)


def _append_markdown_to_message(msg: dict, md_block: str) -> None:
    if not md_block:
        return
    cur_content = msg.get("content") or ""
    if cur_content:
        msg["content"] = cur_content.rstrip() + "\n\n---\n\n" + md_block
    else:
        msg["content"] = md_block

    output = msg.get("output")
    if isinstance(output, list):
        appended = False
        for item in reversed(output):
            if isinstance(item, dict) and item.get("type") == "message":
                parts = item.get("content")
                if isinstance(parts, list) and parts:
                    last_part = parts[-1]
                    if isinstance(last_part, dict) and "text" in last_part:
                        last_part["text"] = (last_part.get("text") or "").rstrip() + "\n\n---\n\n" + md_block
                        appended = True
                        break
        if not appended:
            output.append(
                {
                    "type": "message",
                    "status": "completed",
                    "content": [{"type": "output_text", "text": msg["content"]}],
                }
            )


def _strip_refusal_lines(text: str) -> str:
    patterns = [
        r"很抱歉，我目前无法直接生成[^\n]*\n*",
        r"很遗憾，目前仍然无法[^\n]*\n*",
        r"I can't directly create a \.doc file[^\n]*\n*",
    ]
    cleaned = text
    for pat in patterns:
        cleaned = re.sub(pat, "", cleaned, flags=re.IGNORECASE)
    return cleaned.strip()


class Filter:
    def __init__(self):
        self.toggle = False

    def inlet(
        self,
        body: dict,
        __user__: Optional[dict] = None,
        __metadata__: Optional[dict] = None,
    ) -> dict:
        # 1. Always attach doc_revision_writer tool
        tool_ids = list(body.get("tool_ids") or [])
        if "doc_revision_writer" not in tool_ids:
            tool_ids.append("doc_revision_writer")
        body["tool_ids"] = tool_ids

        # 2. Ensure full context for attached files so contracts are not truncated
        for f_list in (body.get("files"), (__metadata__ or {}).get("files")):
            if isinstance(f_list, list):
                for f in f_list:
                    if isinstance(f, dict):
                        f["context"] = "full"

        # 3. Inject system capability reminder
        messages = body.get("messages") or []
        if messages:
            sys_msg = next((m for m in messages if m.get("role") == "system"), None)
            if sys_msg is not None:
                if "revise_doc_with_track_changes" not in (sys_msg.get("content") or ""):
                    sys_msg["content"] = (sys_msg.get("content") or "") + SYSTEM_INJECTION
            else:
                messages.insert(
                    0,
                    {
                        "role": "system",
                        "content": "你是一名资深的企业法务与技术商务合同专家。" + SYSTEM_INJECTION,
                    },
                )
            body["messages"] = messages

        return body

    async def outlet(
        self,
        body: dict,
        __user__: Optional[dict] = None,
        __metadata__: Optional[dict] = None,
        __event_emitter__=None,
    ) -> dict:
        messages = body.get("messages") or []
        if not messages:
            return body

        last_msg = messages[-1]
        if last_msg.get("role") != "assistant":
            return body

        assistant_text = last_msg.get("content") or ""
        output_items = last_msg.get("output") or []

        # 1. Check if a tool call in this turn already produced download_markdown
        tool_download_mds = []
        if isinstance(output_items, list):
            for item in output_items:
                if isinstance(item, dict) and item.get("type") == "function_call_output":
                    for part in item.get("output") or []:
                        txt = part.get("text") if isinstance(part, dict) else ""
                        if txt and "download_markdown" in txt:
                            try:
                                data = json.loads(txt)
                                if isinstance(data, dict) and data.get("download_markdown"):
                                    tool_download_mds.append(data["download_markdown"])
                            except Exception:
                                pass

        if tool_download_mds:
            for md in tool_download_mds:
                # Extract file URL to verify if already included in assistant_text
                m_url = re.search(r"/api/v1/files/[0-9a-fA-F-]+/content", md)
                if not m_url or m_url.group(0) not in assistant_text:
                    _append_markdown_to_message(last_msg, md)
                    assistant_text = last_msg.get("content") or ""
            return body

        # 2. Fallback: If the model didn't call the tool, check if user asked for .doc / 修订模式 / 审阅模式
        user_msgs = [m for m in messages if m.get("role") == "user"]
        last_user_text = (user_msgs[-1].get("content") if user_msgs else "") or ""
        if isinstance(last_user_text, list):
            last_user_text = " ".join(p.get("text", "") for p in last_user_text if isinstance(p, dict))

        combined_check = (last_user_text + " " + assistant_text).lower()
        wants_doc = any(
            kw in combined_check
            for kw in (".doc", "doc ", "doc文件", "docx", "word", "修订模式", "审阅模式", "track changes", "直接给我修改好的")
        )
        if not wants_doc:
            return body

        user_id = (__user__ or {}).get("id")
        chat_id = body.get("chat_id") or (__metadata__ or {}).get("chat_id")

        # Clean any refusal banner
        cleaned_text = _strip_refusal_lines(assistant_text)
        if cleaned_text != assistant_text:
            last_msg["content"] = cleaned_text
            if isinstance(output_items, list):
                for item in output_items:
                    if isinstance(item, dict) and item.get("type") == "message":
                        for part in item.get("content") or []:
                            if isinstance(part, dict) and "text" in part:
                                part["text"] = _strip_refusal_lines(part.get("text") or "")

        # Try extracting replacements from assistant text (e.g. Markdown replacement table)
        repls = eng.parse_replacements(assistant_text)
        target = eng.resolve_uploaded_doc_file(
            files_meta=(__metadata__ or {}).get("files"),
            messages=messages,
            chat_id=chat_id,
            user_id=user_id,
        )

        try:
            if repls and target:
                res = eng.revise_document_file(
                    source_path=target["path"],
                    original_filename=target["filename"],
                    raw_replacements=repls,
                    author="AI法务审阅",
                    user_id=user_id,
                )
                _append_markdown_to_message(last_msg, res["download_markdown"])
            elif cleaned_text and len(cleaned_text) > 30:
                res = eng.create_document_file(
                    title="审阅修订文档",
                    content=cleaned_text,
                    filename="审阅修订文档",
                    author="AI法务审阅",
                    raw_replacements=repls if repls else None,
                    user_id=user_id,
                )
                _append_markdown_to_message(last_msg, res["download_markdown"])
        except Exception:
            pass

        return body
