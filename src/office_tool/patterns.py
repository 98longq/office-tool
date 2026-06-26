"""Regular expressions used by the document auditor and formatter."""

from __future__ import annotations

import re

CHINESE_NUM = r"[一二三四五六七八九十百千万零〇两]+"

RE_COPY_NUMBER = re.compile(r"^\d{6}$")
RE_SECRECY = re.compile(r"^(绝密|机密|秘密)(?:\s*★\s*\d+\s*年)?$")
RE_URGENCY = re.compile(r"^(特急|加急|急件|平急|特提|特急件)$")
RE_INTERNAL_NOTICE = re.compile(r"^[（(]?\s*内部资料[\s　]*不得外传\s*[）)]?$")
RE_DOCUMENT_NUMBER = re.compile(r"[一-鿿A-Za-z0-9]{0,24}[〔\[]\s*\d{4}\s*[〕\]]\s*\d+\s*号")
RE_LETTER_DOCUMENT_NUMBER = re.compile(r"[一-鿿A-Za-z0-9]{0,24}函[〔\[]\s*\d{4}\s*[〕\]]\s*\d+\s*号")
RE_LETTER_CONTACT = re.compile(r"^[（(]?\s*(?:联系人|电话)\s*[:：].*[）)]?$", re.S)
RE_MEETING_RED_HEAD = re.compile(r"^\s*会\s*议\s*纪\s*要\s*$")
RE_MEETING_NUMBER = re.compile(r"^[（(]\s*\d+\s*[）)]$")
RE_MEETING_ISSUE_LINE = re.compile(r"^(.+?)\s+(\d{4}年\d{1,2}月\d{1,2}日)$")
RE_MEETING_ATTENDEES = re.compile(r"^\s*出席\s*[:：]")
RE_DISTRIBUTION = re.compile(r"^\s*分送\s*[:：]")
RE_SIGNER = re.compile(r"签发人\s*[:：]\s*\S+")
RE_RED_HEAD_KEYWORDS = re.compile(
    r"(文件|办公室文件|委员会文件|人民政府文件|党组文件|党委文件|局文件|厅文件|部文件|令)$"
)
RE_TITLE_LIKE = re.compile(r"(关于.+的(?:通知|通报|报告|请示|批复|函|意见|决定|公告|通告|纪要)|.+(?:通知|通报|报告|请示|批复|函|意见|决定|公告|通告|纪要|令)$)")
TITLE_DATE_PART = (
    r"(?:\d{4}|[二〇零一二三四五六七八九十]{4})年"
    r"\s*(?:\d{1,2}|[一二三四五六七八九十]+)月"
    r"\s*(?:\d{1,2}|[一二三四五六七八九十]+)日"
)
RE_TITLE_DATE_LINE = re.compile(
    rf"^[（(]?\s*{TITLE_DATE_PART}(?:\s*[-—－至到]\s*{TITLE_DATE_PART})?\s*[）)]?$"
)
RE_HEADING_H1 = re.compile(rf"^{CHINESE_NUM}\s*、")
RE_HEADING_H2 = re.compile(rf"^[（(]\s*{CHINESE_NUM}\s*[）)]")
RE_HEADING_H3 = re.compile(r"^\d{1,2}\s*[.．、](?!\d)")
RE_HEADING_H4 = re.compile(r"^[（(]\s*\d+\s*[）)]")
RE_REGULATION_CODE = re.compile(r"^[A-Za-z][A-Za-z0-9]*/[A-Za-z0-9]+-\d{4}$")
RE_REGULATION_CHAPTER = re.compile(rf"^(第{CHINESE_NUM}章)(?:[\s　]+)?(.+)$")
RE_REGULATION_ARTICLE_PARTS = re.compile(rf"^(第{CHINESE_NUM}条)(?:[\s　]+)?(.*)$")
RE_MAIN_SEND = re.compile(r"^[^。；;!?！？]{2,80}[:：]$")
RE_ATTACHMENT_NOTE = re.compile(r"^附件(?:\s*\d+|\s*" + CHINESE_NUM + r")?\s*[:：]")
RE_ATTACHMENT_MARK = re.compile(r"^附件(?:\s*\d+|\s*" + CHINESE_NUM + r")?\s*[:：]?$")
RE_ATTACHMENT_END_PUNCT = re.compile(r"[。；;，,、.!！?？]$")
RE_DATE = re.compile(
    r"((?:\d{4}|[二〇零一二三四五六七八九十]{4})年\s*(?:\d{1,2}|" + CHINESE_NUM + r")月\s*(?:\d{1,2}|" + CHINESE_NUM + r")日)"
)
RE_ARABIC_DATE = re.compile(r"\d{4}年(?:0[1-9]|1[0-2])月(?:0[1-9]|[12]\d|3[01])日")
RE_COPY_TO = re.compile(r"^(抄送|抄报|发送)\s*[:：]")
RE_PRINT_ORG_DATE = re.compile(r"(印发|印制)\s*$|(?:\d{4}年\d{1,2}月\d{1,2}日印发)")
RE_OBSOLETE_SUBJECT = re.compile(r"^主题词\s*[:：]")
RE_BAD_EFFECTIVE_DATE = re.compile(r"自(?:本文件)?(?:发布|印发)之日起(?:执行|施行)")
RE_HALFWIDTH_PUNCT = re.compile(r"[,:;!?]")
RE_EMPTY = re.compile(r"^\s*$")


def is_heading(text: str) -> bool:
    stripped = text.strip()
    return bool(
        RE_HEADING_H1.match(stripped)
        or RE_HEADING_H2.match(stripped)
        or RE_HEADING_H3.match(stripped)
        or RE_HEADING_H4.match(stripped)
    )


def heading_role(text: str) -> str | None:
    stripped = text.strip()
    if RE_HEADING_H1.match(stripped):
        return "h1"
    if RE_HEADING_H2.match(stripped):
        return "h2"
    if RE_HEADING_H3.match(stripped):
        return "h3"
    if RE_HEADING_H4.match(stripped):
        return "h4"
    return None
