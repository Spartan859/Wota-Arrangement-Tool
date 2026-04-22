from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Tuple

SECTION_MAP = {
    "p": "前奏",
    "a": "A melo",
    "b": "B melo",
    "c": "C melo",
    "r": "副歌",
    "i": "间奏",
    "o": "尾奏",
}
PURE_ACTION_LYRIC = ("（纯动作/无歌词）", "（纯动作/无歌词）")
HEADERS = ["段落", "拍数", "日文歌词", "中文歌词", "技 / 动作编排", "备注"]
TITLE_PATTERN = re.compile(r"^(?P<name>.*)\s+\(BPM:\s*(?P<bpm>.*)\)$")


@dataclass
class Block:
    block_type: str
    beats: str
    lyrics: List[Tuple[str, str]]
    arrangement: str = ""
    remarks: str = ""


@dataclass
class ArrangementState:
    idx: int = 0
    current_lyrics: List[Tuple[str, str]] = field(default_factory=list)
    blocks: List[Block] = field(default_factory=list)


def resolve_section(block_code: str) -> str:
    return SECTION_MAP.get(block_code, block_code)
