from __future__ import annotations

import copy
import io
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

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


class WotaArrangementTool:
    def __init__(self) -> None:
        self.song_name = ""
        self.bpm = ""
        self.lyric_pairs: List[Tuple[str, str]] = []
        self.state = ArrangementState()
        self.history: List[ArrangementState] = []
        self.redo_history: List[ArrangementState] = []
        self.default_output_path: Path | None = None

    @staticmethod
    def configure_stdio() -> None:
        if sys.platform == "win32" and hasattr(sys.stdout, "buffer") and hasattr(sys.stdin, "buffer"):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
            sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8")

    def run(self) -> None:
        self._print_banner()
        mode = self._choose_mode()
        if mode == "xlsx_edit":
            self._load_from_xlsx_prompt()
            self._run_edit_session()
            self._save_workbook_with_retry()
            return

        self._load_inputs_from_txt()
        self._run_intro_phase()
        self._run_lyric_pipeline()
        self._run_tail_pack_phase()
        self._run_outro_phase()

        if mode == "txt_edit":
            self._run_edit_session()
        self._save_workbook_with_retry()

    @staticmethod
    def _print_banner() -> None:
        print("\n" + "=" * 60)
        print("   🌟  Wota-Arrangement-Tool v1.1 🌟")
        print("=" * 60)

    def _choose_mode(self) -> str:
        print("\n 请选择启动模式：")
        print("  1) 新建编排（原流程）")
        print("  2) 读取已生成 xlsx 继续编辑")
        print("  3) 从 txt 新建并进入完整编辑模式")
        while True:
            choice = input(" >> 模式编号 [默认 1]: ").strip() or "1"
            if choice == "1":
                return "new"
            if choice == "2":
                return "xlsx_edit"
            if choice == "3":
                return "txt_edit"
            print(" ⚠️ 无效输入，请输入 1 / 2 / 3")

    def _load_inputs_from_txt(self) -> None:
        while True:
            filepath = input("\n 📂 1. 请输入歌词txt路径: ").strip().strip('"').strip("'")
            path = Path(filepath)
            try:
                self.lyric_pairs = self._read_lyric_pairs_from_txt_path(path)
                break
            except Exception as exc:
                print(f" ❌ {exc}")

        self.song_name = input(" 🎵 2. 请输入歌曲名称: ").strip() or "Untitled"
        self.bpm = input(" ⚡ 3. 请输入歌曲 BPM: ").strip() or "120"

    @staticmethod
    def _read_lyric_pairs_from_txt_path(path: Path) -> List[Tuple[str, str]]:
        if not path.exists():
            raise FileNotFoundError("路径不存在，请重试。")
        if not path.is_file():
            raise ValueError("路径不是文件，请重试。")

        with path.open("r", encoding="utf-8") as file:
            lines = [line.strip() for line in file if line.strip()]
        if len(lines) < 2:
            raise ValueError("歌词内容不足（至少两行）。")
        return [(lines[i], lines[i + 1]) for i in range(0, len(lines) - 1, 2)]

    @staticmethod
    def _normalize_nav_command(raw_cmd: str) -> str:
        cmd = raw_cmd.strip()
        lowered = cmd.lower()
        if lowered in {"\x1b[a", "{up}", "up", "↑", "k"}:
            return "u"
        if lowered in {"\x1b[b", "{down}", "down", "↓", "j", "skip"}:
            return "skip"
        return lowered

    @staticmethod
    def _cell_str(value: object) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @classmethod
    def _parse_title(cls, raw_title: str) -> tuple[str, str]:
        match = TITLE_PATTERN.match(raw_title.strip())
        if not match:
            return raw_title.strip() or "Untitled", "120"
        return match.group("name").strip() or "Untitled", match.group("bpm").strip() or "120"

    @classmethod
    def _read_state_from_xlsx(cls, path: Path) -> tuple[str, str, ArrangementState]:
        workbook = load_workbook(path)
        worksheet = workbook.active

        actual_headers = [cls._cell_str(worksheet.cell(row=2, column=col).value) for col in range(1, 7)]
        if actual_headers != HEADERS:
            raise ValueError("xlsx 表头不匹配，当前仅支持本工具生成的标准模板。")

        raw_title = cls._cell_str(worksheet["A1"].value)
        song_name, bpm = cls._parse_title(raw_title)

        blocks: List[Block] = []
        current_section = ""
        current_block: Block | None = None

        for row in range(3, worksheet.max_row + 1):
            values = [cls._cell_str(worksheet.cell(row=row, column=col).value) for col in range(1, 7)]
            if not any(values):
                continue

            section, beats, jp, cn, arrangement, remarks = values
            if section:
                current_section = section

            if beats:
                if not current_section:
                    raise ValueError(f"第 {row} 行拍数存在，但段落为空，无法解析。")
                current_block = Block(
                    block_type=current_section,
                    beats=beats,
                    lyrics=[],
                    arrangement=arrangement,
                    remarks=remarks,
                )
                blocks.append(current_block)
            elif current_block is None:
                raise ValueError(f"第 {row} 行无法定位所属段落，xlsx 合并结构可能已损坏。")

            if current_block is not None and (jp or cn):
                current_block.lyrics.append((jp, cn))

        if not blocks:
            raise ValueError("xlsx 中没有可解析的段落数据。")

        for block in blocks:
            if not block.lyrics:
                block.lyrics = [PURE_ACTION_LYRIC]

        return song_name, bpm, ArrangementState(idx=0, current_lyrics=[], blocks=blocks)

    def _load_from_xlsx_prompt(self) -> None:
        while True:
            file_path = input("\n 📂 请输入已生成的编排表 xlsx 路径: ").strip().strip('"').strip("'")
            path = Path(file_path)
            if not path.exists():
                print(" ❌ 文件不存在，请重试。")
                continue
            try:
                self.song_name, self.bpm, self.state = self._read_state_from_xlsx(path)
            except Exception as exc:
                print(f" ❌ 解析失败: {exc}")
                continue
            self.default_output_path = path
            self.history = []
            self.redo_history = []
            print(f" ✅ 已载入：{path.name}（段落 {len(self.state.blocks)} 个）")
            return

    def _record_snapshot(self) -> None:
        self.history.append(copy.deepcopy(self.state))
        self.redo_history = []

    def _undo(self) -> bool:
        if self.history:
            self.redo_history.append(copy.deepcopy(self.state))
            self.state = self.history.pop()
            print("\n   ↩️ 撤销成功")
            return True
        print("\n   ⚠️ 已经是最初状态，无法撤销")
        return False

    def _redo(self) -> bool:
        if self.redo_history:
            self.history.append(copy.deepcopy(self.state))
            self.state = self.redo_history.pop()
            print("\n   ↪️ 重做成功")
            return True
        print("\n   ⚠️ 没有可重做的操作")
        return False

    @staticmethod
    def _resolve_section(block_code: str) -> str:
        return SECTION_MAP.get(block_code, block_code)

    def _append_pure_action_block(self, block_type: str, beats: str) -> None:
        self.state.blocks.append(Block(block_type=block_type, beats=beats, lyrics=[PURE_ACTION_LYRIC]))

    def _run_intro_phase(self) -> None:
        print("\n" + "─" * 30 + "\n 🎹 [前奏阶段] 输入 'p 拍数'，回车进入歌词")
        while True:
            intro_count = sum(1 for block in self.state.blocks if block.block_type == "前奏")
            cmd = input(f" [已记录 {intro_count} 段前奏] >> ").strip().lower()
            if not cmd:
                return
            if cmd == "u":
                self._undo()
                continue

            parts = cmd.split()
            if len(parts) >= 2 and parts[0] == "p":
                self._record_snapshot()
                self._append_pure_action_block("前奏", parts[1])
            else:
                print("   ⚠️ 格式错误，例如 'p 8'")

    def _run_lyric_pipeline(self) -> None:
        print("\n" + "═" * 60 + "\n 【 歌词处理中 】 回车收纳 / 代码打包 / u(或↑/k)撤销 / skip(或↓/j)跳过\n" + "═" * 60)
        while self.state.idx < len(self.lyric_pairs):
            jp, cn = self.lyric_pairs[self.state.idx]
            cache_info = f"📦 暂存: {len(self.state.current_lyrics)}句" if self.state.current_lyrics else ""
            print(f"\n 🟢 [进度: {self.state.idx + 1}/{len(self.lyric_pairs)}] {cache_info}")
            print(f" --------------------------------------------------\n  日: {jp}\n  中: {cn}\n --------------------------------------------------")
            raw_cmd = input(" >> 指令: ")
            cmd = self._normalize_nav_command(raw_cmd)

            if cmd == "u":
                self._undo()
                continue

            if raw_cmd.strip() == "":
                self._record_snapshot()
                self.state.current_lyrics.append((jp, cn))
                self.state.idx += 1
                continue

            if cmd == "skip":
                self._record_snapshot()
                self.state.idx += 1
                print(" ⏭️ 已跳过当前歌词行")
                continue

            parts = cmd.split()
            if len(parts) < 2:
                print(" ⚠️ 格式不对（例：r 16）")
                continue

            self._record_snapshot()
            block_code, beats = parts[0], parts[1]
            if block_code == "i" and not self.state.current_lyrics:
                self._append_pure_action_block("间奏", beats)
                continue

            self.state.current_lyrics.append((jp, cn))
            self.state.blocks.append(
                Block(
                    block_type=self._resolve_section(block_code),
                    beats=beats,
                    lyrics=self.state.current_lyrics.copy(),
                )
            )
            self.state.current_lyrics = []
            self.state.idx += 1

    def _run_tail_pack_phase(self) -> None:
        while self.state.current_lyrics:
            print(f"\n 📢 剩余 {len(self.state.current_lyrics)} 句歌词，请输入 [类型 拍数] 打包，或 'u' 撤销:")
            cmd = input(" >> ").strip().lower()
            if cmd == "u":
                self._undo()
                if self.state.idx < len(self.lyric_pairs):
                    break
                continue

            parts = cmd.split()
            if len(parts) < 2:
                print(" ⚠️ 格式不对（例：r 16）")
                continue

            self._record_snapshot()
            self.state.blocks.append(
                Block(
                    block_type=self._resolve_section(parts[0]),
                    beats=parts[1],
                    lyrics=self.state.current_lyrics.copy(),
                )
            )
            self.state.current_lyrics = []

    def _run_outro_phase(self) -> None:
        print("\n" + "─" * 30 + "\n 🎸 [尾奏阶段] 输入 'o 拍数'，直接回车保存")
        while True:
            outro_count = sum(1 for block in self.state.blocks if block.block_type == "尾奏")
            cmd = input(f" [已记录 {outro_count} 段尾奏] >> ").strip().lower()
            if not cmd:
                return
            if cmd == "u":
                self._undo()
                continue

            parts = cmd.split()
            if len(parts) >= 2 and parts[0] == "o":
                self._record_snapshot()
                self._append_pure_action_block("尾奏", parts[1])
            else:
                print("   ⚠️ 格式错误，例如 'o 4'")

    @staticmethod
    def _sanitize_output_name(song_name: str) -> str:
        output_name = f"{song_name}_编排表.xlsx"
        for char in ['\\', "/", ":", "*", "?", '"', "<", ">", "|"]:
            output_name = output_name.replace(char, "_")
        return output_name

    def _parse_index(self, idx_text: str) -> int:
        idx = int(idx_text)
        if idx < 1 or idx > len(self.state.blocks):
            raise ValueError("索引越界")
        return idx - 1

    def _parse_lyrics_payload(self, payload: str) -> tuple[str, str]:
        if "|" not in payload:
            raise ValueError("歌词参数格式应为 日文|中文")
        jp, cn = payload.split("|", 1)
        return jp.strip(), cn.strip()

    def _print_blocks(self) -> None:
        if not self.state.blocks:
            print(" （当前没有段落）")
            return
        print("\n idx | 段落 | 拍数 | 歌词行数 | 编排 | 备注")
        print("-----+------+-----+---------+------+-----")
        for i, block in enumerate(self.state.blocks, 1):
            print(
                f" {i:>3} | {block.block_type} | {block.beats} | {len(block.lyrics):>7} | "
                f"{block.arrangement or '-'} | {block.remarks or '-'}"
            )

    def _print_edit_help(self) -> None:
        print(
            "\n编辑命令：\n"
            "  ls\n"
            "  ins <idx> <type> <beats> [pure]\n"
            "  inject <idx> <txt_path>\n"
            "  del <idx>\n"
            "  set <idx> type|beats|arrangement|remarks <value>\n"
            "  mv <from> <to>\n"
            "  lyrics add <idx> <pos> <jp>|<cn>\n"
            "  lyrics set <idx> <pos> <jp>|<cn>\n"
            "  lyrics del <idx> <pos>\n"
            "  u / redo\n"
            "  save [path]\n"
            "  fix-merge\n"
            "  done\n"
            "\n注入子模式快捷键：Enter=收纳，u/↑/k=回退，skip/↓/j=跳过\n"
        )

    def _run_edit_session(self) -> None:
        print("\n" + "═" * 60 + "\n 【 完整编辑模式 】输入 help 查看命令\n" + "═" * 60)
        self._print_blocks()
        while True:
            raw_cmd = input(" edit >> ").strip()
            if not raw_cmd:
                continue
            cmd = raw_cmd.split()[0].lower()
            parts = raw_cmd.split(maxsplit=3)

            try:
                if cmd in {"help", "h"}:
                    self._print_edit_help()
                elif cmd == "ls":
                    self._print_blocks()
                elif cmd == "u":
                    self._undo()
                elif cmd == "redo":
                    self._redo()
                elif cmd == "done":
                    return
                elif cmd == "fix-merge":
                    print(" ℹ️ 保存时会自动重建合并结构，无需单独修复。")
                elif cmd == "inject":
                    inject_parts = raw_cmd.split(maxsplit=2)
                    if len(inject_parts) < 3:
                        raise ValueError("inject 用法: inject <idx> <txt_path>")
                    self._handle_inject_command(inject_parts[1], inject_parts[2])
                elif cmd == "del" and len(parts) >= 2:
                    block_idx = self._parse_index(parts[1])
                    self._record_snapshot()
                    deleted = self.state.blocks.pop(block_idx)
                    print(f" ✅ 已删除段落: {deleted.block_type}")
                elif cmd == "ins" and len(parts) >= 4:
                    insert_pos = int(parts[1])
                    if insert_pos < 1 or insert_pos > len(self.state.blocks) + 1:
                        raise ValueError("插入位置越界")
                    block_type = self._resolve_section(parts[2].lower())
                    beats = parts[3].split()[0]
                    tail = raw_cmd.split(maxsplit=4)
                    pure = len(tail) >= 5 and tail[4].strip().lower() == "pure"
                    lyrics = [PURE_ACTION_LYRIC] if pure else [("", "")]
                    self._record_snapshot()
                    self.state.blocks.insert(insert_pos - 1, Block(block_type=block_type, beats=beats, lyrics=lyrics))
                    print(" ✅ 插入成功")
                elif cmd == "set" and len(parts) >= 4:
                    block_idx = self._parse_index(parts[1])
                    field = parts[2].lower()
                    value = parts[3].strip()
                    if field not in {"type", "beats", "arrangement", "remarks"}:
                        raise ValueError("set 字段仅支持 type/beats/arrangement/remarks")
                    self._record_snapshot()
                    target = self.state.blocks[block_idx]
                    if field == "type":
                        target.block_type = self._resolve_section(value.lower())
                    elif field == "beats":
                        target.beats = value
                    elif field == "arrangement":
                        target.arrangement = value
                    else:
                        target.remarks = value
                    print(" ✅ 修改成功")
                elif cmd == "mv" and len(parts) >= 3:
                    from_idx = self._parse_index(parts[1])
                    to_idx = int(parts[2])
                    if to_idx < 1 or to_idx > len(self.state.blocks):
                        raise ValueError("目标位置越界")
                    self._record_snapshot()
                    block = self.state.blocks.pop(from_idx)
                    self.state.blocks.insert(to_idx - 1, block)
                    print(" ✅ 重排成功")
                elif cmd == "lyrics":
                    self._handle_lyrics_command(raw_cmd)
                elif cmd == "save":
                    save_parts = raw_cmd.split(maxsplit=1)
                    save_path = Path(save_parts[1].strip()) if len(save_parts) == 2 else None
                    self._save_workbook_with_retry(save_path)
                else:
                    print(" ⚠️ 未识别命令，输入 help 查看可用命令。")
            except Exception as exc:
                print(f" ⚠️ 命令失败: {exc}")

    def _handle_lyrics_command(self, raw_cmd: str) -> None:
        parts = raw_cmd.split(maxsplit=4)
        if len(parts) < 4:
            raise ValueError("lyrics 命令格式不正确")
        action = parts[1].lower()
        block_idx = self._parse_index(parts[2])
        block = self.state.blocks[block_idx]

        if action == "del":
            pos = int(parts[3])
            if pos < 1 or pos > len(block.lyrics):
                raise ValueError("歌词行索引越界")
            self._record_snapshot()
            block.lyrics.pop(pos - 1)
            if not block.lyrics:
                block.lyrics = [PURE_ACTION_LYRIC]
            print(" ✅ 歌词行删除成功")
            return

        if len(parts) < 5:
            raise ValueError("lyrics add/set 需要提供歌词内容")
        pos = int(parts[3])
        jp, cn = self._parse_lyrics_payload(parts[4])
        if action == "add":
            if pos < 1 or pos > len(block.lyrics) + 1:
                raise ValueError("插入位置越界")
            self._record_snapshot()
            if block.lyrics == [PURE_ACTION_LYRIC]:
                block.lyrics = []
            block.lyrics.insert(pos - 1, (jp, cn))
            print(" ✅ 歌词行插入成功")
            return
        if action == "set":
            if pos < 1 or pos > len(block.lyrics):
                raise ValueError("歌词行索引越界")
            self._record_snapshot()
            block.lyrics[pos - 1] = (jp, cn)
            print(" ✅ 歌词行修改成功")
            return

        raise ValueError("lyrics 子命令仅支持 add/set/del")

    def _insert_blocks_after(self, idx_1based: int, new_blocks: List[Block]) -> None:
        if idx_1based < 0 or idx_1based > len(self.state.blocks):
            raise ValueError("插入目标 idx 越界")
        self._record_snapshot()
        insert_pos = idx_1based
        for offset, block in enumerate(new_blocks):
            self.state.blocks.insert(insert_pos + offset, block)

    def _handle_inject_command(self, idx_text: str, txt_path_text: str) -> None:
        idx_1based = int(idx_text)
        if idx_1based < 1 or idx_1based > len(self.state.blocks):
            raise ValueError("inject 的 idx 需要是现有段落编号")
        txt_path = Path(txt_path_text.strip().strip('"').strip("'"))
        lyric_pairs = self._read_lyric_pairs_from_txt_path(txt_path)
        new_blocks = self._run_inject_pipeline(lyric_pairs)
        if not new_blocks:
            print(" ℹ️ 未生成任何新段落，取消插入。")
            return
        self._insert_blocks_after(idx_1based, new_blocks)
        print(f" ✅ 已在段落 {idx_1based} 后插入 {len(new_blocks)} 个新段落")

    def _run_inject_pipeline(self, lyric_pairs: List[Tuple[str, str]]) -> List[Block]:
        print("\n" + "═" * 60 + "\n 【 TXT 注入子模式 】回车收纳 / 代码打包 / u(↑/k)回退 / skip(↓/j)跳过\n" + "═" * 60)
        generated_blocks: List[Block] = []
        current_lyrics: List[Tuple[str, str]] = []
        idx = 0
        history: List[tuple[int, List[Tuple[str, str]], List[Block]]] = []

        def snapshot() -> None:
            history.append((idx, current_lyrics.copy(), copy.deepcopy(generated_blocks)))

        while True:
            while idx < len(lyric_pairs):
                jp, cn = lyric_pairs[idx]
                print(f"\n 🟢 [注入进度: {idx + 1}/{len(lyric_pairs)}] 暂存{len(current_lyrics)}句 已产出{len(generated_blocks)}段")
                print(f" --------------------------------------------------\n  日: {jp}\n  中: {cn}\n --------------------------------------------------")
                raw_cmd = input(" inject >> ")
                cmd = self._normalize_nav_command(raw_cmd)

                if cmd == "u":
                    if history:
                        old_idx, old_lyrics, old_blocks = history.pop()
                        idx = old_idx
                        current_lyrics = old_lyrics
                        generated_blocks = old_blocks
                        print(" ↩️ 已回退")
                    else:
                        print(" ⚠️ 已经是注入模式最初状态")
                    continue

                if raw_cmd.strip() == "":
                    snapshot()
                    current_lyrics.append((jp, cn))
                    idx += 1
                    continue

                if cmd == "skip":
                    snapshot()
                    idx += 1
                    print(" ⏭️ 已跳过当前歌词行")
                    continue

                parts = cmd.split()
                if len(parts) < 2:
                    print(" ⚠️ 格式不对（例：r 16）")
                    continue

                snapshot()
                block_code, beats = parts[0], parts[1]
                if block_code == "i" and not current_lyrics:
                    generated_blocks.append(Block(block_type="间奏", beats=beats, lyrics=[PURE_ACTION_LYRIC]))
                    continue

                current_lyrics.append((jp, cn))
                generated_blocks.append(
                    Block(
                        block_type=self._resolve_section(block_code),
                        beats=beats,
                        lyrics=current_lyrics.copy(),
                    )
                )
                current_lyrics = []
                idx += 1

            while current_lyrics:
                cmd = input(f"\n 📢 注入剩余 {len(current_lyrics)} 句，请输入 [类型 拍数] 打包，或 u 回退: ").strip()
                normalized = self._normalize_nav_command(cmd)
                if normalized == "u":
                    if history:
                        old_idx, old_lyrics, old_blocks = history.pop()
                        idx = old_idx
                        current_lyrics = old_lyrics
                        generated_blocks = old_blocks
                        print(" ↩️ 已回退")
                    else:
                        print(" ⚠️ 已经是注入模式最初状态")
                    continue
                parts = normalized.split()
                if len(parts) < 2:
                    print(" ⚠️ 格式不对（例：r 16）")
                    continue
                snapshot()
                generated_blocks.append(
                    Block(
                        block_type=self._resolve_section(parts[0]),
                        beats=parts[1],
                        lyrics=current_lyrics.copy(),
                    )
                )
                current_lyrics = []

            if idx >= len(lyric_pairs):
                return generated_blocks

    def _build_workbook(self) -> Workbook:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "打艺编排脚本"

        fill_odd = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
        fill_even = PatternFill(start_color="EFF3F6", end_color="EFF3F6", fill_type="solid")
        header_fill = PatternFill(start_color="1F2D3D", end_color="1F2D3D", fill_type="solid")
        border = Border(
            left=Side(style="thin", color="DCDFE6"),
            right=Side(style="thin", color="DCDFE6"),
            top=Side(style="thin", color="DCDFE6"),
            bottom=Side(style="thin", color="DCDFE6"),
        )

        worksheet.merge_cells("A1:F1")
        worksheet["A1"] = f"{self.song_name} (BPM: {self.bpm})"
        worksheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        worksheet["A1"].font = Font(size=18, bold=True)

        for col, header in enumerate(HEADERS, 1):
            cell = worksheet.cell(row=2, column=col, value=header)
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")

        curr_row = 3
        block_row_info: List[Tuple[int, int]] = []
        last_color = fill_even
        for idx, block in enumerate(self.state.blocks):
            same_as_prev = idx > 0 and block.block_type == self.state.blocks[idx - 1].block_type
            color = last_color if same_as_prev else (fill_odd if last_color == fill_even else fill_even)
            last_color = color
            start_row = curr_row
            for jp, cn in block.lyrics:
                worksheet.cell(row=curr_row, column=3, value=jp)
                worksheet.cell(row=curr_row, column=4, value=cn)
                for col in range(1, 7):
                    cell = worksheet.cell(row=curr_row, column=col)
                    cell.fill = color
                    cell.border = border
                    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
                worksheet.row_dimensions[curr_row].height = 30
                curr_row += 1
            block_row_info.append((start_row, curr_row - 1))

        for idx, block in enumerate(self.state.blocks):
            start_row, end_row = block_row_info[idx]
            for col in (2, 5, 6):
                if end_row > start_row:
                    worksheet.merge_cells(start_row=start_row, start_column=col, end_row=end_row, end_column=col)
            worksheet.cell(row=start_row, column=2, value=block.beats).alignment = Alignment(horizontal="center", vertical="center")
            worksheet.cell(row=start_row, column=5, value=block.arrangement).alignment = Alignment(horizontal="center", vertical="center")
            worksheet.cell(row=start_row, column=6, value=block.remarks).alignment = Alignment(horizontal="center", vertical="center")

        i = 0
        while i < len(self.state.blocks):
            j = i
            while j + 1 < len(self.state.blocks) and self.state.blocks[j + 1].block_type == self.state.blocks[i].block_type:
                j += 1
            start_merge = block_row_info[i][0]
            end_merge = block_row_info[j][1]
            if end_merge > start_merge:
                worksheet.merge_cells(start_row=start_merge, start_column=1, end_row=end_merge, end_column=1)
            section_cell = worksheet.cell(row=start_merge, column=1, value=self.state.blocks[i].block_type)
            section_cell.font = Font(bold=True, color="409EFF")
            section_cell.alignment = Alignment(horizontal="center", vertical="center")
            i = j + 1

        worksheet.column_dimensions["C"].width = 40
        worksheet.column_dimensions["D"].width = 40
        return workbook

    def _resolve_output_path(self, output_path: Path | None = None) -> Path:
        if output_path is not None:
            return output_path
        if self.default_output_path is not None:
            return self.default_output_path
        return Path(self._sanitize_output_name(self.song_name))

    def _save_workbook_with_retry(self, output_path: Path | None = None) -> None:
        target_path = self._resolve_output_path(output_path)
        while True:
            try:
                workbook = self._build_workbook()
                target_path.parent.mkdir(parents=True, exist_ok=True)
                workbook.save(target_path)
                self.default_output_path = target_path
                print(f"\n ✨ 成功！文件已保存至: {target_path}")
                return
            except PermissionError:
                input(f"\n ⚠️ 文件 [{target_path}] 被占用，请关闭它后按回车重试...")


def main() -> None:
    WotaArrangementTool.configure_stdio()
    tool = WotaArrangementTool()
    tool.run()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n 👋 已退出")
