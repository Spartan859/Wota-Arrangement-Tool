from __future__ import annotations

from pathlib import Path
from typing import List

from .models import ArrangementState, Block, PURE_ACTION_LYRIC, resolve_section


class Editor:
    state: ArrangementState

    def _parse_index(self, idx_text: str) -> int:
        idx = int(idx_text)
        if idx < 1 or idx > len(self.state.blocks):
            raise ValueError("索引越界")
        return idx - 1

    @staticmethod
    def _parse_lyrics_payload(payload: str) -> tuple[str, str]:
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

    def _print_block_lyrics(self, idx_1based: int) -> None:
        if idx_1based < 1 or idx_1based > len(self.state.blocks):
            raise ValueError("索引越界")
        block = self.state.blocks[idx_1based - 1]
        print(f"\n 📄 段落详情 idx={idx_1based} | 段落={block.block_type} | 拍数={block.beats}")
        if block.arrangement:
            print(f" 编排: {block.arrangement}")
        if block.remarks:
            print(f" 备注: {block.remarks}")
        print(" 歌词：")
        for i, (jp, cn) in enumerate(block.lyrics, 1):
            print(f"  [{i}] 日: {jp}")
            print(f"      中: {cn}")

    def _print_edit_help(self) -> None:
        print(
            "\n编辑命令：\n"
            "  ls\n"
            "  show <idx>\n"
            "  ins <idx> <type> <beats> [pure]   # 在 idx 后插入；idx=0 表示最前面\n"
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
            "\n注入子模式快捷键：Enter=收纳，u/↑=回退，skip/↓/j=跳过，q/quit=中途退出并保留已生成部分\n"
        )

    def _insert_blocks_after(self, idx_1based: int, new_blocks: List[Block]) -> None:
        if idx_1based < 0 or idx_1based > len(self.state.blocks):
            raise ValueError("插入目标 idx 越界")
        self._record_snapshot()
        insert_pos = idx_1based
        for offset, block in enumerate(new_blocks):
            self.state.blocks.insert(insert_pos + offset, block)

    def _handle_inject_command(self, idx_text: str, txt_path_text: str) -> None:
        idx_1based = int(idx_text)
        if idx_1based < 0 or idx_1based > len(self.state.blocks):
            raise ValueError("inject 的 idx 需在 0..当前段落数 范围内")
        txt_path = Path(txt_path_text.strip().strip('"').strip("'"))
        lyric_pairs = self._read_lyric_pairs_from_txt_path(txt_path)
        new_blocks = self._run_inject_pipeline(lyric_pairs)
        if not new_blocks:
            print(" ℹ️ 未生成任何新段落，取消插入。")
            return
        self._insert_blocks_after(idx_1based, new_blocks)
        print(f" ✅ 已在段落 {idx_1based} 后插入 {len(new_blocks)} 个新段落")

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
                elif cmd == "show" and len(parts) >= 2:
                    self._print_block_lyrics(int(parts[1]))
                elif cmd == "u":
                    if self._undo():
                        self._print_blocks()
                elif cmd == "redo":
                    if self._redo():
                        self._print_blocks()
                elif cmd == "done":
                    return
                elif cmd == "fix-merge":
                    print(" ℹ️ 保存时会自动重建合并结构，无需单独修复。")
                elif cmd == "inject":
                    inject_parts = raw_cmd.split(maxsplit=2)
                    if len(inject_parts) < 3:
                        raise ValueError("inject 用法: inject <idx> <txt_path>")
                    self._handle_inject_command(inject_parts[1], inject_parts[2])
                    self._print_blocks()
                elif cmd == "del" and len(parts) >= 2:
                    block_idx = self._parse_index(parts[1])
                    self._record_snapshot()
                    deleted = self.state.blocks.pop(block_idx)
                    print(f" ✅ 已删除段落: {deleted.block_type}")
                    self._print_blocks()
                elif cmd == "ins" and len(parts) >= 4:
                    insert_after_idx = int(parts[1])
                    if insert_after_idx < 0 or insert_after_idx > len(self.state.blocks):
                        raise ValueError("插入位置越界")
                    block_type = resolve_section(parts[2].lower())
                    beats = parts[3].split()[0]
                    tail = raw_cmd.split(maxsplit=4)
                    pure = len(tail) >= 5 and tail[4].strip().lower() == "pure"
                    lyrics = [PURE_ACTION_LYRIC] if pure else [("", "")]
                    self._record_snapshot()
                    self.state.blocks.insert(insert_after_idx, Block(block_type=block_type, beats=beats, lyrics=lyrics))
                    print(" ✅ 插入成功")
                    self._print_blocks()
                elif cmd == "set" and len(parts) >= 4:
                    block_idx = self._parse_index(parts[1])
                    field_name = parts[2].lower()
                    value = parts[3].strip()
                    if field_name not in {"type", "beats", "arrangement", "remarks"}:
                        raise ValueError("set 字段仅支持 type/beats/arrangement/remarks")
                    self._record_snapshot()
                    target = self.state.blocks[block_idx]
                    if field_name == "type":
                        target.block_type = resolve_section(value.lower())
                    elif field_name == "beats":
                        target.beats = value
                    elif field_name == "arrangement":
                        target.arrangement = value
                    else:
                        target.remarks = value
                    print(" ✅ 修改成功")
                    self._print_blocks()
                elif cmd == "mv" and len(parts) >= 3:
                    from_idx = self._parse_index(parts[1])
                    to_idx = int(parts[2])
                    if to_idx < 1 or to_idx > len(self.state.blocks):
                        raise ValueError("目标位置越界")
                    self._record_snapshot()
                    block = self.state.blocks.pop(from_idx)
                    self.state.blocks.insert(to_idx - 1, block)
                    print(" ✅ 重排成功")
                    self._print_blocks()
                elif cmd == "lyrics":
                    self._handle_lyrics_command(raw_cmd)
                    self._print_blocks()
                elif cmd == "save":
                    save_parts = raw_cmd.split(maxsplit=1)
                    save_path = Path(save_parts[1].strip()) if len(save_parts) == 2 else None
                    self._save_workbook_with_retry(save_path)
                else:
                    print(" ⚠️ 未识别命令，输入 help 查看可用命令。")
            except Exception as exc:
                print(f" ⚠️ 命令失败: {exc}")
