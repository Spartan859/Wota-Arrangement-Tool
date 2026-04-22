from __future__ import annotations

import copy
from typing import List, Tuple

from .models import ArrangementState, Block, resolve_section


class Pipeline:
    state: ArrangementState
    lyric_pairs: List[Tuple[str, str]]

    @staticmethod
    def _normalize_nav_command(raw_cmd: str) -> str:
        cmd = raw_cmd.strip()
        lowered = cmd.lower()
        if lowered in {"\x1b[a", "{up}", "up", "↑"}:
            return "u"
        if lowered in {"\x1b[b", "{down}", "down", "↓", "j", "skip"}:
            return "skip"
        return lowered

    @staticmethod
    def _print_staging_preview(staging: List[Tuple[str, str]]) -> None:
        if not staging:
            print(" 📦 暂存区: （空）")
            return
        print(f" 📦 暂存区: {len(staging)} 句")
        for i, (jp, cn) in enumerate(staging, 1):
            print(f"   [{i}] 日: {jp}")
            print(f"       中: {cn}")

    def _run_lyric_pipeline(self) -> None:
        print("\n" + "═" * 60 + "\n 【 歌词处理中 】 回车收纳 / 代码打包 / u(或↑)撤销 / skip(或↓/j)跳过\n" + "═" * 60)
        while self.state.idx < len(self.lyric_pairs):
            jp, cn = self.lyric_pairs[self.state.idx]
            print(f"\n 🟢 [进度: {self.state.idx + 1}/{len(self.lyric_pairs)}]")
            self._print_staging_preview(self.state.current_lyrics)
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
            block_lyrics = self.state.current_lyrics.copy() if self.state.current_lyrics else [("", "")]
            self.state.blocks.append(
                Block(
                    block_type=resolve_section(block_code),
                    beats=beats,
                    lyrics=block_lyrics,
                )
            )
            self.state.current_lyrics = []
            print(" ✅ 已打包暂存区；当前行已保留")

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
                    block_type=resolve_section(parts[0]),
                    beats=parts[1],
                    lyrics=self.state.current_lyrics.copy(),
                )
            )
            self.state.current_lyrics = []

    def _run_inject_pipeline(self, lyric_pairs: List[Tuple[str, str]]) -> List[Block]:
        print("\n" + "═" * 60 + "\n 【 TXT 注入子模式 】回车收纳 / 代码打包 / u(↑)回退 / skip(↓/j)跳过 / q退出\n" + "═" * 60)
        generated_blocks: List[Block] = []
        current_lyrics: List[Tuple[str, str]] = []
        idx = 0
        history: List[tuple[int, List[Tuple[str, str]], List[Block]]] = []

        def snapshot() -> None:
            history.append((idx, current_lyrics.copy(), copy.deepcopy(generated_blocks)))

        while True:
            while idx < len(lyric_pairs):
                jp, cn = lyric_pairs[idx]
                print(f"\n 🟢 [注入进度: {idx + 1}/{len(lyric_pairs)}] 已产出{len(generated_blocks)}段")
                self._print_staging_preview(current_lyrics)
                print(f" --------------------------------------------------\n  日: {jp}\n  中: {cn}\n --------------------------------------------------")
                raw_cmd = input(" inject >> ")
                cmd = self._normalize_nav_command(raw_cmd)

                if cmd in {"q", "quit"}:
                    return generated_blocks

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
                block_lyrics = current_lyrics.copy() if current_lyrics else [("", "")]
                generated_blocks.append(
                    Block(
                        block_type=resolve_section(block_code),
                        beats=beats,
                        lyrics=block_lyrics,
                    )
                )
                current_lyrics = []
                print(" ✅ 已打包暂存区；当前行已保留")

            while current_lyrics:
                cmd = input(f"\n 📢 注入剩余 {len(current_lyrics)} 句，请输入 [类型 拍数] 打包，或 u 回退: ").strip()
                normalized = self._normalize_nav_command(cmd)
                if normalized in {"q", "quit"}:
                    return generated_blocks
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
                        block_type=resolve_section(parts[0]),
                        beats=parts[1],
                        lyrics=current_lyrics.copy(),
                    )
                )
                current_lyrics = []

            if idx >= len(lyric_pairs):
                return generated_blocks
