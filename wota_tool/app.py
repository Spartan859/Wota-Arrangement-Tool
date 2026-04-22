from __future__ import annotations

import copy
import io
import sys
from pathlib import Path
from typing import List, Tuple

from .models import ArrangementState, Block, PURE_ACTION_LYRIC, resolve_section
from .xlsx_io import XlsxIO
from .pipeline import Pipeline
from .editor import Editor


class WotaArrangementTool(XlsxIO, Pipeline, Editor):
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
        return resolve_section(block_code)

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
