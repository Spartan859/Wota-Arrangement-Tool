# Copilot Instructions for Wota-Arrangement-Tool

## Build, test, and lint commands

This repository is a single interactive Python CLI project. It has dependency management via `requirements.txt`, but still does **not** define a formal test suite or lint pipeline.

Use these practical commands:

| Purpose | Command | Notes |
| --- | --- | --- |
| Install dependencies | `pip install -r requirements.txt` | Installs `openpyxl` used by Excel export. |
| Run tool | `python wota_tool.py` | Main workflow from README; interactive prompts for lyrics path, song name, BPM, and section commands. |
| Quick syntax check (single-file sanity) | `python -m py_compile wota_tool.py` | Useful as a lightweight check for the only code file. |
| Manual single-case test | `python wota_tool.py` | Use `lyrics.txt` with one JP/CN pair to validate one full interaction path and XLSX output. |

## High-level architecture

The application is a **stateful CLI pipeline** that turns bilingual lyric pairs into an Excel arrangement sheet:

1. **Input & parsing phase (`wota_tool.py`)**
   - Reads a UTF-8 text file path from user input.
   - Expects lyrics in alternating JP/CN lines and converts them into `(jp, cn)` pairs via `lyric_pairs`.
   - Collects song metadata (name, BPM).

2. **Arrangement state machine (`wota_tool.py`)**
   - Implemented as `WotaArrangementTool` with dataclass-based domain state (`ArrangementState`, `Block`).
   - Central mutable `state` object:
     - `idx`: current lyric pair cursor
     - `current_lyrics`: buffered lines not yet packaged
     - `blocks`: finalized arrangement blocks
   - Workflow phases: intro (`p`), lyric pipeline (Enter to accumulate, `<type> <beats>` to pack), tail-buffer flush, outro (`o`).
   - Undo is snapshot-based: each state-changing action stores `copy.deepcopy(state)` in `history`; `u` restores the previous snapshot.

3. **Excel rendering (`wota_tool.py`)**
   - Generates `<song_name>_编排表.xlsx` with sanitized filename characters.
   - Writes block rows, alternates fills by block groups, applies borders/alignment, and performs merge logic:
     - merges beats/arrangement/remarks columns within each block
     - merges section column across consecutive blocks of the same type
   - Retries save on `PermissionError` if the output file is open.

## Key conventions specific to this codebase

- **Section shorthand mapping is centralized in `SECTION_MAP`**: `p/a/b/c/r/i/o` map to display names; unknown section code is accepted as-is (custom section labels are a supported behavior).
- **Lyrics format convention is strict and domain-specific**: text input is expected as alternating Japanese/Chinese non-empty lines; processing logic operates on paired lines.
- **Pure-action blocks use a fixed placeholder lyric tuple**: `("（纯动作/无歌词）", "（纯动作/无歌词）")` for intro/interlude/outro blocks without lyrics.
- **Undo reliability depends on snapshot timing**: always snapshot before mutating `state`; this is an intentional pattern used across pipeline phases.
- **CLI interaction language and user-facing output are Chinese-first**; keep prompts/messages consistent with existing style when editing behavior.
