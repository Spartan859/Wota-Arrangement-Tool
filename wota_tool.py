import sys
import io
import os
import copy
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# 强制 UTF-8
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding='utf-8')

SECTION_MAP = {'p': '前奏', 'a': 'A melo', 'b': 'B melo', 'c': 'C melo', 'r': '副歌', 'i': '间奏', 'o': '尾奏'}

def main():
    print("\n" + "="*60)
    print("   🌟  Wota-Arrangement-Tool v1.0 🌟")
    print("="*60)
    
    # --- 阶段 0: 文件读取 ---
    lines = []
    while True:
        filepath = input("\n 📂 1. 请输入歌词txt路径: ").strip().strip('"').strip("'")
        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                lines = [line.strip() for line in f if line.strip()]
            if len(lines) >= 2: break
        print(" ❌ 路径无效或内容为空，请重试。")

    song_name = input(" 🎵 2. 请输入歌曲名称: ").strip() or "Untitled"
    bpm = input(" ⚡ 3. 请输入歌曲 BPM: ").strip() or "120"

    lyric_pairs = [ (lines[i], lines[i+1]) for i in range(0, len(lines)-1, 2) ]

    # --- 状态核心 ---
    # 我们把所有数据封装在一个 state 里，undo 时直接替换整个 state
    state = {
        'idx': 0,
        'current_lyrics': [],
        'blocks': []
    }
    history = [] # 存储 state 的快照

    def save_history():
        history.append(copy.deepcopy(state))

    def undo():
        nonlocal state
        if history:
            state = history.pop()
            print("\n   ↩️  撤销成功，已回退到上一步状态")
            return True
        print("\n   ⚠️  已经是最初状态，无法撤销")
        return False

    # --- 阶段 1: 前奏 ---
    print("\n" + "─"*30 + "\n 🎹 [前奏阶段] 输入 'p 拍数'，回车进入歌词")
    while True:
        cmd = input(f" [已记录 {len([b for b in state['blocks'] if b['type']=='前奏'])} 段前奏] >> ").strip().lower()
        if not cmd: break
        if cmd == 'u':
            undo()
            continue
        
        parts = cmd.split()
        if len(parts) >= 2 and parts[0] == 'p':
            save_history()
            state['blocks'].append({
                'type': '前奏', 'beats': parts[1], 
                'lyrics': [("（纯动作/无歌词）", "（纯动作/无歌词）")]
            })
        else: print("   ⚠️ 格式错误，例如 'p 8'")

    # --- 阶段 2: 歌词流水线 ---
    print("\n" + "═"*60 + "\n 【 歌词处理中 】 回车收纳 / 代码打包 / u 撤销\n" + "═"*60)

    while state['idx'] < len(lyric_pairs):
        jp, cn = lyric_pairs[state['idx']]
        print(f"\n 🟢 [进度: {state['idx']+1}/{len(lyric_pairs)}] " + (f"📦 暂存: {len(state['current_lyrics'])}句" if state['current_lyrics'] else ""))
        print(f" --------------------------------------------------\n  日: {jp}\n  中: {cn}\n --------------------------------------------------")
        
        cmd = input(" >> 指令: ").strip().lower()

        if cmd == 'u':
            undo()
            continue

        if cmd == "":
            save_history()
            state['current_lyrics'].append((jp, cn))
            state['idx'] += 1
        else:
            parts = cmd.split()
            if len(parts) >= 2:
                save_history()
                b_code, beats = parts[0], parts[1]
                # 兼容处理：如果是间奏且暂存为空
                if b_code == 'i' and not state['current_lyrics']:
                    state['blocks'].append({
                        'type': '间奏', 'beats': beats, 'lyrics': [("（纯动作/无歌词）", "（纯动作/无歌词）")]
                    })
                else:
                    # 正常的打包逻辑
                    state['current_lyrics'].append((jp, cn))
                    state['blocks'].append({
                        'type': SECTION_MAP.get(b_code, b_code),
                        'beats': beats,
                        'lyrics': state['current_lyrics'].copy()
                    })
                    state['current_lyrics'] = []
                    state['idx'] += 1
            else:
                print(" ⚠️ 格式不对（例：r 16）")

    # --- 阶段 3: 扫尾与尾奏 (逻辑简化，统一使用 state 控制) ---
    while state['current_lyrics']:
        print(f"\n 📢 剩余 {len(state['current_lyrics'])} 句歌词，请输入 [类型 拍数] 打包，或 'u' 撤销:")
        cmd = input(" >> ").strip().lower()
        if cmd == 'u': 
            undo()
            if state['idx'] < len(lyric_pairs): break # 滚回主循环
            continue
        parts = cmd.split()
        if len(parts) >= 2:
            save_history()
            state['blocks'].append({
                'type': SECTION_MAP.get(parts[0], parts[0]), 'beats': parts[1], 
                'lyrics': state['current_lyrics'].copy()
            })
            state['current_lyrics'] = []

    print("\n" + "─"*30 + "\n 🎸 [尾奏阶段] 输入 'o 拍数'，直接回车保存")
    while True:
        cmd = input(f" [已记录 {len([b for b in state['blocks'] if b['type']=='尾奏'])} 段尾奏] >> ").strip().lower()
        if not cmd: break
        if cmd == 'u':
            undo()
            continue
        parts = cmd.split()
        if len(parts) >= 2 and parts[0] == 'o':
            save_history()
            state['blocks'].append({
                'type': '尾奏', 'beats': parts[1], 
                'lyrics': [("（纯动作/无歌词）", "（纯动作/无歌词）")]
            })

    # --- 阶段 4: Excel 渲染 (逻辑与之前一致，增强保存鲁棒性) ---
    out_name = f"{song_name}_编排表.xlsx"
    for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']: out_name = out_name.replace(char, '_')

    while True:
        try:
            if os.path.exists(out_name): os.remove(out_name)
            wb = Workbook()
            ws = wb.active
            ws.title = "打艺编排脚本"

            # 样式定义
            fill_odd = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
            fill_even = PatternFill(start_color="EFF3F6", end_color="EFF3F6", fill_type="solid")
            header_fill = PatternFill(start_color="1F2D3D", end_color="1F2D3D", fill_type="solid")
            border = Border(left=Side(style='thin', color="DCDFE6"), right=Side(style='thin', color="DCDFE6"),
                            top=Side(style='thin', color="DCDFE6"), bottom=Side(style='thin', color="DCDFE6"))

            # 写表头
            ws.merge_cells('A1:F1')
            ws['A1'] = f"{song_name} (BPM: {bpm})"
            ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
            ws['A1'].font = Font(size=18, bold=True)
            
            headers = ["段落", "拍数", "日文歌词", "中文歌词", "技 / 动作编排", "备注"]
            for col, h in enumerate(headers, 1):
                cell = ws.cell(row=2, column=col, value=h)
                cell.fill, cell.font = header_fill, Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # 填充数据
            curr_row = 3
            block_row_info = []
            last_color = fill_even
            for idx, b in enumerate(state['blocks']):
                color = last_color if (idx > 0 and b['type'] == state['blocks'][idx-1]['type']) else (fill_odd if last_color == fill_even else fill_even)
                last_color = color
                start_r = curr_row
                for jp, cn in b['lyrics']:
                    ws.cell(row=curr_row, column=3, value=jp)
                    ws.cell(row=curr_row, column=4, value=cn)
                    for c in range(1, 7):
                        cell = ws.cell(row=curr_row, column=c)
                        cell.fill, cell.border = color, border
                        cell.alignment = Alignment(vertical='center', wrap_text=True, indent=1)
                    ws.row_dimensions[curr_row].height = 30
                    curr_row += 1
                block_row_info.append((start_r, curr_row - 1))

            # 合并逻辑 (Column 1, 2, 5, 6)
            for idx, b in enumerate(state['blocks']):
                sr, er = block_row_info[idx]
                for col in [2, 5, 6]:
                    if er > sr: ws.merge_cells(start_row=sr, start_column=col, end_row=er, end_column=col)
                    if col == 2: ws.cell(row=sr, column=col, value=b['beats']).alignment = Alignment(horizontal='center', vertical='center')

            # 智能合并段落列
            i = 0
            while i < len(state['blocks']):
                j = i
                while j + 1 < len(state['blocks']) and state['blocks'][j+1]['type'] == state['blocks'][i]['type']: j += 1
                sm, em = block_row_info[i][0], block_row_info[j][1]
                if em > sm: ws.merge_cells(start_row=sm, start_column=1, end_row=em, end_column=1)
                cell = ws.cell(row=sm, column=1, value=state['blocks'][i]['type'])
                cell.font = Font(bold=True, color="409EFF")
                cell.alignment = Alignment(horizontal='center', vertical='center')
                i = j + 1

            ws.column_dimensions['C'].width = ws.column_dimensions['D'].width = 40
            wb.save(out_name)
            print(f"\n ✨ 成功！文件已保存至: {out_name}")
            break
        except PermissionError:
            input(f"\n ⚠️  文件 [{out_name}] 被占用，请关闭它后按回车重试...")

if __name__ == "__main__":
    try: main()
    except KeyboardInterrupt: print("\n 👋 已退出")