'''
安装依赖: pip install pyperclip
运行: 复制表格(roam不含{{[[table]]}}), 然后运行此脚本
'''
import pyperclip
import re
import copy


def roamToExcel(text, removeMarks=True):
    text = '\n' + text.replace('    ', '\t')
    if removeMarks:
        text = re.sub(r'[*]{2}|~~|\^\^|__', '', text)
    lines = re.sub(r'(?<=\n|\t)[-] ', '', text).strip().split('\n')
    table = []
    layer = 0
    while lines:
        line = lines.pop(0)
        if not table:
            table.append([line])
            continue
        tab_num = line.count('\t')
        line = line.strip('\t')
        if tab_num <= layer:
            table.append(copy.deepcopy(table[-1][0:tab_num]))
            table[-1].append(line)
            layer = tab_num
            continue
        table[-1].append(line)
        layer += 1
    max_len = max([len(i) for i in table])
    table_s = ''
    for i in table:
        i += [''] * (max_len - len(i))
        table_s += '\t'.join(i) + '\n'
    return table_s.strip()


def excelToRoam(text, mergeCellsColNum = -1):
    seg = ' ' * 4
    blankPad = '- '
    table = text.strip().split('\n')
    table = [[j if j.strip() else blankPad for j in i.rstrip().split('\t')] for i in table]
    max_len = max([len(i) for i in table])
    if mergeCellsColNum < 0:
        mergeCellsColNum = max_len
    for i in range(1, len(table)):
        for j in range(min(mergeCellsColNum, len(table[i]), len(table[i-1]))):
            if table[i][j] == blankPad and len(table[i-1]) > j:
                table[i][j] = table[i-1][j]
            else:
                break
    lines = ''
    for i in range(len(table)):
        if i == 0:
            for j, v in enumerate(table[i]):
                lines += seg * j + v + '\n'
            continue
        sim_num = 0
        for j in range(min(len(table[i]), len(table[i-1]))):
            if table[i][j] == table[i-1][j]:
                sim_num += 1
            else:
                break
        if sim_num == len(table[i-1]):
            lines += seg * sim_num + blankPad + '\n'
        for j, v in enumerate(table[i][sim_num:]):
            lines += seg * sim_num + seg * j + v + '\n'
    return lines.strip()


if __name__ == '__main__':
    removeMarks = True  # 是否移除加粗/斜体/删除线/高亮格式
    mergeCellsColNum = -1  # 合并单元格涉及的前几列, 小于0表示全部考虑

    text = pyperclip.paste().strip()
    print('-'*10+'转换前'+'-'*10)
    print(text)
    if '\t' in text:
        print('-' * 10 + '转换后(可以复制到roam)' + '-' * 10)
        print(excelToRoam(text, mergeCellsColNum))
    else:
        print('-' * 10 + '转换后(可以复制到excel)' + '-' * 10)
        print(roamToExcel(text, removeMarks))
