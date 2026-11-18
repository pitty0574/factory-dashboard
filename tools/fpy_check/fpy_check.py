# -*- coding: utf-8 -*-
"""
一次性調查：直行率負值原因（只讀 Google Sheets，不改任何資料）
讀取方式與 person_part_analysis.html 相同：同一個 Sheet ID / API Key、
補漆＝「整修日報」A:G、終檢＝「終檢階段」A:F、員工表排除「離職」。

用法：python tools/fpy_check/fpy_check.py [起日 YYYY/MM/DD] [迄日 YYYY/MM/DD]
輸出：markdown 印到 stdout，並寫入 tools/fpy_check/report.md
"""
import json
import os
import re
import sys
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime

SHEET_ID = '1xQoIxoLlM4LkKjDrYZ_87uIahtebPgPcG6Hitck1T5c'   # person_part_analysis.html DEFAULT_CONFIG
API_KEY = 'AIzaSyBMgXHsua64mciOFJwqH4NreGI11gMXR-s'
PAINT_RANGE = '整修日報!A:G'      # A 時間戳記, B 記錄日期, C 工號, D 車型, E 種類, F 顏色代號, G 補漆總點數
FINAL_RANGE = '終檢階段!A:F'      # A 時間戳記, B 記錄日期, C 工號, D 車型, E 部件, F 顏色
WORKER_RANGE = '員工表!A2:D1000'  # A 工號, B 姓名, C 崗位, D 狀態
RVPU = ('RV', 'PU')
FOCUS_BL = ('AQ', 'RV', 'PU')
DUP_WINDOW_SEC = 120

START = sys.argv[1] if len(sys.argv) > 1 else '2026/11/01'
END = sys.argv[2] if len(sys.argv) > 2 else '2026/11/17'


def fetch(rng):
    url = (f'https://sheets.googleapis.com/v4/spreadsheets/{SHEET_ID}/values/'
           f'{urllib.parse.quote(rng)}?key={API_KEY}')
    req = urllib.request.Request(url, headers={'Referer': 'http://localhost/'})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.load(r).get('values', [])


def cell(row, i):
    return str(row[i]).strip() if i < len(row) and row[i] is not None else ''


# 與看板 normalizeDate 相同的規則 → 'YYYY/MM/DD' 或 ''
def normalize_date(s):
    s = re.sub(r'\s*(上午|下午|AM|PM)[\s\d:：]*$', '', s.strip(), flags=re.I).strip()
    s = s.split('T')[0].strip().replace('-', '/')
    if re.fullmatch(r'\d{4}/\d{2}/\d{2}', s):
        return s
    m = re.fullmatch(r'(\d{4})/(\d{1,2})/(\d{1,2})', s)
    if m:
        return f'{m[1]}/{int(m[2]):02d}/{int(m[3]):02d}'
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return f'{m[3]}/{int(m[1]):02d}/{int(m[2]):02d}'
    return ''


# 時間戳記「2026/4/7 上午9:22:42」「2026/10/10 下午08:08:58」→ datetime
def parse_ts(raw):
    m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})[\sT ]*(上午|下午|AM|PM)?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?',
                  raw, re.I)
    if not m:
        return None
    h = int(m[5])
    ap = (m[4] or m[8] or '').upper()
    if ap in ('下午', 'PM') and h < 12:
        h += 12
    if ap in ('上午', 'AM') and h == 12:
        h = 0
    return datetime(int(m[1]), int(m[2]), int(m[3]), h, int(m[6]), int(m[7] or 0))


# 與 JS parseFloat 相同：取開頭的數字，取不到 → None
def js_parse_float(s):
    m = re.match(r'\s*([+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?)', s)
    return float(m[1]) if m else None


def load():
    workers = {}
    for r in fetch(WORKER_RANGE):
        wid = cell(r, 0)
        if wid:
            workers[wid] = cell(r, 3) != '離職'
    active = lambda pid: workers.get(pid, True)   # 查無員工 → 視為在職（同看板）

    in_range = lambda d: START <= d <= END
    paint = []
    for r in fetch(PAINT_RANGE)[1:]:
        date = normalize_date(cell(r, 1))
        if not date or not in_range(date) or not active(cell(r, 2)):
            continue
        paint.append({
            'date': date, 'pid': cell(r, 2), 'model': cell(r, 3).upper(), 'part': cell(r, 4),
            'color': cell(r, 5).upper(), 'pts_raw': cell(r, 6), 'pts': js_parse_float(cell(r, 6)),
            'ts': parse_ts(cell(r, 0)), 'ts_raw': cell(r, 0),
        })
    final = []
    for r in fetch(FINAL_RANGE)[1:]:
        date = normalize_date(cell(r, 1))
        pid = cell(r, 2)
        if not date or not pid or not in_range(date) or not active(pid):
            continue
        final.append({'date': date, 'pid': pid, 'model': cell(r, 3).upper(), 'part': cell(r, 4), 'color': cell(r, 5).upper()})
    return paint, final


# 與看板 calcPaintNet 相同：排除點數 ≤ 0（空白／非數字算有補漆），工號＋車型＋部件分組，
# 與同組前一筆間隔 ≤ PAINT_MERGE_SEC 視為同一件（連鎖）；無時間戳記各算一件
PAINT_MERGE_SEC = 120


def calc_paint_net(rows):
    groups = defaultdict(list)
    net = 0
    for p in rows:
        if p['pts'] is not None and p['pts'] <= 0:
            continue
        if not p['ts']:
            net += 1
            continue
        groups[(p['pid'].upper(), p['model'].upper(), p['part'].upper())].append(p['ts'])
    for ts in groups.values():
        ts.sort()
        net += sum(1 for i, t in enumerate(ts) if i == 0 or (t - ts[i - 1]).total_seconds() > PAINT_MERGE_SEC)
    return net


# 各車型（依終檢件數排序）：final / paint / net / over（補漆淨 > 終檢）
def new_formula(paint, final):
    models = sorted({p['model'] for p in paint} | {f['model'] for f in final})
    out = []
    for m in models:
        fc = sum(1 for f in final if f['model'] == m)
        pr = [p for p in paint if p['model'] == m]
        net = calc_paint_net(pr)
        out.append({'model': m, 'final': fc, 'paint': len(pr), 'net': net, 'over': net > fc})
    return sorted(out, key=lambda d: (-d['final'], -d['paint'], d['model']))


def pct(num, den):
    return '—' if not den else f'{num / den * 100:.1f}%'


def md_table(head, rows):
    out = ['| ' + ' | '.join(head) + ' |', '|' + '|'.join('---' for _ in head) + '|']
    out += ['| ' + ' | '.join(str(c) for c in r) + ' |' for r in rows]
    return '\n'.join(out)


def main():
    paint, final = load()
    out = [f'# 直行率負值調查 {START} ～ {END}', '',
           f'補漆（整修日報，排除離職）{len(paint)} 筆；終檢（終檢階段，排除離職）{len(final)} 件', '']
    pts_kind = lambda p: 'zero' if p['pts'] == 0 else 'pos' if (p['pts'] or 0) > 0 else 'na'

    # ── 1. 車型 × 顏色 ──
    mc = defaultdict(lambda: {'n': 0, 'zero': 0, 'pos': 0, 'na': 0})
    for p in paint:
        d = mc[(p['model'], p['color'])]
        d['n'] += 1
        d[pts_kind(p)] += 1
    rows = [[m or '（未填）', c or '（未填）', d['n'], d['zero'], d['pos'], d['na']]
            for (m, c), d in sorted(mc.items(), key=lambda kv: (kv[0][0], -kv[1]['n']))]
    out += ['## 1. 補漆：車型 × 顏色', '',
            md_table(['車型', '顏色', '補漆筆數', '點數=0', '點數>0', '點數空白/非數字'], rows), '',
            '### 1a. AQ / RV / PU 的 BL', '']
    out.append(md_table(['車型', '顏色', '補漆筆數', '點數=0', '點數>0', '點數空白/非數字'],
                        [[m, 'BL'] + [mc[(m, 'BL')][k] for k in ('n', 'zero', 'pos', 'na')] for m in FOCUS_BL]))
    out.append('')

    # ── 2. 各車型：現行公式 vs 只算點數 > 0 ──
    fin_by_model = defaultdict(int)
    for f in final:
        fin_by_model[f['model']] += 1
    pb = defaultdict(lambda: {'n': 0, 'pos': 0})
    for p in paint:
        pb[p['model']]['n'] += 1
        pb[p['model']]['pos'] += pts_kind(p) == 'pos'
    rows = []
    summary = []
    for m in sorted(set(fin_by_model) | set(pb), key=lambda k: (-fin_by_model[k], -pb[k]['n'], k)):
        fc, pc, pos = fin_by_model[m], pb[m]['n'], pb[m]['pos']
        bl = min(pc, fc) if m in RVPU else 0          # 現行：整段期間 min()（renderModelTable）
        net_now, net_new = pc - bl, pos
        f_now = (fc - net_now) / fc if fc else None
        f_new = (fc - net_new) / fc if fc else None
        rows.append([m or '（未填）', fc, pc, f'{net_now}' + (f'（扣 {bl}）' if bl else ''), pct(fc - net_now, fc),
                     pos, pct(fc - net_new, fc)])
        summary.append((m, fc, f_now, f_new))
    out += ['## 2. 各車型直行率：現行公式 vs 只算點數 > 0', '',
            '- 現行＝看板「各車型直行率明細」：補漆淨 = 補漆筆數 − (RV/PU 時 min(補漆, 終檢)，整段期間算一次)',
            '- 新＝補漆淨 = 點數 > 0 的筆數（不再另扣 min()）', '',
            md_table(['車型', '終檢', '補漆總筆數', '現行補漆淨', '現行直行率', '點數>0 補漆淨', '新直行率'], rows), '']

    # ── 3. RV 逐日 ──
    days = sorted({p['date'] for p in paint if p['model'] == 'RV'} | {f['date'] for f in final if f['model'] == 'RV'})
    rows = []
    sum_daily = tot_f = tot_p = 0
    for d in days:
        fr = sum(1 for f in final if f['date'] == d and f['model'] == 'RV')
        pr = sum(1 for p in paint if p['date'] == d and p['model'] == 'RV')
        bl = sum(1 for p in paint if p['date'] == d and p['model'] == 'RV' and p['color'] == 'BL')
        mn = min(pr, fr)
        sum_daily += mn
        tot_f += fr
        tot_p += pr
        flag = '⚠ 補漆有 RV、終檢沒 RV' if pr and not fr else ('補漆 > 終檢' if pr > fr else '')
        rows.append([d, fr, pr, bl, mn, flag])
    rows.append(['**合計**', tot_f, tot_p, sum(r[3] for r in rows), f'逐日加總 {sum_daily}／整段 {min(tot_p, tot_f)}', ''])
    out += ['## 3. RV 逐日', '',
            '- 「各車型直行率明細」表格：min() 整段期間算一次（renderModelTable）',
            '- 「直行率推移」圖：min() 逐日算，且全部車型時是 RV+PU 合併（computeDailyFPY）', '',
            md_table(['日期', '終檢 RV', '補漆 RV', '其中 BL', '當天 min()', '備註'], rows), '']

    # ── 4. 一件多筆：同工號、同車型、同部件，時間差 ≤ ±2 分鐘 ──
    groups = defaultdict(list)
    no_ts = 0
    for p in paint:
        if not p['ts']:
            no_ts += 1
            continue
        groups[(p['pid'], p['model'], p['part'])].append(p)
    clusters = []
    for key, lst in groups.items():
        lst.sort(key=lambda p: p['ts'])
        cur = [lst[0]]
        for p in lst[1:]:
            if (p['ts'] - cur[-1]['ts']).total_seconds() <= DUP_WINDOW_SEC:
                cur.append(p)
            else:
                if len(cur) > 1:
                    clusters.append((key, cur))
                cur = [p]
        if len(cur) > 1:
            clusters.append((key, cur))
    by_model = defaultdict(lambda: {'c': 0, 'rows': 0, 'extra': 0, 'same_color': 0, 'extra_pos': 0})
    for (pid, m, part), cl in clusters:
        d = by_model[m]
        d['c'] += 1
        d['rows'] += len(cl)
        d['extra'] += len(cl) - 1
        d['same_color'] += len({p['color'] for p in cl}) == 1
        d['extra_pos'] += max(0, sum(1 for p in cl if pts_kind(p) == 'pos') - 1)
    rows = [[m or '（未填）', d['c'], d['rows'], d['extra'], d['same_color'], d['extra_pos'], pb[m]['n'], pct(d['extra'], pb[m]['n'])]
            for m, d in sorted(by_model.items(), key=lambda kv: -kv[1]['extra'])]
    tot_extra = sum(d['extra'] for d in by_model.values())
    rows.append(['**合計**', sum(d['c'] for d in by_model.values()), sum(d['rows'] for d in by_model.values()), tot_extra,
                 sum(d['same_color'] for d in by_model.values()), sum(d['extra_pos'] for d in by_model.values()), len(paint), pct(tot_extra, len(paint))])
    out += ['## 4. 一件多筆（同工號＋同車型＋同部件，相鄰兩筆 ≤ 2 分鐘串成一群）', '',
            f'時間戳記無法解析而略過：{no_ts} 筆', '',
            md_table(['車型', '群數', '群內筆數', '多出筆數(群內−1)', '同顏色群數', '多出且點數>0', '該車型補漆筆數', '多出占比'], rows), '',
            '### 4a. 範例（前 10 群）', '']
    ex = []
    for (pid, m, part), cl in sorted(clusters, key=lambda kv: -len(kv[1]))[:10]:
        ex.append([pid, m, part, len(cl), cl[0]['ts'].strftime('%m/%d %H:%M:%S') + ' ～ ' + cl[-1]['ts'].strftime('%H:%M:%S'),
                   '、'.join(p['color'] for p in cl), '、'.join(p['pts_raw'] or '空' for p in cl)])
    out += [md_table(['工號', '車型', '部件', '筆數', '時間', '顏色', '點數'], ex), '']

    # ── 5. 新公式（與看板 calcPaintNet／calcFPY 同一套規則）──
    stats = new_formula(paint, final)
    valid = [d for d in stats if d['final'] > 0 and not d['over']]
    tf, tn = sum(d['final'] for d in valid), sum(d['net'] for d in valid)
    rows = [[d['model'] or '（未填）', d['final'], d['paint'], d['net'],
             '資料不足' if d['over'] and d['final'] else '—' if not d['final'] else pct(d['final'] - d['net'], d['final'])]
            for d in stats]
    rows.append(['**合計（有效車型）**', tf, sum(d['paint'] for d in valid), tn, pct(tf - tn, tf)])
    over = [d['model'] for d in stats if d['over'] and d['final']]
    nofinal = [d['model'] or '（未填）' for d in stats if d['over'] and not d['final']]
    out += ['## 5. 新公式：補漆淨 = 點數 > 0，同工號同車型同部件 2 分鐘內合併（calcPaintNet）', '',
            '- 補漆淨 > 終檢 → 資料不足，不列入合計',
            f'- 已排除資料不足車型：{"、".join(over) or "無"}；無終檢資料（補漆未計入）：{"、".join(nofinal) or "無"}', '',
            md_table(['車型', '終檢', '補漆登錄', '補漆淨', '直行率'], rows), '']

    # 供總結用的機器可讀摘要
    out += ['## 摘要資料', '', md_table(['車型', '終檢', '現行直行率', '新直行率'],
                                       [[m or '（未填）', fc, '—' if a is None else f'{a*100:.1f}%', '—' if b is None else f'{b*100:.1f}%']
                                        for m, fc, a, b in summary if fc]), '']
    text = '\n'.join(out)
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'report.md'), 'w', encoding='utf-8') as f:
        f.write(text)
    sys.stdout.reconfigure(encoding='utf-8')
    print(text)


if __name__ == '__main__':
    main()
