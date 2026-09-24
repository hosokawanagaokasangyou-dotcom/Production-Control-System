from __future__ import annotations

import re
import unicodedata


def _index_by_ym(months_data: list[dict]) -> dict[tuple[int, int], dict]:
    return {m['ym']: m for m in months_data if m.get('ym')}


def _norm_proc(name: str) -> str:
    s = unicodedata.normalize('NFKC', str(name or '')).replace('\n', ' ').strip()
    s = re.sub(r'\s+', ' ', s)
    return s


# グループ名 → エイリアス（正規化後に前方一致/完全一致で判定）
# 国分「スライス1/3」と湖南「スライス」を同一視する
_GROUP_RULES: list[tuple[str, tuple[str, ...]]] = [
    ('スライス', ('スライス1', 'スライス3', 'スライス１', 'スライス３', 'スライス')),
    ('スリット', ('スリット4 ゴールド', 'スリット4ゴールド', 'スリット')),
    ('LAC(EC)', ('LAC(EC)', 'ＬＡＣ(EC)', 'LAC（EC）')),
    ('SEC(SS)', ('SEC(SS)', 'SEC (SS)', 'SEC（SS）')),
    ('SEC(EC)', ('SEC(EC)', 'SEC (EC)', 'SEC（EC）')),
]


def process_group(name: str) -> str:
    """生工程名 → 比較用グループ名。未定義は正規化名のまま。"""
    n = _norm_proc(name)
    if not n:
        return n
    for group, aliases in _GROUP_RULES:
        for a in aliases:
            if n == a or n.startswith(a):
                # 「スライス」は「スライス1」等に先にマッチさせるため aliases 順が重要
                if group == 'スライス' and n.startswith('スライス'):
                    return 'スライス'
                if group == 'スリット' and n.startswith('スリット'):
                    return 'スリット'
                if n == a:
                    return group
    for group, aliases in _GROUP_RULES:
        for a in aliases:
            if n == a:
                return group
    if n.startswith('スライス'):
        return 'スライス'
    if n.startswith('スリット'):
        return 'スリット'
    return n


def build_cross_kind(
    kokubu: list[dict],
    konan: list[dict],
    months: list[tuple[int, int]],
    metric: str,
) -> dict:
    """横断: processes ソート済み, values[工場][工程][ym] = float|None（生名）"""
    kb = _index_by_ym(kokubu)
    kn = _index_by_ym(konan)
    processes = set()
    for src in (kb, kn):
        for ym in months:
            kinds = (src.get(ym) or {}).get('kinds') or {}
            processes.update(kinds)
    processes_l = sorted(processes)
    values = {'国分': {}, '湖南': {}}
    for factory, src in (('国分', kb), ('湖南', kn)):
        for proc in processes_l:
            values[factory][proc] = {}
            for ym in months:
                kinds = (src.get(ym) or {}).get('kinds') or {}
                cell = kinds.get(proc)
                if cell is None:
                    values[factory][proc][ym] = None
                else:
                    values[factory][proc][ym] = float(cell.get(metric, 0) or 0)
    return {'processes': processes_l, 'months': months, 'metric': metric, 'values': values}


def build_grouped_cross(
    kokubu: list[dict],
    konan: list[dict],
    months: list[tuple[int, int]],
    metric: str,
) -> dict:
    """工程グループで合算した横断。values[工場][グループ][ym]=float|None"""
    kb = _index_by_ym(kokubu)
    kn = _index_by_ym(konan)
    groups: set[str] = set()
    raw: dict[str, dict[str, dict]] = {'国分': {}, '湖南': {}}

    for factory, src in (('国分', kb), ('湖南', kn)):
        for ym in months:
            kinds = (src.get(ym) or {}).get('kinds') or {}
            bucket: dict[str, float] = {}
            for name, cell in kinds.items():
                g = process_group(name)
                if not g:
                    continue
                groups.add(g)
                bucket[g] = bucket.get(g, 0.0) + float((cell or {}).get(metric, 0) or 0)
            for g, v in bucket.items():
                raw[factory].setdefault(g, {})[ym] = v

    groups_l = sorted(groups)
    values = {'国分': {}, '湖南': {}}
    for factory in ('国分', '湖南'):
        for g in groups_l:
            values[factory][g] = {}
            for ym in months:
                # その月に工場データ自体が無ければ None（ゼロ埋めしない）
                src = kb if factory == '国分' else kn
                if ym not in src:
                    values[factory][g][ym] = None
                else:
                    values[factory][g][ym] = raw[factory].get(g, {}).get(ym, 0.0)
    return {
        'processes': groups_l,
        'months': months,
        'metric': metric,
        'values': values,
        'grouped': True,
    }


def build_sheet_series(
    months_data: list[dict],
    months: list[tuple[int, int]],
    metric: str,
) -> dict:
    by_ym = _index_by_ym(months_data)
    names = set()
    for ym in months:
        sheets = (by_ym.get(ym) or {}).get('sheets') or {}
        names.update(sheets)
    names_l = sorted(names)
    values = {}
    for name in names_l:
        values[name] = {}
        for ym in months:
            sheets = (by_ym.get(ym) or {}).get('sheets') or {}
            cell = sheets.get(name)
            values[name][ym] = None if cell is None else float(cell.get(metric, 0) or 0)
    return {'sheet_names': names_l, 'months': months, 'metric': metric, 'values': values}


def build_kind_series(
    months_data: list[dict],
    months: list[tuple[int, int]],
    metric: str,
) -> dict:
    by_ym = _index_by_ym(months_data)
    names = set()
    for ym in months:
        names.update((by_ym.get(ym) or {}).get('kinds') or {})
    names_l = sorted(names)
    values = {}
    for name in names_l:
        values[name] = {}
        for ym in months:
            kinds = (by_ym.get(ym) or {}).get('kinds') or {}
            cell = kinds.get(name)
            values[name][ym] = None if cell is None else float(cell.get(metric, 0) or 0)
    return {'processes': names_l, 'months': months, 'metric': metric, 'values': values}


def mom_share(
    kinds_by_ym: dict[tuple[int, int], dict],
    months: list[tuple[int, int]],
    metric: str,
) -> dict:
    if not months:
        return {'mom': {}, 'share': {}, 'latest': None, 'prev': None}
    latest = months[-1]
    prev = months[-2] if len(months) >= 2 else None
    latest_kinds = kinds_by_ym.get(latest) or {}
    prev_kinds = kinds_by_ym.get(prev) or {} if prev else {}

    total = sum(float((v or {}).get(metric, 0) or 0) for v in latest_kinds.values())
    mom = {}
    share = {}
    procs = set(latest_kinds) | set(prev_kinds)
    for p in sorted(procs):
        cur = float((latest_kinds.get(p) or {}).get(metric, 0) or 0)
        share[p] = (cur / total) if total else None
        if prev is None:
            mom[p] = None
            continue
        before = float((prev_kinds.get(p) or {}).get(metric, 0) or 0)
        if before == 0:
            mom[p] = None
        else:
            mom[p] = (cur - before) / before
    return {'mom': mom, 'share': share, 'latest': latest, 'prev': prev}


def shift_badge(d_k, d_h) -> str | None:
    """前月差から振替疑いバッジ。断定しない表示用。"""
    if d_k is None or d_h is None:
        return None
    if d_k == 0 and d_h == 0:
        return '変化なし'
    # 異符号 = 相殺パターン
    if d_k * d_h < 0:
        ratio = min(abs(d_k), abs(d_h)) / max(abs(d_k), abs(d_h))
        if ratio >= 0.5:
            return '強い相殺'
        if ratio >= 0.2:
            return '弱い相殺'
        return '弱い相殺'
    return '同方向'


def _pair_months(months: list[tuple[int, int]], values: dict) -> tuple | None:
    """両工場に値がある直近2か月を返す (prev, latest)"""
    usable = []
    for ym in months:
        k = values['国分']
        h = values['湖南']
        # いずれかの工程に数値があればその月は候補 — 呼び出し側で工程単位に見る
        usable.append(ym)
    if len(usable) < 2:
        return None
    return usable[-2], usable[-1]


def rank_shift_suspects(grouped: dict, metric_label: str = 'wage', top_n: int = 20) -> list[dict]:
    """グループ横断から、直近2か月の前月差で振替疑いを順位付け。"""
    months = grouped['months']
    values = grouped['values']
    if len(months) < 2:
        return []
    prev, latest = months[-2], months[-1]
    rows = []
    for proc in grouped['processes']:
        k0 = values['国分'].get(proc, {}).get(prev)
        k1 = values['国分'].get(proc, {}).get(latest)
        h0 = values['湖南'].get(proc, {}).get(prev)
        h1 = values['湖南'].get(proc, {}).get(latest)
        if k0 is None or k1 is None or h0 is None or h1 is None:
            # 片方欠損月は比較しない（偽の急減を避ける）
            continue
        dk, dh = k1 - k0, h1 - h0
        badge = shift_badge(dk, dh)
        offset = min(abs(dk), abs(dh)) if dk * dh < 0 else 0.0
        combined = (k1 + h1) - (k0 + h0)
        rows.append({
            'process': proc,
            'prev': prev,
            'latest': latest,
            'kokubu_prev': k0,
            'kokubu_latest': k1,
            'konan_prev': h0,
            'konan_latest': h1,
            'd_kokubu': dk,
            'd_konan': dh,
            'd_combined': combined,
            'offset': offset,
            'badge': badge,
            'metric': metric_label,
        })
    # 相殺を優先、次いで |dk|+|dh|
    def sort_key(r):
        pri = 0 if r['badge'] in ('強い相殺', '弱い相殺') else 1
        return (pri, -r['offset'], -(abs(r['d_kokubu']) + abs(r['d_konan'])))

    rows.sort(key=sort_key)
    return rows[:top_n]


def pick_focus_processes(grouped_wage: dict, grouped_qty: dict, max_n: int = 6) -> list[str]:
    """注目工程: スライス固定 + 量の振替疑い上位 + 賃の変動上位。"""
    focus: list[str] = []
    if 'スライス' in grouped_wage.get('processes', []):
        focus.append('スライス')
    for src in (grouped_qty, grouped_wage):
        for row in rank_shift_suspects(src, top_n=max_n):
            p = row['process']
            if p not in focus:
                focus.append(p)
            if len(focus) >= max_n:
                return focus
    # まだ足りなければ金額規模上位
    months = grouped_wage['months']
    latest = months[-1] if months else None
    if latest:
        sized = []
        for p in grouped_wage['processes']:
            k = grouped_wage['values']['国分'].get(p, {}).get(latest) or 0
            h = grouped_wage['values']['湖南'].get(p, {}).get(latest) or 0
            sized.append((k + h, p))
        for _, p in sorted(sized, reverse=True):
            if p not in focus:
                focus.append(p)
            if len(focus) >= max_n:
                break
    return focus[:max_n]


def factory_month_total(months_data: list[dict], ym: tuple[int, int], metric: str) -> float | None:
    by = _index_by_ym(months_data)
    m = by.get(ym)
    if not m:
        return None
    kinds = m.get('kinds') or {}
    return sum(float((v or {}).get(metric, 0) or 0) for v in kinds.values())
