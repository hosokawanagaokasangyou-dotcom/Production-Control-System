from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

from common import resolve_base
import common
from read_kokubu import read_kokubu_month
from read_konan import read_konan_month
from scan import list_kokubu_files, list_konan_files, pick_period, resolve_end_ym
from write_excel import archive_old_trends, write_trend_book


def _ym_label(ym: tuple[int, int]) -> str:
    return f'{ym[0]}年{ym[1]}月'


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(description='後加工工賃 工程別・シート区分別 月次トレンド')
    p.add_argument('--months', type=int, default=6, help='直近月数 (既定 6)')
    p.add_argument('--factory', choices=('kokubu', 'konan', 'both'), default='both')
    p.add_argument('--out', type=str, default=None)
    p.add_argument('--base', type=str, default=None, help='●自動検証ルート')
    args = p.parse_args(argv)

    base = Path(args.base) if args.base else resolve_base(__file__)
    warnings: list[str] = []
    kokubu_avail: dict = {}
    konan_avail: dict = {}

    if args.factory in ('kokubu', 'both'):
        kdirs = [Path(p) for p in common.KOKUBU_MEISAI_DIRS]
        missing = [str(d) for d in kdirs if not d.is_dir()]
        for m in missing:
            warnings.append(f'国分フォルダにアクセスできません: {m}')
        if all(d.is_dir() for d in kdirs) or any(d.is_dir() for d in kdirs):
            kokubu_avail = list_kokubu_files(kdirs)
        if not kokubu_avail and args.factory == 'kokubu':
            print('国分の後加工工賃明細が見つかりません', file=sys.stderr)
            return 1

    if args.factory in ('konan', 'both'):
        root = Path(common.KONAN_SHISAN_ROOT)
        if not root.is_dir():
            msg = f'湖南試算ルートにアクセスできません: {root}'
            if args.factory == 'konan':
                print(msg, file=sys.stderr)
                return 1
            warnings.append(msg)
        else:
            konan_avail = list_konan_files(root)

    end = resolve_end_ym(
        kokubu_avail if args.factory in ('kokubu', 'both') else {},
        konan_avail if args.factory in ('konan', 'both') else {},
    )
    if end is None:
        print('対象工場で読める月次ファイルがありません', file=sys.stderr)
        return 1

    months_k, files_k = pick_period(kokubu_avail, args.months, end=end)
    months_n, files_n = pick_period(konan_avail, args.months, end=end)
    months = months_k  # same end/n → same month list

    kokubu_data = []
    konan_data = []
    for ym in months:
        if ym in files_k:
            agg = read_kokubu_month(files_k[ym])
            warnings.extend(agg.get('warnings') or [])
            kokubu_data.append(agg)
        elif args.factory in ('kokubu', 'both'):
            warnings.append(f'国分: {_ym_label(ym)} の明細がありません')
        if ym in files_n:
            agg = read_konan_month(files_n[ym])
            warnings.extend(agg.get('warnings') or [])
            konan_data.append(agg)
        elif args.factory in ('konan', 'both') and konan_avail is not None:
            if Path(common.KONAN_SHISAN_ROOT).is_dir() or konan_avail:
                warnings.append(f'湖南: {_ym_label(ym)} の試算がありません')

    if not kokubu_data and not konan_data:
        print('対象期間に集計できるデータがありません', file=sys.stderr)
        return 1

    # 単位変換は write_trend_book 側（量: m→km、賃: 円→千円）

    if args.out:
        out = Path(args.out)
    else:
        stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        out = base / f'月トレンド_{stamp}.xlsx'

    payload = {
        'months': months,
        'kokubu': kokubu_data,
        'konan': konan_data,
        'warnings': warnings,
        'files': {
            '国分': files_k,
            '湖南': files_n,
        },
    }
    write_trend_book(out, payload)
    if out.parent.resolve() == base.resolve() and out.name.startswith('月トレンド_'):
        archive_old_trends(base, keep=out)

    print(f'対象期間: {_ym_label(months[0])} 〜 {_ym_label(months[-1])}（{len(months)}か月）')
    print(f'国分ファイル: {len(files_k)} / 湖南ファイル: {len(files_n)}')
    for w in warnings:
        print(f'【警告】{w}')
    print(f'出力: {out}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
