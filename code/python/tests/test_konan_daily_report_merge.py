# -*- coding: utf-8 -*-
from __future__ import annotations

import csv
from pathlib import Path

import pandas as pd
import pytest

from planning_core.konan_daily_report import (
    ENV_DAILY_REPORT_CSV_PATH,
    ENV_DAILY_REPORT_MERGE,
    aggregate_daily_report_latest,
    merge_daily_report_into_plan_df,
    read_konan_daily_report_csv,
)

_HEADER = (
    "倉庫コード,倉庫,機械コード,機械名,工程コード,工程名,加工日付,依頼NO,受注NO,加工実績NO,"
    "加工担当者1コード,加工担当者1名,加工担当者2コード,加工担当者2名,加工担当者3コード,加工担当者3名,"
    "開始時間,終了時間,残業時間_分,残業当者1コード,残業当者1名,残業当者2コード,残業当者2名,"
    "残業当者3コード,残業当者3名,停機時間_分,休憩時間_分,稼働時間_分,所要時間,製品,製品梱包,製品色,製品区分,"
    "投入原反,原反梱包,原反色,原反区分,受注数量,換算数量,加工日加工予定数,実加工量,実製品出来高,原反着予定,"
    "指定納期,加工開始予定日,加工完了予定日,加工内容,商品特記事項,製造条件特記事項,納入先1,納入先2,コア着日,"
    "回答納期,出荷予定日,完了区分,加工完了日,注文単位加工完了区分,注文単位加工完了日,加工実績累計,実製品出来高累計"
)


def _row(**kwargs: str) -> list[str]:
    cols = _HEADER.split(",")
    idx = {name: i for i, name in enumerate(cols)}
    row = [""] * len(cols)
    for key, val in kwargs.items():
        if key in idx:
            row[idx[key]] = val
    return row


def _write_sample_daily_report(path: Path) -> None:
    rows = [
        _row(
            機械名="スライス機1　湖南",
            工程名="スライス",
            加工日付="2026/06/22",
            依頼NO="Y6-19",
            換算数量="10000",
            完了区分="0:未完",
            注文単位加工完了区分="0:未完",
            加工実績累計="2000",
            実製品出来高累計="4000",
            製品="PROD-A",
        ),
        _row(
            機械名="スライス機1　湖南",
            工程名="スライス",
            加工日付="2026/06/23",
            依頼NO="Y6-19",
            換算数量="10000",
            完了区分="0:未完",
            注文単位加工完了区分="0:未完",
            加工実績累計="3800",
            実製品出来高累計="7600",
            製品="PROD-A",
        ),
        _row(
            機械名="スライス機1　湖南",
            工程名="スライス",
            加工日付="2026/06/23",
            依頼NO="Y6-18",
            換算数量="6000",
            完了区分="1:完了",
            加工完了日="2026/06/23",
            注文単位加工完了区分="1:完了",
            注文単位加工完了日="2026/06/23",
            加工実績累計="6000",
            実製品出来高累計="12000",
            製品="PROD-B",
        ),
        _row(
            機械名="巻返し機",
            工程名="巻返し",
            加工日付="2026/06/23",
            依頼NO="ONLY-DR",
            換算数量="500",
            完了区分="0:未完",
            注文単位加工完了区分="0:未完",
            加工実績累計="200",
            実製品出来高累計="200",
            製品="PROD-C",
        ),
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as fp:
        fp.write("倉庫     : 520201 湖南工場01本倉庫\n")
        fp.write("加工日付 : 2026年06月23日 ～ 2026年06月23日\n")
        fp.write("機械     : 2011           ～ 2901\n")
        writer = csv.writer(fp)
        writer.writerow(_HEADER.split(","))
        writer.writerows(rows)


@pytest.fixture()
def sample_daily_report_csv(tmp_path: Path) -> Path:
    p = tmp_path / "加工日報発行問合せ_test.csv"
    _write_sample_daily_report(p)
    return p


def test_read_and_aggregate_latest(sample_daily_report_csv: Path) -> None:
    df = read_konan_daily_report_csv(str(sample_daily_report_csv))
    assert len(df) == 4
    latest = aggregate_daily_report_latest(df)
    assert len(latest) == 3
    y619 = latest[latest["依頼NO"] == "Y6-19"].iloc[0]
    assert y619["加工日付"] == "2026/06/23"
    assert y619["加工実績累計"] == "3800"
    assert y619["完了区分"] == "0:未完"


def test_merge_updates_plan_and_appends_missing(
    sample_daily_report_csv: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv(ENV_DAILY_REPORT_CSV_PATH, str(sample_daily_report_csv))
    monkeypatch.setenv(ENV_DAILY_REPORT_MERGE, "1")

    plan = pd.DataFrame(
        [
            {
                "依頼NO": "Y6-19",
                "工程名": "スライス",
                "機械名": "スライス機1　湖南",
                "換算数量": "10000",
                "未加工": "10000",
                "実加工数": "0",
                "加工完了区分": "",
            },
            {
                "依頼NO": "Y6-18",
                "工程名": "スライス",
                "機械名": "スライス機1　湖南",
                "換算数量": "6000",
                "未加工": "6000",
                "実加工数": "0",
                "加工完了区分": "",
            },
        ]
    )

    merged = merge_daily_report_into_plan_df(plan)
    assert len(merged) == 3

    y619 = merged[merged["依頼NO"] == "Y6-19"].iloc[0]
    assert y619["実加工数"] == "3800"
    assert y619["未加工"] == "6200"
    assert y619["加工完了区分"] == "0:未完"

    y618 = merged[merged["依頼NO"] == "Y6-18"].iloc[0]
    assert y618["実加工数"] == "6000"
    assert y618["未加工"] == "0"
    assert y618["加工完了区分"] == "1:完了"

    only = merged[merged["依頼NO"] == "ONLY-DR"]
    assert len(only) == 1
    assert only.iloc[0]["工程名"] == "巻返し"
    assert only.iloc[0]["実加工数"] == "200"
    assert only.iloc[0]["未加工"] == "300"
    assert only.iloc[0]["加工完了区分"] == "0:未完"


def test_merge_disabled(sample_daily_report_csv: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(ENV_DAILY_REPORT_CSV_PATH, str(sample_daily_report_csv))
    monkeypatch.setenv(ENV_DAILY_REPORT_MERGE, "0")
    plan = pd.DataFrame([{"依頼NO": "Y6-19", "工程名": "スライス", "機械名": "x"}])
    merged = merge_daily_report_into_plan_df(plan)
    assert len(merged) == 1
