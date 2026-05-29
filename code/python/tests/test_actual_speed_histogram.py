# -*- coding: utf-8 -*-
from __future__ import annotations

from planning_core.actual_speed_distribution import (
    histogram_from_observations,
    merge_observation_into_histogram,
    percentile,
)


def test_percentile_and_histogram():
    vals = [10.0, 12.0, 14.0, 16.0, 18.0]
    assert percentile(vals, 50) == 14.0
    hist = histogram_from_observations(vals, bin_width=2.0)
    assert len(hist["counts"]) >= 1
    entry: dict = {"histogram": hist}
    merge_observation_into_histogram(entry, 15.0, bin_width=2.0)
    assert sum(entry["histogram"]["counts"]) == len(vals) + 1
