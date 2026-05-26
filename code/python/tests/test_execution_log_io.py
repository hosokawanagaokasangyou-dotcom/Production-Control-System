# -*- coding: utf-8 -*-
import os
import tempfile

from execution_log_io import EXECUTION_LOG_MAX_BYTES, trim_execution_log_if_oversized


def test_trim_execution_log_if_oversized_keeps_tail():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "execution_log.txt")
        line = "2026-01-01 00:00:00 - INFO - hello\n"
        payload = (line * 400_000).encode("utf-8")
        with open(path, "wb") as f:
            f.write(b"\xef\xbb\xbf")
            f.write(payload)
        assert os.path.getsize(path) > EXECUTION_LOG_MAX_BYTES

        assert trim_execution_log_if_oversized(path)
        size = os.path.getsize(path)
        assert size <= EXECUTION_LOG_MAX_BYTES
        with open(path, "rb") as f:
            data = f.read()
        assert data.startswith(b"\xef\xbb\xbf")
        assert b"hello" in data


def test_trim_execution_log_if_oversized_noop_when_small():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "execution_log.txt")
        with open(path, "wb") as f:
            f.write(b"\xef\xbb\xbfok\n")
        assert not trim_execution_log_if_oversized(path)
        assert os.path.getsize(path) == 6
