# -*- coding: utf-8 -*-
import os
import tempfile

from execution_log_io import EXECUTION_LOG_MAX_LINES, trim_execution_log_if_oversized


def test_trim_execution_log_if_oversized_keeps_tail():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "execution_log.txt")
        line = b"2026-01-01 00:00:00 - INFO - hello\n"
        with open(path, "wb") as f:
            f.write(b"\xef\xbb\xbf")
            f.write(line * (EXECUTION_LOG_MAX_LINES + 500))

        assert trim_execution_log_if_oversized(path)
        with open(path, "rb") as f:
            data = f.read()
        assert data.startswith(b"\xef\xbb\xbf")
        assert data.count(b"\n") == EXECUTION_LOG_MAX_LINES
        assert data.endswith(b"hello\n")
        assert data.find(b"hello\n") >= 0


def test_trim_execution_log_if_oversized_noop_when_small():
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "execution_log.txt")
        with open(path, "wb") as f:
            f.write(b"\xef\xbb\xbfok\n")
        assert not trim_execution_log_if_oversized(path)
        assert os.path.getsize(path) == 6
