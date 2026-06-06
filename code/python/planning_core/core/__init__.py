# -*- coding: utf-8 -*-
"""
planning_core.core — _core.py から機械分割した内部サブパッケージ。

外部からは ``planning_core._core``（ファサード）経由で import すること。
実体は ``_core.py`` が ``core/_bootstrap.py`` + 各 ``core/*.py`` を exec 連結した共有名前空間。
再生成: ``code/python/scripts/split_core_modules.py``（正本 ``_core.py.bak``）。
"""
