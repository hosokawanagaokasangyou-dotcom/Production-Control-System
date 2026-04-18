# -*- coding: utf-8 -*-
"""planning_core/_core.py の誤 Unicode（酝坰系・U+37xx 仮名置換）を一括修正する。"""
from __future__ import annotations

import ast
from pathlib import Path


def _eval_assign(src: str, name: str):
    tree = ast.parse(src)
    for node in tree.body:
        if isinstance(node, ast.AnnAssign) and isinstance(node.target, ast.Name) and node.target.id == name:
            return eval(  # noqa: S307
                compile(ast.Expression(node.value), "<bootstrap>", "eval"),
                {"sorted": sorted, "tuple": tuple},
            )
        if isinstance(node, ast.Assign):
            for t in node.targets:
                if isinstance(t, ast.Name) and t.id == name:
                    return eval(  # noqa: S307
                        compile(ast.Expression(node.value), "<bootstrap>", "eval"),
                        {"sorted": sorted, "tuple": tuple},
                    )
    raise KeyError(name)


_SINGLE_NO_QUOTE: tuple[tuple[str, str], ...] = (
    ("㝕", "さ"),
    ("㝮", "の"),
    ("㝯", "は"),
    ("㝫", "に"),
    ("㝧", "で"),
    ("㝨", "と"),
    ("㝌", "は"),
    ("㝓", "こ"),
    ("㝋", "か"),
    ("㝿", "み"),
    ("㝾", "ま"),
    ("㝛", "せ"),
    ("㝙", "れ"),
    ("㝟", "た"),
    ("㝝", "し"),
    ("㝪", "な"),
    ("㝄", "い"),
    ("㝂", "あ"),
    ("㝤", "つ"),
    ("㝸", "へ"),
    ("㝲", "参"),
    ("㝠", "て"),
)

_EXTRA: tuple[tuple[str, str], ...] = (
    ("酝坰㝧加工㝗㝪㝄", "配台で加工している"),
    ("㝗〝", "し、"),
    ("㝗㝾㝗㝟", "しました"),
    ("㝗㝾㝛ん", "しません"),
    ("㝗㝟", "した"),
    ("㝗㝦", "して"),
    ("㝠㝑を", "の値を"),
    ("㝠㝑", "の値"),
    ("load_workbook を試行㝗㝪㝄", "load_workbook を試行する"),
    ("他タスクを試行㝗㝪㝄", "他タスクを試行する"),
    ("試行㝗㝪㝄", "試行する"),
    ("㝗㝪㝄", "しない"),
    ("㝣㝦いる", "っている"),
    ("書かれ㝦いる", "書かれている"),
    ("開い㝦いる", "開いている"),
    ("い㝦いる", "いている"),
    ("ゝ㝦も", "んでも"),
    ("使㝣㝦も", "使っても"),
    ("行㝔とに", "行ごとに"),
    ("相い坈ゝせ㝦してさい", "相談してください"),
    ("割り当㝦ない", "割り当てない"),
    ("当㝦はまる", "当てはまる"),
    ("全㝦の", "全体の"),
    ("適用れる", "適用する"),
    ("坈計", "合計"),
    ("折れ線の坝剝", "折れ線グラフ名"),
    ("書弝", "書式"),
    ("形弝", "形式"),
    ("㝣㝦", "って"),
    ("坈ゝせ", "合わせ"),
    ("編戝", "編集"),
    ("引し継し", "引き継ぎ"),
    ("のとし", "のとき"),
    ("概㝭", "概ね"),
    ("使っても」", "使っても、"),
    ("VBA から渡れ", "VBA から渡される"),
    ("中抜㝑", "中抜き"),
    ("開㝑ません", "開きません"),
    ("開㝑ませ", "開きませ"),
    ("避㝑られ", "避けられ"),
    ("避㝑ら", "避けら"),
    ("切り分㝑", "切り分け"),
    ("文脈切り分㝑", "文脈切り分け"),
    ("分㝑る", "分ける"),
    ("掛㝑", "掛け"),
    ("付㝑", "付け"),
    ("読み分㝑", "読み分け"),
    ("付㝑替", "付け替"),
    ("付㝑直", "付け直"),
    ("空酝列", "空の配列"),
    ("絝㝳付", "結び付"),
)


def main() -> None:
    here = Path(__file__).resolve().parent
    boot = (here / "planning_core" / "bootstrap.py").read_text(encoding="utf-8")
    pairs: tuple[tuple[str, str], ...] = _eval_assign(boot, "_LOG_MOJIBAKE_PAIRS")
    core_path = here / "planning_core" / "_core.py"
    text = core_path.read_text(encoding="utf-8")
    orig = text

    merged = list(pairs) + list(_EXTRA)
    merged.sort(key=lambda x: -len(x[0]))
    for old, new in merged:
        text = text.replace(old, new)

    text = text.replace("\u3757", "し")  # 㝗

    for old, new in _SINGLE_NO_QUOTE:
        text = text.replace(old, new)

    text = text.replace("\u301d", "」")
    text = text.replace("# 」設定】", "# 【設定】")

    # 㝦 の残り（「いで」誤変換を「いて」に戻す）
    for wrong, ok in (
        ("開いでいる", "開いている"),
        ("書いでいる", "書いている"),
        ("入いでいる", "入っている"),
    ):
        text = text.replace(wrong, ok)

    text = text.replace("\u3766", "で")  # 㝦（残りは主に で）

    # 㝑 は多くが「け」だが「開き」系は上で修正済み想定
    text = text.replace("\u3751", "け")  # 㝑

    for wrong, ok in (
        ("開けません", "開きません"),
        ("開けませ", "開きませ"),
    ):
        text = text.replace(wrong, ok)

    # 残存しやすい仮名置換（U+37xx）
    for old, new in (
        ("未保存て㝣た E", "未保存だった E"),
        ("無か㝣た", "なかった"),
        ("でしなか㝣た", "でしなかった"),
        ("しなか㝣た", "しなかった"),
        ("返さなか㝣た", "返さなかった"),
        ("あ㝣た", "あった"),
        ("誤㝣た", "誤った"),
        ("使い切㝣た", "使い切った"),
        ("入㝣た", "入った"),
        ("暗号化形弝て㝣たか", "暗号化形式であったか"),
    ):
        text = text.replace(old, new)

    text = text.replace("\u3763", "っ")  # 㝣（小書きつ）

    for old, new in (
        ("\u3769", "ど"),  # 㝩
        ("\u375a", "う"),  # 㝚
        ("\u3746", "ご"),  # 㝆
        ("\u3748", "ご"),  # 㝈
        ("\u3770", "み"),  # 㝰
        ("\u3754", "と"),  # 㝔
        ("\u374a", "よ"),  # 㝊
        ("\u3761", "う"),  # 㝡
        ("\u3773", "よ"),  # 㝳
        ("\u3758", "も"),  # 㝘
        ("\u377b", "ろ"),  # 㝻
        ("\u3752", "き"),  # 㝒
        ("\u3779", "き"),  # 㝹
        ("\u374e", "ね"),  # 㝎
        ("\u3776", "め"),  # 㝶
        ("\u3765", "る"),  # 㝥
        ("\u375e", "ず"),  # 㝞
        ("\u377c", "れ"),  # 㝼
        ("\u376d", "ん"),  # 㝭
        ("\u375c", "ず"),  # 㝜
        ("\u3771", "り"),  # 㝱
    ):
        text = text.replace(old, new)

    if text == orig:
        raise SystemExit("no changes")
    core_path.write_text(text, encoding="utf-8", newline="\n")
    print("ok", core_path, "delta", len(text) - len(orig))


if __name__ == "__main__":
    main()
