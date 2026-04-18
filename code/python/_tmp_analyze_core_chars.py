# -*- coding: utf-8 -*-
import collections
from unicodedata import name

p = "planning_core/_core.py"
text = open(p, encoding="utf-8").read()
ctr = collections.Counter(text)
for ch, n in ctr.most_common(120):
    o = ord(ch)
    if ch in " \n\t\r":
        continue
    if o < 128 and ch.isascii():
        continue
    try:
        nm = name(ch)
    except ValueError:
        nm = "?"
    if n > 30 and ("CJK" in nm or "HIRAGANA" in nm or "KATAKANA" in nm or o == 0x301D):
        print(hex(o), repr(ch), n, nm[:50])
