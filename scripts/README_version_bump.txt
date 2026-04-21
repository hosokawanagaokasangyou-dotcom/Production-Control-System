version.txt auto bump (+0.1 on each commit)
============================================

One-time per clone (stored in local .git/config, not pushed):
  git config core.hooksPath .githooks

Skip for one commit (e.g. git commit --amend without bumping again):
  PowerShell:  $env:SKIP_BUMP_CODE_VERSION = "1"; git commit ...
  cmd.exe:     set SKIP_BUMP_CODE_VERSION=1 && git commit ...

Target file: code/version.txt (same folder as the macro workbook in this repo)
