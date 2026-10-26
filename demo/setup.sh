#!/usr/bin/env bash
# Hidden setup for demo.tape: a throwaway workspace with the generated, brand-free
# fixtures. Sourced from the repo root; nothing here touches the repo.
REPO="$(pwd)"
export PROFILE="$REPO/skills/pptx-style-profile/scripts"
export BUILD="$REPO/skills/pptx-style-build/scripts"
export PPTX_STYLE_FONT_DIRS="$REPO/tests/fonts"
WS="$(mktemp -d)/workspace"
mkdir -p "$WS" && cd "$WS" || return 1
python3 "$REPO/tests/fixtures/make_fixtures.py" . >/dev/null
cp "$REPO/demo/spec.json" "$REPO/demo/offcontract.json" .
export PS1='\[\e[2m\]$\[\e[0m\] '
