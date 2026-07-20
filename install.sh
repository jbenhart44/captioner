#!/usr/bin/env bash
# captioner install — clone-and-symlink into the Claude Code skills directory.
#
# Usage:
#   git clone https://github.com/jbenhart44/captioner.git
#   cd captioner
#   bash install.sh
#
# Then restart Claude Code and run:  /captioner <path-to-deck-or-folder>
set -euo pipefail

SKILLS_DIR="${CLAUDE_SKILLS_DIR:-$HOME/.claude/skills}"
REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="$SKILLS_DIR/captioner"

echo "captioner installer"
echo "  source : $REPO_DIR"
echo "  target : $TARGET"

echo
echo "This installer will touch exactly one path on your system:"
echo "  $TARGET"
echo "It creates that path as a symlink pointing at this checkout. If a captioner"
echo "install already exists there it is replaced (announced below before removal)."
echo "Nothing outside that one path is read, written, or deleted."
echo

mkdir -p "$SKILLS_DIR"

# Resolve the existing target (if any) and this checkout to canonical paths so
# absolute-vs-relative and trailing-slash variants compare equal.
REPO_CANON="$(readlink -f "$REPO_DIR" 2>/dev/null || echo "$REPO_DIR")"

if [ -e "$TARGET" ] || [ -L "$TARGET" ]; then
  TARGET_CANON="$(readlink -f "$TARGET" 2>/dev/null || echo "$TARGET")"
  if [ -L "$TARGET" ] && [ "$TARGET_CANON" = "$REPO_CANON" ]; then
    echo "  note   : $TARGET already links to this checkout — nothing to do"
  else
    echo "  action : removing existing $TARGET (currently -> $TARGET_CANON) before relinking"
    rm -rf "$TARGET"
    ln -s "$REPO_DIR" "$TARGET"
    echo "  linked : $TARGET -> $REPO_DIR"
  fi
else
  ln -s "$REPO_DIR" "$TARGET"
  echo "  linked : $TARGET -> $REPO_DIR"
fi

echo
echo "Checking Python dependencies..."
MISSING=0
if python3 -c "import pptx" 2>/dev/null;        then echo "  python-pptx   : OK"; else echo "  python-pptx   : MISSING"; MISSING=1; fi
if python3 -c "import lxml" 2>/dev/null;         then echo "  lxml          : OK"; else echo "  lxml          : MISSING"; MISSING=1; fi
if python3 -c "import PIL" 2>/dev/null;          then echo "  Pillow        : OK"; else echo "  Pillow        : MISSING"; MISSING=1; fi
if python3 -c "import spellchecker" 2>/dev/null; then echo "  pyspellchecker: OK"; else echo "  pyspellchecker: MISSING"; MISSING=1; fi
if [ "$MISSING" -eq 1 ]; then
  echo
  echo "  Some dependencies are missing. Install them with:"
  echo "    pip install -r \"$REPO_DIR/requirements.txt\""
fi

echo
echo "Done. Restart Claude Code, then run:  /captioner <path-to-.pptx-or-folder>"
