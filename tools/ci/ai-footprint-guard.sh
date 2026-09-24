#!/usr/bin/env bash
set -euo pipefail

BASE_SHA="${1:-}"
HEAD_SHA="${2:-}"

if [[ -z "$HEAD_SHA" ]]; then
  echo "ai-footprint-guard: missing HEAD sha" >&2
  exit 2
fi

if [[ -z "$BASE_SHA" || "$BASE_SHA" == "0000000000000000000000000000000000000000" ]]; then
  if git rev-parse "${HEAD_SHA}^" >/dev/null 2>&1; then
    BASE_SHA="${HEAD_SHA}^"
  else
    BASE_SHA="$HEAD_SHA"
  fi
fi

AI_ANYWHERE='claude|anthropic|copilot|chatgpt|openai|gpt-[0-9]|codeium|tabnine|windsurf|deepseek|grok|mistral|perplexity'
AI_WORD='codex|kimi|gemini|llama'
PATTERN="$AI_ANYWHERE|\\b($AI_WORD)\\b"
COAUTHOR_PATTERN='co-?authored-?by:|generated with|authored by|written by|assisted by|created by|powered by'
CURSOR_PATTERN='(co-?authored|generated|written|authored|created|assisted|made|built|powered)([^a-z]|[a-z]){0,24}cursor|cursor\.(sh|so|com|ai|dev)|cursor[-_ ](ide|editor|ai|agent|composer|rules)|\.cursorrules|\.cursor/'
ROBOT="🤖"

HITS_FILE="$(mktemp)"
PATHS_FILE="$(mktemp)"
ADDED_LINES_FILE="$(mktemp)"
trap 'rm -f "$HITS_FILE" "$PATHS_FILE" "$ADDED_LINES_FILE"' EXIT

record_hit() {
  printf '%s\n' "$1" >> "$HITS_FILE"
}

# Commit messages in the incoming range
while IFS= read -r commit; do
  [[ -z "$commit" ]] && continue
  msg="$(git show -s --format=%B "$commit")"
  robot_hit=1
  case "$msg" in
    *"$ROBOT"*) robot_hit=0 ;;
  esac

  if printf '%s' "$msg" | grep -qiE "$PATTERN"; then
    record_hit "COMMIT_MESSAGE: $commit"
  elif printf '%s' "$msg" | grep -qiE "$COAUTHOR_PATTERN" && printf '%s' "$msg" | grep -qiE "$PATTERN|$CURSOR_PATTERN"; then
    record_hit "COMMIT_MESSAGE: $commit"
  elif [[ "$robot_hit" -eq 0 ]]; then
    record_hit "COMMIT_MESSAGE: $commit"
  fi
done < <(git rev-list "${BASE_SHA}..${HEAD_SHA}" 2>/dev/null || true)


# Added lines from changed files, computed once for the whole range
git diff -U0 --diff-filter=ACMR "$BASE_SHA" "$HEAD_SHA" \
  | awk '
      /^\+\+\+ / { next }
      /^--- / { next }
      /^@@/ { next }
      /^diff --git / { next }
      /^index / { next }
      /^\+/ { print substr($0, 2) }
    ' > "$ADDED_LINES_FILE" || true

if grep -qiE "$PATTERN" "$ADDED_LINES_FILE"; then
  record_hit "FILE_CONTENT: added lines in this change set"
elif grep -qiE "$CURSOR_PATTERN" "$ADDED_LINES_FILE"; then
  record_hit "FILE_CONTENT: added lines in this change set"
fi

# Changed file names (new/renamed/modified/copied)
git diff --name-only --diff-filter=ACMR -z "$BASE_SHA" "$HEAD_SHA" | tr '\0' '\n' > "$PATHS_FILE"
while IFS= read -r path; do
  [[ -z "$path" ]] && continue

  case "$path" in
    .githooks/*|*/.githooks/*|AGENTS.md|RULES.md|.ai-rules.md|*/AGENTS.md|*/RULES.md|*/.ai-rules.md|.github/workflows/ai-footprint-guard.yml|tools/ci/ai-footprint-guard.sh)
      continue ;;
  esac

  if printf '%s' "$path" | grep -qiE "$PATTERN"; then
    record_hit "FILE_NAME: $path"
  elif printf '%s' "$path" | grep -qiE "$CURSOR_PATTERN"; then
    record_hit "FILE_NAME: $path"
  fi

done < "$PATHS_FILE"

if [[ -s "$HITS_FILE" ]]; then
  echo "ai-footprint-guard: rejected, AI footprint detected:" >&2
  sort -u "$HITS_FILE" | sed 's/^/  /' >&2
  exit 1
fi

echo "ai-footprint-guard: no AI footprints detected."
