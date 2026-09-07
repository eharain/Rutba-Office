#!/usr/bin/env bash
# Scan tracked files and commit messages (including trailers such as
# "Co-authored-by:") for disallowed AI-footprint content, as defined in
# scripts/ai-footprint-patterns.txt.
#
# Exit status is 0 when clean, non-zero when a match is found.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PATTERNS_FILE="$REPO_ROOT/scripts/ai-footprint-patterns.txt"

# Files that are allowed to reference the banned terms because they define
# or implement the policy itself, not because they add an AI footprint.
SELF_EXCLUDES=(
  "scripts/ai-footprint-patterns.txt"
  "scripts/check-ai-footprint.sh"
  ".githooks/pre-commit"
  ".githooks/commit-msg"
)

if [[ ! -f "$PATTERNS_FILE" ]]; then
  echo "AI footprint pattern file not found: $PATTERNS_FILE" >&2
  exit 1
fi

mapfile -t PATTERNS < <(grep -vE '^[[:space:]]*(#|$)' "$PATTERNS_FILE")

if [[ ${#PATTERNS[@]} -eq 0 ]]; then
  echo "No AI footprint patterns configured." >&2
  exit 1
fi

fail=0

is_self_excluded() {
  local candidate="$1"
  local excluded
  for excluded in "${SELF_EXCLUDES[@]}"; do
    if [[ "$candidate" == "$excluded" ]]; then
      return 0
    fi
  done
  return 1
}

echo "Scanning tracked files for AI footprints..."
while IFS= read -r -d '' file; do
  if is_self_excluded "$file"; then
    continue
  fi
  for pattern in "${PATTERNS[@]}"; do
    if grep -inE -a "$pattern" -- "$REPO_ROOT/$file" 2>/dev/null; then
      echo "::error file=${file}::Disallowed AI footprint pattern '${pattern}' found in ${file}"
      fail=1
    fi
  done
done < <(git -C "$REPO_ROOT" ls-files -z)

echo "Scanning commit messages for AI footprints..."
while IFS= read -r commit; do
  [[ -z "$commit" ]] && continue
  message="$(git -C "$REPO_ROOT" log -1 --format='%B' "$commit")"
  for pattern in "${PATTERNS[@]}"; do
    if grep -inE "$pattern" <<<"$message" >/dev/null; then
      echo "::error::Commit ${commit} message contains disallowed AI footprint pattern '${pattern}'"
      fail=1
    fi
  done
done < <(git -C "$REPO_ROOT" log --all --format='%H' 2>/dev/null || true)

if [[ "$fail" -ne 0 ]]; then
  echo "" >&2
  echo "AI footprint check FAILED: remove any AI provider/tool names, generated-by disclosures, and Co-authored-by trailers before committing." >&2
  exit 1
fi

echo "AI footprint check passed."
