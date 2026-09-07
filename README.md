# Rutba-Office
Rutba Office

## No AI Footprint Policy

This repository must never contain any footprint that identifies an AI
tool, assistant, or provider as an author, co-author, or generator of
content. This includes, but is not limited to:

* Commit trailers that credit an AI tool or assistant as a co-author of a
  commit (git's mechanism for crediting multiple authors on one commit).
* Comments, headers, commit messages, or file content disclosing that an
  AI tool played any part in producing the content (including disclosures
  added to satisfy AI transparency/labeling requirements).
* Any AI provider, product, or tool name appearing anywhere in the
  repository's files, commit history, or metadata.

The exact list of disallowed patterns lives in
[`scripts/ai-footprint-patterns.txt`](scripts/ai-footprint-patterns.txt) and
is enforced by [`scripts/check-ai-footprint.sh`](scripts/check-ai-footprint.sh).

### Automated enforcement

* A GitHub Actions workflow (`.github/workflows/ai-footprint-check.yml`)
  runs the check on every push and pull request and fails the build if a
  violation is found.
* Contributors can enable the same check locally as git hooks:

  ```sh
  git config core.hooksPath .githooks
  ```

  This blocks commits (`pre-commit`) and commit messages (`commit-msg`,
  including multi-author trailers) that contain disallowed content before
  they are ever created. 
