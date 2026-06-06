# Contribution Guide

Contributions to Mindoff Dataport are welcome: a bug fix, a docs gap that tripped you up, or a feature that grew out of something you needed in practice.

The most useful contributions tend to come from people who hit the actual problem. If you worked around something or tracked down why a template behaved unexpectedly, that experience is worth sharing.

## Ways to Contribute

- ✨ Build new features or extend existing capabilities.
- 🐛 Fix bugs and security concerns.
- ⚡ Improve or optimize the codebase, including performance.
- 📝 Improve documentation, fill gaps, or resolve inconsistencies.
- 🚨 Open Issues to report bugs or highlight documentation gaps.
- 💬 Use Discussions to request features, ask/answer questions, or share what you built.

Mindoff Dataport follows a linear versioning system where only the latest version is actively supported. Before opening any issue, discussion, or pull request:

- Search the GitHub repository for related items to avoid duplicates.
- Make sure you're using the latest version of the `mindoff-dataport` package.

## Code of Conduct

> **Be respectful, open-minded, and constructive.**

Read the guidelines in the repo: [Code of Conduct Document](https://github.com/mindoffwork/mindoff-dataport/blob/root/.github/CODE_OF_CONDUCT.md)

## Issues & Discussions

### Issues

Open an Issue only when something is broken or needs correction, following the issue template carefully.

Examples:

- Bugs or incorrect output
- Security concerns
- Incorrect documentation
- Compatibility problems with upcoming Python versions or core dependencies (openpyxl, xlsxwriter, reportlab, pyarrow, polars)

You can also contribute by confirming bugs, sharing reproducible examples, or submitting a pull request that fixes the issue.

### Discussions

Use Discussions for everything that is not a bug.

Examples:

- Feature ideas
- Questions about usage
- Architectural ideas
- Sharing projects built with mindoff-dataport

Mindoff Dataport's features are curated by the maintainers. Reasonable feature ideas that align with the project's principles are welcome. Open feature ideas in **Discussions (not Issues)** and tag them with the `feature` label.

Avoid requesting changes to the behavior of existing features unless they are required for security reasons or to ensure compatibility with future versions of core dependencies.

## Development

### 1. Fork or Clone the Repository

**Option A: Fork**

Click **Fork** in the GitHub UI to create your own copy, then clone your fork locally:

```bash
git clone https://github.com/mindoffwork/mindoff-dataport.git
cd mindoff-dataport
```

Add the upstream remote (so you can sync changes from the original repo):

```bash
git remote add upstream https://github.com/mindoffwork/mindoff-dataport.git
git fetch upstream
git rebase upstream/root
```

**Option B: Clone directly (recommended only for maintainers and collaborators)**

```bash
git clone https://github.com/mindoffwork/mindoff-dataport.git
cd mindoff-dataport
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
source .venv/bin/activate         # Windows: .venv\Scripts\Activate.ps1
```

### 3. Install Dependencies

```bash
pip install -e ".[dev]"
```

If you plan to edit documentation, install the docs dependencies as well:

```bash
pip install -e ".[dev,docs]"
```

Mindoff Dataport uses `mkdocs` for documentation. Preview it live and open the local URL shown in the terminal:

```bash
mkdocs serve
```

### 4. Running Tests

Before submitting a pull request, ensure the test suite passes and that **coverage is at 90% minimum**.

```bash
PYTHONPATH=src python -m pytest -q
```

With the coverage gate:

```bash
PYTHONPATH=src python -m pytest --cov=mindoff_dataport --cov-branch --cov-report=term-missing:skip-covered --cov-fail-under=90
```

If your change touches the build or serialization path, also run the roundtrip guard:

```bash
PYTHONPATH=src python -m pytest -q src/tests/test_roundtrip.py::test_roundtrip_schema_identical
```

### 5. Branch Naming Strategy

_Good practice, recommended for maintainers and collaborators._

All branches must start with one of the following prefixes:

- `feature/*`: New user-facing functionality or capabilities.
- `bug/*`: Fixes for incorrect behavior, regressions, or defects.
- `documentation/*`: Docs-only updates (guides, README, examples).
- `enhancement/*`: Improvements to existing behavior without adding a new feature.
- `internal/*`: Refactors, tooling, CI changes, or non-user-facing maintenance.

Examples:

```text
feature/streaming-pivot-tables
bug/pdf-merge-chunk-boundary
documentation/contribution-guide
enhancement/shift-performance
internal/ci-changelog-step
```

### 6. Commit Message Guidelines

Use [Gitmoji](https://gitmoji.dev) codes in commit messages that match the change type. Keep the text short and action-oriented. Favor clarity over creativity. The project standard is `:<gitmoji_code>: <Verb> <short description>`.

Example:

```text
:sparkles: Add streaming pivot table support
```

### 7. Open a Pull Request

1. Push your branch: `git push origin <branch-name>`
2. Open a Pull Request targeting the `root` branch on `mindoffwork/mindoff-dataport`. If you are using a fork, read about [creating a pull request from a fork](https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request-from-a-fork).

## Pull Request Guidelines

- Ensure tests pass and coverage is at least 90%, as mentioned earlier.
- Follow the coding conventions and project structure used in the existing codebase. See the [Architecture documentation](../architecture/index.md) for details.
- Update documentation based on your changes. Otherwise, mention clearly in the PR description that it's not documented.
- Run `black` formatting before submitting.

### 1. PR Title

Pattern:

```
{emoji} {Verb-starting phrase}
```

- Use the **actual emoji** in place of `emoji` (not the `:gitmoji:` code) so it renders correctly in release notes. Use `:sparkles:`-style codes in commit messages, and raw emoji in PR titles.
- The `Verb-starting phrase` should not start with a participle verb and should not end with a period; it's a phrase, not a sentence. (Use "Add", not "Added".)

Good example ✅:

```
✨ Add streaming pivot table support
```

Bad example ❌:

```
:sparkles: Added streaming pivot table support.
```

### 2. PR Labels (Required)

Pick the single label that best matches the main change in your branch. Do not choose more than one:

- `feature`: New functionality or capability.
- `bug`: Fix for incorrect behavior or regression.
- `documentation`: Docs-only changes.
- `enhancement`: Improvements to existing behavior without a new feature.
- `internal`: Refactors, tooling, CI, or changes that don't affect users.

### 3. PR Description

<div class="admonition question">
  <p class="admonition-title">When it's optional</p>
  <p>PR descriptions are optional if the title itself is self-explanatory, or if you are the reviewer of the PR you're creating.</p>
</div>

Keep the title and description simple, accurate, and easy to understand. Clearly explain what changed and why, how to test it (including specific commands or steps), and any potential risks or breaking changes. Always review before submitting.

**General rule:** a good PR description should take reviewers less time to read than it took you to write it.

Suggested structure (styling encouraged, if you have the time):

- Summary
- Motivation / context
- Testing
- Risks / migrations (if applicable)

Example:

```text
Summary:
Repeat dataframe headers across paginated PDF chunks.

Motivation / context:
Long tables lost their column labels on later pages. This repeats the header where a dataframe-header anchor exists.

Testing:
PYTHONPATH=src python -m pytest -q src/tests/test_pdf_renderer.py

Risks / migrations:
None. Opt-in via repeat_dataframe_headers=True; default unchanged.
```

## One Last Thing

Please respect the maintainers' time and effort.

AI tools can churn out code, PRs, or comments in a flash, but reviewing them still takes real human work on our end. Flooding the repo with automated or half-baked submissions creates a kind of Human Effort Denial-of-Service attack on the project.

That said, modern tools like AI can be a huge help when used wisely. They speed up development, boost code quality, and help you learn faster. Just don't let them replace careful thinking and intent.

If you're using AI or similar tools, always review, understand, and test your changes before submitting a PR. Make sure your submissions are spot-on, genuinely useful, and easy on the reviewer.

<div class="admonition failure">
  <p class="admonition-title">Rejection note</p>
  <p>Pull requests, Issues, or Discussion threads that do not follow this guide properly may get rejected, closed, or removed from the repo without review or notice.</p>
</div>
