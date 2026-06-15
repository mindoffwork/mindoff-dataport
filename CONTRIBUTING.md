# Contributing to Mindoff Dataport

Contributions to Mindoff Dataport are welcome: bug fixes, docs improvements, performance work, or features that grew out of something you actually needed.

The best contributions usually come from real friction. If a template behaved in a surprising way, a workflow felt awkward, or a docs gap slowed you down, that is useful context for the project.

Before opening an issue or pull request, it helps to read:

- The [README](https://github.com/mindoffwork/mindoff-dataport/blob/root/README.md) for the project flow and quick start
- The [Developer Guide](https://dataport.mindoff.work/latest-release/developer_guide/) for user-facing behavior and constraints
- The [Architecture docs](https://dataport.mindoff.work/latest-release/architecture/) for internals and rendering details

## Ways to Contribute

- Build new features or extend existing capabilities
- Fix bugs and security concerns
- Improve or optimize the codebase, including performance
- Improve documentation, fill gaps, or resolve inconsistencies
- Open issues to report bugs or highlight documentation gaps
- Use discussions to request features, ask questions, answer questions, or share what you built

Mindoff Dataport follows a linear versioning system where only the latest version is actively supported. Before opening any issue, discussion, or pull request:

- Search the GitHub repository for related items first
- Make sure you are using the latest version of `mindoff-dataport`

## Code of Conduct

Be respectful, open-minded, and constructive.

Read the repo guidelines here:

- [Code of Conduct](https://github.com/mindoffwork/mindoff-dataport/blob/root/.github/CODE_OF_CONDUCT.md)

## Issues and Discussions

### Issues

Open an issue when something is broken, incorrect, or needs correction.

Good issue candidates:

- Bugs or incorrect output
- Security concerns
- Incorrect documentation
- Compatibility problems with upcoming Python versions or core dependencies (`openpyxl`, `xlsxwriter`, `reportlab`, `pyarrow`, `polars`)

You can also help by confirming bugs, attaching a minimal reproduction, or submitting a pull request that fixes the issue.

### Discussions

Use GitHub Discussions for everything that is not a bug.

Good discussion candidates:

- Feature ideas
- Usage questions
- Architectural ideas
- Projects built with `mindoff-dataport`

Feature ideas are welcome when they fit the project's direction. Open them in Discussions, not Issues, and tag them with the `feature` label.

Avoid requesting behavior changes to existing features unless the change is needed for security reasons or future compatibility with core dependencies.

## Development Setup

### 1. Fork or Clone the Repository

Option A: fork in the GitHub UI, then clone your fork locally:

```bash
git clone https://github.com/mindoffwork/mindoff-dataport.git
cd mindoff-dataport
```

Add the upstream remote so you can sync from the main repository:

```bash
git remote add upstream https://github.com/mindoffwork/mindoff-dataport.git
git fetch upstream
git rebase upstream/root
```

Option B: clone directly. This is usually only for maintainers and collaborators:

```bash
git clone https://github.com/mindoffwork/mindoff-dataport.git
cd mindoff-dataport
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
source .venv/bin/activate
```

On Windows PowerShell:

```powershell
.venv\Scripts\Activate.ps1
```

### 3. Install Dependencies

```bash
pip install -e ".[dev]"
```

If you plan to edit documentation too:

```bash
pip install -e ".[dev,docs]"
```

To preview the docs locally:

```bash
mkdocs serve
```

### 4. Run Tests

Before submitting a pull request, make sure the test suite passes and coverage stays at 90% or higher.

```bash
PYTHONPATH=src python -m pytest -q
```

Coverage gate:

```bash
PYTHONPATH=src python -m pytest --cov=mindoff_dataport --cov-branch --cov-report=term-missing:skip-covered --cov-fail-under=90
```

If your change touches the build or serialization path, also run the roundtrip guard:

```bash
PYTHONPATH=src python -m pytest -q src/tests/test_roundtrip.py::test_roundtrip_schema_identical
```

## Branches and Commits

### Branch Naming

Recommended branch prefixes:

- `feature/*`: New user-facing functionality or capabilities
- `bug/*`: Fixes for incorrect behavior, regressions, or defects
- `documentation/*`: Docs-only updates (guides, README, examples)
- `enhancement/*`: Improvements to existing behavior without adding a new feature
- `internal/*`: Refactors, tooling, CI changes, or other non-user-facing maintenance

Examples:

```text
feature/streaming-pivot-tables
bug/pdf-merge-chunk-boundary
documentation/contribution-guide
enhancement/shift-performance
internal/ci-changelog-step
```

### Commit Messages

Use Gitmoji codes in commit messages. Keep the message short and action-oriented.

Project standard:

```text
:<gitmoji_code>: <Verb> <short description>
```

Example:

```text
:sparkles: Add streaming pivot table support
```

## Pull Requests

### 1. Open the Pull Request

1. Push your branch: `git push origin <branch-name>`
2. Open a pull request targeting the `root` branch on `mindoffwork/mindoff-dataport`
3. If you are contributing from a fork, use GitHub's fork PR flow:
   https://docs.github.com/en/pull-requests/collaborating-with-pull-requests/proposing-changes-to-your-work-with-pull-requests/creating-a-pull-request-from-a-fork

### 2. Pull Request Checklist

- Run the relevant tests
- Keep coverage at 90% or higher
- Follow the coding conventions and project structure already in the repo
- Update documentation when behavior changes, or say clearly in the PR description that no docs change was needed
- Run `black` before submitting

### 3. Pull Request Title

Pattern:

```text
{emoji} {Verb-starting phrase}
```

Use the actual emoji in the PR title, not the `:gitmoji:` code, so it renders correctly in release notes. Use `:sparkles:`-style codes in commit messages, and raw emoji in PR titles.

Good example:

```text
✨ Add streaming pivot table support
```

Bad example:

```text
:sparkles: Added streaming pivot table support.
```

### 4. Pull Request Label

Pick the single label that best matches the main change:

- `feature`: New functionality or capability
- `bug`: Fix for incorrect behavior or regression
- `documentation`: Docs-only changes
- `enhancement`: Improvement to existing behavior without a new feature
- `internal`: Refactors, tooling, CI, or other changes that do not affect users directly

### 5. Pull Request Description

PR descriptions are optional when the title is already self-explanatory, or when you are also the reviewer. When you do write one, keep it short and useful.

Suggested structure:

- Summary
- Motivation / context
- Testing
- Risks / migrations, if applicable

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

Please respect the maintainers' time.

AI tools can generate code, pull requests, and comments very quickly. Reviewing them still takes human time. Flooding a project with automated or half-checked submissions shifts the cost onto reviewers.

AI can still be genuinely useful here. Just make sure you review, understand, and test what you submit. The goal is not to avoid tools. The goal is to avoid handing reviewers work you did not finish yourself.

Pull requests, issues, or discussion threads that do not follow this guide may be closed or removed without review.
