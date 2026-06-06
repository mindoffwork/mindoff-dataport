# Deployment and Versioning

<div class="admonition info">
<p class="admonition-title">For maintainers</p>
<p>This is the day-to-day guide for shipping releases and docs for <code>mindoff-dataport</code>. If you are a contributor or user, feel free to skim and adapt it for your own projects.</p>
</div>

Mindoff Dataport uses linear versioning: only the latest release is supported for security and bug fixes. Some teams may stay on an older feature-complete release, so documentation is preserved per release series to keep those projects unblocked.

Releases and their changes are tracked in the [Release Notes](../release_notes.md), the source of truth for what changed in each version and the best starting point when upgrading.

## Deployment Workflow

### Deploy the Package

Releases are created in GitHub and finalized by the CD pipeline (`.github/workflows/cd.yml`).

Create a tag in the format `vX.Y.Z` and push it:

```bash
git tag -a v0.7.0 -m "Release v0.7.0"
git push origin v0.7.0
```

Then publish the tag as a GitHub Release. On publish, the pipeline:

- Validates the tag matches `vX.Y.Z`.
- Replaces the `## Recent Changes` heading in `CHANGELOG.md` with the release version and commits it.
- Syncs `pyproject.toml`'s `version` to the tag (committing if it differed).
- Builds the distribution, runs `twine check`, and publishes to PyPI.
- Reverts the changelog heading if the publish step fails.

Once the pipeline finishes cleanly, the release is live on PyPI.

### Deploy the Documentation

Source docs live in `src/mindoff_dataport/docs/` (inside the package, so they ship on PyPI too). Rendered docs are built with MkDocs + Mike and pushed to the `docs` branch, which GitHub Pages serves.

<div class="admonition note">
<p class="admonition-title">Key points to remember</p>
<ul>
<li>The <code>docs</code> branch contains generated output only. Do not edit it by hand.</li>
<li>Both the <code>root</code> branch and the <code>docs</code> branch are protected. Do not merge one into the other.</li>
</ul>
</div>

The public documentation site is:

```
https://dataport.mindoff.work/
```

Before deploying, make sure your working tree is clean, `src/mindoff_dataport/docs/` contains the final content, and you are on the correct branch.

Example for the `0.7` series:

```bash
mike deploy --branch docs --push --update-aliases v0.7 latest-release
mike set-default --branch docs --push latest-release
```

This builds the docs, publishes them to the `docs` branch, and points the `latest-release` alias at the new version.

### Update an Existing Documentation Version

To fix docs within the same series without creating a new version:

```bash
mike deploy --push 0.7
```

This replaces the existing rendered docs for that series.

## Versioning Policy

Release tags always follow the `vX.Y.Z` format:

| Component | Meaning       |
| --------- | ------------- |
| X         | Major version |
| Y         | Minor version |
| Z         | Patch version |

Rules we follow:

1. A release is official only after the tag is published as a GitHub Release.
2. The GitHub Release title must match the tag exactly.
3. The CD pipeline updates `pyproject.toml` from the tag.
4. Only the latest release line receives fixes and security updates.

### Documentation Versioning

Docs versions are created per major release series, not per patch. Patch releases update the existing docs for that series. The version selector on the site is driven by Mike's `versions.json`.

| Release | Documentation Version |
| ------- | --------------------- |
| v0.7.0  | 0.7                   |
| v0.7.3  | 0.7                   |
| v1.0.0  | 1.0                   |

**Documentation aliases**

The alias `latest-release` always points to the newest published docs version, so users land on the recommended docs by default.

| Alias          | Points to                          |
| -------------- | ---------------------------------- |
| latest-release | 1.0                                |
| 0.7            | documentation for the 0.7.x series |

## Rollbacks

### Release Rollback

1. Unpublish or delete the GitHub Release for the bad tag.
2. To reuse the same tag for another release, delete the tag itself.
3. Yank the bad release on PyPI so new installs no longer resolve to it.
4. If a corrected release is needed, create a new patch tag and publish it as a new GitHub Release.
5. The CD pipeline updates `pyproject.toml` from the corrected tag.

When you're ready to cut the corrected release, return to the [Deployment Workflow](#deploy-the-package).

### Documentation Rollback

If a docs deploy introduces an issue, redeploy the last known good version:

```bash
mike delete --branch docs --push v0.7
mike deploy --push <version>
```

This restores the previous docs for that series.
