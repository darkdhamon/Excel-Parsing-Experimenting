# AGENTS.md

Follow these rules before making changes in this repository:

## Repository Workflow

- Read this file before starting work.
- Do not push directly to `master`.
- Make changes on a feature branch and deliver them through a pull request into `master`.
- If a conversation starts with an issue number, create the feature branch from `dev`. If `dev` does not exist, branch from `master`.
- When tracked changes are made, commit and push them.

## Changelog Rules

- Every pull request targeting `master` must update [CHANGELOG.md](./CHANGELOG.md).
- Add a concise description of the pull request's repo-significant changes under the `Unreleased` section.
- Keep the changelog entry aligned with the pull request summary.
- Update the changelog before committing the branch that will be opened against `master`.

## Tooling

- If Visual Studio is required, use Visual Studio 2026 Insider Preview.
- Prefer the existing repo conventions and avoid introducing extra project layouts unless the task requires them.
