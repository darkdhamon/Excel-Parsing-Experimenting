# Changelog

All notable changes intended for `master` should be recorded in this file.

## Unreleased

### Changed

- Rebuilt the repository as an ASP.NET Core MVC application targeting `.NET 10`.
- Replaced the legacy ASP.NET MVC 5 / .NET Framework project structure with SDK-style projects.
- Implemented equivalent Fitbit parsing flows for `ExcelDataReader`, `EPPlus`, and `Open XML SDK`.
- Added sample-driven regression tests for the Fitbit workbook parsing paths.
- Added PR governance for future `master` merges through `AGENTS.md`, a PR template, and a changelog enforcement workflow.

## How To Update This File

- Every pull request targeting `master` must update this file.
- Add short, reviewable bullets under `Unreleased`.
- Record user-visible, workflow, or repository-significant changes.
- Keep the changelog wording close to the pull request summary so reviewers can match them quickly.
