# v0.1.2

The `v0.1.1` tag is retained as an unreleased failed-preflight tag. Its workflow stopped before any build or publication step, so it has no release or assets and will not be moved or deleted.

This release keeps the toolkit source-only and adds the changes made since `v0.1.0`:

- add native Microsoft Excel acceptance coverage for the Power Query functions;
- harden Excel COM cleanup and failure reporting in that acceptance harness;
- correct the documented boundary between static guards and native Excel acceptance; and
- add workflow-built source archives, SHA-256 checksums, an SPDX SBOM and GitHub build attestations.

No client data or binary workbook is included. The CSV fixtures are fabricated.
