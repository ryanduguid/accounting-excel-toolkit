# Releasing

Releases are built by GitHub Actions from an annotated tag on the exact `main` commit. Do not create or upload release assets by hand.

Before tagging:

1. Merge the release pull request and require every `main` check to pass.
2. Enable release immutability in the repository settings.
3. From an operator session authenticated with repository Administration read access, run:

    ```bash
    gh api -H "X-GitHub-Api-Version: 2026-03-10" repos/ryanduguid/SirAlexanderFitzgerald/immutable-releases --jq .enabled
    ```

    Do not push the tag unless the output is exactly `true`. The Actions `GITHUB_TOKEN` cannot be granted repository Administration read access, so the tag workflow cannot perform this preflight itself.
4. Confirm `VERSION` is the intended version and the first line of `RELEASE_NOTES.md` is the matching tag.
5. Create an annotated tag on the current remote `main` commit, for example `git tag -a v0.1.2 -m "v0.1.2"` (use `-s` instead of `-a` when a signing key is configured), then push only that tag.

The existing `v0.1.1` tag is retained as immutable history. [Pilot run 31822769922](https://github.com/ryanduguid/SirAlexanderFitzgerald/actions/runs/31822769922) stopped on the now-removed workflow attempt to read this Administration setting, before any build or publication step; no `v0.1.1` release or assets were created. Do not move or delete that tag. The next release is `v0.1.2`.

The workflow reruns the regression suite, builds deterministic ZIP and tar.gz source archives, generates an SPDX 2.3 SBOM and `SHA256SUMS`, records GitHub provenance and SBOM attestations, then publishes a draft release only after every asset is uploaded. The archive helper fixes the timezone to UTC and Git text conversion to LF so the same tagged tree produces the same archive bytes on Linux and Windows. Existing releases are refused rather than overwritten.

After publication, download the assets and verify them:

```bash
gh release download v0.1.2 -R ryanduguid/SirAlexanderFitzgerald --dir release-v0.1.2
cd release-v0.1.2
sha256sum --check SHA256SUMS
gh attestation verify accounting-excel-toolkit-0.1.2.zip -R ryanduguid/SirAlexanderFitzgerald
gh attestation verify accounting-excel-toolkit-0.1.2.zip -R ryanduguid/SirAlexanderFitzgerald --predicate-type https://spdx.dev/Document/v2.3
gh release view v0.1.2 -R ryanduguid/SirAlexanderFitzgerald --json isImmutable
gh release verify v0.1.2 -R ryanduguid/SirAlexanderFitzgerald
gh release verify-asset v0.1.2 accounting-excel-toolkit-0.1.2.zip -R ryanduguid/SirAlexanderFitzgerald
```

If any gate fails, leave the tag and any draft release untouched until the failure is understood. Never move an already published tag.
