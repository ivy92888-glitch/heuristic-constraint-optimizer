# Push guide (GitHub)

This repository was published from a local working copy using the GitHub REST
API (the `git push` HTTPS path is blocked in the build environment, so files are
uploaded individually via the Contents API). These are the steps that were used;
they are recorded so the process can be repeated.

## Steps performed

1. Create the repository (public, default branch `main`):
   `POST /user/repos` with `{"name":"heuristic-constraint-optimizer", ...}`.
2. For each file, base64-encode it and `PUT /repos/ivy92888-glitch/
   heuristic-constraint-optimizer/contents/<path>` with
   `{"message":"Add <path>","content":"<base64>"}`.

## Files uploaded

- `README.md`
- `LICENSE`
- `PUSH_GUIDE.md`
- `.gitignore`
- `.gitattributes`
- `docs/Configuration.md`
- `src/ConstrainedAllocator.bas`
- `examples/sample-data.xlsx`

## Note on credentials

The personal access token used for publishing was passed only through an
environment variable and was **never written to disk or committed**. Rotate /
delete the token after use from GitHub &rarr; Settings &rarr; Developer settings
&rarr; Personal access tokens.
