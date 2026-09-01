# No CI workflow here, and why

This repository has no `.github/workflows`. That is a decision, and this file is
it -- so the absence is on record rather than looking like an oversight.

An empty `.github/workflows` directory is worse than no directory: it reads as a
promise of a check with nothing behind it, and git cannot track an empty
directory anyway, so it is invisible to everyone except whoever is looking at
that working copy. The release tooling refuses on both states by name.

## What this package has to run

- **own specs:** 12
- **scripts:** `build`, `clean`, `lint`, `typecheck`

## Why none of it can run here yet

Measured from a clean checkout -- the git-tracked files only, copied into an
empty directory in a container with no sibling packages visible, which is what
CI gets:

```
npm install
npm error 404 Not Found - GET https://registry.npmjs.org/@coolms%2fcore-angular
```

npm auto-installs peer dependencies, so an **unpublished peer 404s the whole
install as hard as a dependency would**. This package declares these `@coolms/*`
peers, and none of them is published:

- `@coolms/core-angular`
- `@coolms/document-engine`
- `@coolms/ui-angular`

With `--legacy-peer-deps` the install succeeds and the remaining steps fail on the same missing modules, so the workaround moves the failure rather than removing it.

A workflow added today would be red on its first run, or would have to be
written around the failure -- a green check that verifies nothing, which is the
one thing worse than no check.

## What does cover this package

All 12 of this package's specs run in the admin SPA suite, which imports each one by path from the application's own spec bridge. That suite last ran 1215 of 1215 green. The tests are real and they do execute -- just not here, and not as a gate on a pull request opened against this repository.

## What would unblock it

Publishing the `@coolms/*` packages to npm. That is gated on `private: true`
being lifted in their manifests -- a decision, not a missing file. Once the
peers resolve, this repository can install standalone, the workflow becomes a
real gate, and branch protection can follow it.

## The consequence, stated plainly

Until then this repository has **no gate**: code reaches `develop` with nothing
standing in the way. Branch protection cannot substitute for it, because a
required status check that no workflow ever produces blocks every pull request
forever.
