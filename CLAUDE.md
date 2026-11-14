# Gridworks

Gridworks is a spreadsheet application that provides the main features of
Microsoft Excel in the browser, as a static TypeScript app on GitHub Pages.
It is built with the `nightshift` workflow. What it must do is in
`docs/requirements.md`, how it is built is in `docs/architecture.md`, and
progress is in `docs/backlog.md`.

## Commands

- `make check`: Prettier check, ESLint, `tsc` for the engine (without the DOM
  library) and for the whole project, Vitest with the 90% line coverage
  floor, and the document check. Run it after every change.
- `make e2e`: builds the production bundle, serves it under `/gridworks/`, and
  runs the Playwright tests on Chromium, Firefox and WebKit. Run it before a
  pull request. It needs `npx playwright install` once.
- `npm run dev`: runs the app locally at `http://localhost:5173/gridworks/`.
- `npx prettier --write .`: formats everything.

## Rules

- Work on one slice at a time, with `/nightshift:build`. Tests come first.
- The simplest design and code that meet the requirements. Nothing for a
  need that is not stated.
- Deep modules: small interfaces that hide substantial behavior.
- Tests go through public interfaces only and cover the core logic. None is
  redundant or written to raise coverage.
- Never commit to `main`. One branch and one squash-merged pull request per
  slice.
- Code, comments, documents, and commit messages are in English.
- Update the affected documents in the same change as the behavior.
- Do not edit `scripts/check-docs.py`; it is a copy from the plugin.
- `src/engine` has no DOM and never imports `src/app`. Only
  `src/app/platform/web` touches storage, files, the clipboard or printing;
  everything else goes through the `Platform` interface. Lint enforces both.
- In `src/app`, code that touches the DOM goes in `.tsx` files or in
  `src/app/platform/web`; `.ts` files there are unit tested and count toward
  the coverage floor.
