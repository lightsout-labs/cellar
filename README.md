# Cellar

Cellar is a spreadsheet application that provides the main features of
Microsoft Excel in the browser, including opening and saving real `.xlsx`
files, with no installation and no account. It is written in TypeScript, runs
entirely client-side, and is deployed to GitHub Pages.

**Use it at <https://lightsout-labs.github.io/cellar/>.**

The requirements are in [docs/requirements.md](docs/requirements.md), the
design in [docs/architecture.md](docs/architecture.md), and progress in
[docs/backlog.md](docs/backlog.md).

## Development

Cellar needs the Node.js version in [.nvmrc](.nvmrc) and `make`.

| Task | Command |
|---|---|
| Install | `npm ci`, then `npx playwright install` once for the end-to-end tests |
| Run locally | `npm run dev`, then open <http://localhost:5173/cellar/> |
| Test | `make check` (format, lint, types, unit tests with coverage, documents) and `make e2e` (the built app in Chromium, Firefox and WebKit) |
| Build | `npm run build`, which writes the static site to `dist/` |

Every merge to `main` is deployed to the address above by
[a GitHub Actions workflow](.github/workflows/deploy.yml), which then runs a
smoke test against it.

## License

[MIT](LICENSE)
