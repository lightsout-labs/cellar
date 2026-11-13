.PHONY: check e2e

check:
	npx prettier --check .
	npx eslint .
	npx tsc --noEmit -p src/engine/tsconfig.json
	npx tsc --noEmit -p tsconfig.json
	npx vitest run --coverage
	python3 scripts/check-docs.py

e2e:
	npx playwright test
