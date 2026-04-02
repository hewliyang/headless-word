.PHONY: lint typecheck format format-check test check

lint:
	uv run ruff check src tests

typecheck:
	uv run ty check src

format:
	uv run ruff format src tests

format-check:
	uv run ruff format --check src tests

test:
	uv run pytest

check: lint typecheck format-check
