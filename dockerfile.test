FROM python:3.13-slim-trixie

ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    libreoffice-script-provider-python \
    python3-uno \
    poppler-utils \
    && soffice --version \
    && rm -rf /var/lib/apt/lists/* /tmp/* /var/tmp/* \
    && apt-get clean

COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

WORKDIR /app

COPY pyproject.toml uv.lock README.md ./
COPY src/ ./src/
COPY tests/ ./tests/

RUN uv sync --frozen --group dev

CMD ["bash", "-lc", "soffice --version && uv run pytest -v"]
