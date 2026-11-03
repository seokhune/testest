# syntax=docker/dockerfile:1
FROM python:3.14-slim

COPY --from=ghcr.io/astral-sh/uv:0.11.17 /uv /uvx /bin/

# libreoffice-writer: HWP/HWPX -> DOCX conversion, fonts-nanum: Korean text rendering
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-writer fonts-nanum && \
    rm -rf /var/lib/apt/lists/*

# H2Orestart adds the Hancom (HWP 5.x / HWPX) import filters to LibreOffice
RUN python -c "import urllib.request; urllib.request.urlretrieve('https://github.com/ebandal/H2Orestart/releases/download/v0.8.0/H2Orestart.oxt', '/tmp/H2Orestart.oxt')" && \
    echo "f36c19fa6b7dff1cf34114d926bdc088e6bcfa7fdc93d67eaafd3f572be1d62c  /tmp/H2Orestart.oxt" | sha256sum -c - && \
    unopkg add --suppress-license /tmp/H2Orestart.oxt && \
    rm -f /tmp/H2Orestart.oxt

WORKDIR /app
ENV UV_PYTHON_PREFERENCE=only-system \
    UV_NO_PYTHON_DOWNLOADS=1 \
    UV_LINK_MODE=copy \
    UV_COMPILE_BYTECODE=1
COPY pyproject.toml ./
COPY src ./src
RUN uv sync --no-dev

ENV HEALTH_SERVER_HOST=127.0.0.1 HEALTH_SERVER_PORT=3001
HEALTHCHECK --interval=15s --timeout=5s --retries=5 --start-period=20s \
    CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:3001/health', timeout=5).status==200 else 1)"
CMD ["sh", "-c", "exec uv run --no-dev --no-sync python -m worker.entrypoints.worker"]
