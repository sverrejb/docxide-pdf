FROM docker/sandbox-templates:claude-code

USER root

# System dependencies: mutool and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
    mupdf-tools \
    pkg-config \
    fonts-liberation \
    fonts-dejavu \
    fonts-noto-core \
    fonts-freefont-ttf \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

USER agent

# Rust toolchain (installed as agent so ~/.cargo is accessible)
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | \
    sh -s -- -y --default-toolchain stable
ENV PATH="/home/agent/.cargo/bin:${PATH}"
