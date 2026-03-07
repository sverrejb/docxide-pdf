FROM rust:1-bookworm

# System dependencies: mutool (mupdf-tools), Node.js, fonts, and build essentials
RUN apt-get update && apt-get install -y --no-install-recommends \
    mupdf-tools \
    curl \
    ca-certificates \
    gnupg \
    pkg-config \
    # Fonts that substitute for common Windows/Office fonts
    fonts-liberation \
    fonts-dejavu \
    fonts-noto-core \
    fonts-freefont-ttf \
    fontconfig \
    && rm -rf /var/lib/apt/lists/*

# Install Node.js 22 for Claude Code
RUN curl -fsSL https://deb.nodesource.com/setup_22.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

# Install Claude Code
RUN npm install -g @anthropic-ai/claude-code

# Create non-root user
RUN useradd -m -s /bin/bash claude
USER claude
WORKDIR /home/claude/workspace

ENTRYPOINT ["claude"]
