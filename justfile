image := "docxside-pdf"

# Build the dev container
build:
    podman build -t {{image}} .

# Start an interactive shell in the dev container
dev:
    podman run -it --rm \
        --security-opt seccomp=unconfined \
        -e TERM -e COLORTERM \
        -v "$(pwd)":/workspace:Z \
        -v claude-config:/home/dev/.claude:U \
        -v cargo-registry:/home/dev/.cargo/registry \
        -v cargo-target:/home/dev/.cargo-target \
        -v sccache:/home/dev/.cache/sccache \
        {{image}} bash

# Start claude directly in the dev container
claude *args:
    podman run -it --rm \
        --security-opt seccomp=unconfined \
        -e TERM -e COLORTERM \
        -v "$(pwd)":/workspace:Z \
        -v claude-config:/home/dev/.claude:U \
        -v ~/specs:/home/dev/specs:ro,Z \
        -v cargo-registry:/home/dev/.cargo/registry \
        -v cargo-target:/home/dev/.cargo-target \
        -v sccache:/home/dev/.cache/sccache \
        {{image}} claude {{args}}

# One-time setup: configure MCP server inside the container
setup-mcp:
    podman run -it --rm \
        --security-opt seccomp=unconfined \
        -v claude-config:/home/dev/.claude:U \
        {{image}} claude mcp add local-rag --scope user \
            --env BASE_DIR=/home/dev/specs \
            -- npx -y mcp-local-rag
