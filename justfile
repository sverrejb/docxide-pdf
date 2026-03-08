image := "docxside-pdf"

# Build the dev container
build:
    podman build -t {{image}} .

# Start an interactive shell in the dev container
dev:
    podman run -it --rm \
        -v "$(pwd)":/workspace:Z \
        -v claude-config:/home/dev/.claude:U \
        -v cargo-registry:/home/dev/.cargo/registry \
        -v cargo-target:/home/dev/.cargo-target \
        -v sccache:/home/dev/.cache/sccache \
        {{image}} bash

# Start claude directly in the dev container
claude *args:
    podman run -it --rm \
        -v "$(pwd)":/workspace:Z \
        -v claude-config:/home/dev/.claude:U \
        -v cargo-registry:/home/dev/.cargo/registry \
        -v cargo-target:/home/dev/.cargo-target \
        -v sccache:/home/dev/.cache/sccache \
        {{image}} claude {{args}}
