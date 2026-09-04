#!/usr/bin/env sh
# Publish comparison/ (minus its work/ cache, per comparison/.gitignore) to GitHub Pages as a single orphan commit.
#
# The branch carries no history: every run replaces it with one commit holding the current
# site, so nothing accumulates. Git still only uploads blobs the remote lacks, so a re-deploy
# after a re-score sends just the changed pages.
#
#   tools/deploy_comparison.sh                      # push to origin, branch gh-pages
#   tools/deploy_comparison.sh origin comparison    # other branch name
#   tools/deploy_comparison.sh git@github.com:USER/docxide-pdf-comparison.git   # separate repo
#   DRY_RUN=1 tools/deploy_comparison.sh            # build the commit, print, do not push
#
# First time only, enable Pages for the target (branch must exist first):
#   gh api -X POST repos/OWNER/REPO/pages -f build_type=legacy -f 'source[branch]=gh-pages' -f 'source[path]=/'
set -eu

ROOT=$(cd "$(dirname "$0")/.." && pwd)
SITE="$ROOT/comparison"
REMOTE=${1:-origin}
BRANCH=${2:-gh-pages}

[ -f "$SITE/index.html" ] || {
    echo "no $SITE/index.html; build it first: python3 tools/engine_compare.py" >&2
    exit 1
}

# A throwaway .git elsewhere; comparison/ itself is the work tree, nothing is copied.
# `git add -A` honours comparison/.gitignore, which excludes work/ (PDFs, PNGs, LO profiles).
TMP=$(mktemp -d)
trap 'rm -rf "$TMP"' EXIT
g() { git --git-dir="$TMP/.git" --work-tree="$SITE" "$@"; }

g init -q
touch "$SITE/.nojekyll"   # stop GitHub running Jekyll over 9k files
g add -A
g -c user.name="engine_compare" -c user.email="engine_compare@localhost" \
    commit -q -m "engine comparison $(date +%Y-%m-%d) ($(g ls-files | grep -c 'page_' | tr -d ' ') pages)"

URL=$(git -C "$ROOT" remote get-url "$REMOTE" 2>/dev/null || echo "$REMOTE")
SIZE=$(du -sh -I work "$SITE" 2>/dev/null | cut -f1 || du -sh --exclude=work "$SITE" | cut -f1)   # BSD du, then GNU du
echo "commit $(g rev-parse --short HEAD): $(g ls-files | wc -l | tr -d ' ') files, $SIZE -> $URL $BRANCH"

if [ -n "${DRY_RUN:-}" ]; then
    echo "dry run, not pushing"
    exit 0
fi
g push --force "$URL" "HEAD:refs/heads/$BRANCH"
echo "pushed. Site: https://$(echo "$URL" | sed -E 's#.*github.com[:/]([^/]+)/([^/.]+).*#\1.github.io/\2#')/"
