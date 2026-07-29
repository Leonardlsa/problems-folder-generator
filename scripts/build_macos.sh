#!/usr/bin/env bash
set -euo pipefail

EXPECTED_ARCH="${1:-}"
PYTHON_BIN="${PYTHON:-python3}"

case "$(uname -m)" in
  arm64|aarch64)
    ARCH="arm64"
    ;;
  x86_64)
    ARCH="x86_64"
    ;;
  *)
    echo "Unsupported macOS architecture: $(uname -m)" >&2
    exit 1
    ;;
esac

if [[ -n "$EXPECTED_ARCH" && "$EXPECTED_ARCH" != "$ARCH" ]]; then
  echo "Runner architecture mismatch: expected $EXPECTED_ARCH, got $ARCH" >&2
  exit 1
fi

rm -rf build dist

"$PYTHON_BIN" -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "错题集生成器" \
  --osx-bundle-identifier "io.github.leonardlsa.problems-folder-generator" \
  --add-data "ui/icons:ui/icons" \
  main.py

APP_PATH="dist/错题集生成器.app"
DMG_NAME="problems-folder-generator-macos-${ARCH}.dmg"
DMG_PATH="dist/${DMG_NAME}"

if [[ ! -d "$APP_PATH" ]]; then
  echo "macOS application bundle was not created: $APP_PATH" >&2
  exit 1
fi

# Ad-hoc signing keeps the bundle internally consistent. It is not notarization.
/usr/bin/codesign --force --deep --sign - "$APP_PATH"
/usr/bin/codesign --verify --deep --strict "$APP_PATH"

STAGING_DIR="$(mktemp -d "${TMPDIR:-/tmp}/pfg-dmg.XXXXXX")"
cleanup() {
  rm -rf "$STAGING_DIR"
}
trap cleanup EXIT

cp -R "$APP_PATH" "$STAGING_DIR/"
ln -s /Applications "$STAGING_DIR/Applications"

/usr/bin/hdiutil create \
  -volname "错题集生成器" \
  -srcfolder "$STAGING_DIR" \
  -ov \
  -format UDZO \
  "$DMG_PATH"

if [[ ! -s "$DMG_PATH" ]]; then
  echo "DMG was not created: $DMG_PATH" >&2
  exit 1
fi

if [[ "$(stat -f%z "$DMG_PATH")" -lt 1000000 ]]; then
  echo "DMG is unexpectedly small" >&2
  exit 1
fi

(
  cd dist
  shasum -a 256 "$DMG_NAME" > "$DMG_NAME.sha256"
  shasum -a 256 -c "$DMG_NAME.sha256"
)
echo "Created $DMG_PATH"
