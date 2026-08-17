#!/bin/zsh
set -euo pipefail

script_dir="$(cd "$(dirname "$0")" && pwd)"
project_root="$(cd "$script_dir/../.." && pwd)"
source_file="$script_dir/OfficeWindowSwitcher/main.swift"
icon_source="$script_dir/OfficeWindowSwitcher/IconGenerator.swift"
asset_catalog_source="$script_dir/OfficeWindowSwitcher/Assets.xcassets"
plist_file="$script_dir/OfficeWindowSwitcher/Info.plist"
release_dir="$project_root/output/Mac Office Tools/v3.11"
app_path="$release_dir/Mac Office Tools.app"
sdk_path="/Library/Developer/CommandLineTools/SDKs/MacOSX15.4.sdk"
actool_path="/Volumes/Applications/Applications/Xcode.app/Contents/Developer/usr/bin/actool"

if [[ -e "$app_path" ]]; then
    print -u2 "目标版本已有交付物，为保护已有文件，构建已停止：$release_dir"
    exit 1
fi

if [[ ! -d "$sdk_path" ]]; then
    print -u2 "找不到可用的 macOS SDK：$sdk_path"
    exit 1
fi

if [[ ! -x "$actool_path" ]]; then
    print -u2 "找不到 Xcode 资产编译器：$actool_path"
    exit 1
fi

build_root="$(mktemp -d "$project_root/work/Office窗口切换/build.XXXXXX")"
asset_output="$(mktemp -d /private/tmp/OfficeWindowSwitcherAssets.XXXXXX)"
build_app="$build_root/Mac Office Tools.app"
asset_catalog="$build_root/Assets.xcassets"
icon_generator="$build_root/icon-generator"
mkdir -p "$build_app/Contents/MacOS" "$build_app/Contents/Resources"
cp "$plist_file" "$build_app/Contents/Info.plist"
cp -R "$asset_catalog_source" "$asset_catalog"
swiftc -sdk "$sdk_path" -target arm64-apple-macos13.0 -module-cache-path "$build_root/module-cache" \
    "$source_file" -o "$build_app/Contents/MacOS/Mac Office Tools"
swiftc -sdk "$sdk_path" -target arm64-apple-macos13.0 -module-cache-path "$build_root/icon-module-cache" \
    "$icon_source" -o "$icon_generator"
icon_specs=(
    "16:icon_16x16.png" "32:icon_16x16@2x.png"
    "32:icon_32x32.png" "64:icon_32x32@2x.png"
    "128:icon_128x128.png" "256:icon_128x128@2x.png"
    "256:icon_256x256.png" "512:icon_256x256@2x.png"
    "512:icon_512x512.png" "1024:icon_512x512@2x.png"
)
for spec in "${icon_specs[@]}"; do
    size="${spec%%:*}"
    filename="${spec#*:}"
    "$icon_generator" "$asset_catalog/AppIcon.appiconset/$filename" "$size"
done
"$actool_path" --compile "$asset_output" --platform macosx \
    --minimum-deployment-target 13.0 --app-icon AppIcon \
    --output-partial-info-plist "$asset_output/asset-info.plist" "$asset_catalog"
cp "$asset_output/Assets.car" "$asset_output/AppIcon.icns" "$build_app/Contents/Resources/"
codesign --force --sign - "$build_app"
mkdir -p "$release_dir"
ditto "$build_app" "$app_path"

print "构建完成：$app_path"
