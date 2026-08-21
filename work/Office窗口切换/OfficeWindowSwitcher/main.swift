import AppKit
import ApplicationServices

private enum MenuLayout {
    static let width: CGFloat = 220
    static let windowTitleAvailableWidth: CGFloat = 156
    static let windowTitleEllipsisCenter: CGFloat = 78
}

final class AppDelegate: NSObject, NSApplicationDelegate, NSMenuDelegate {
    private let statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.squareLength)
    private let menu = NSMenu()
    private let officeApps: [(name: String, bundleIdentifier: String)] = [
        ("Word", "com.microsoft.Word"),
        ("Excel", "com.microsoft.Excel"),
        ("PowerPoint", "com.microsoft.Powerpoint")
    ]
    private var windowOrder: [String: Int] = [:]
    private var nextWindowOrder = 0
    func applicationDidFinishLaunching(_ notification: Notification) {
        NSApp.setActivationPolicy(.accessory)
        statusItem.button?.image = NSImage(systemSymbolName: "doc.on.doc", accessibilityDescription: "Mac Office Tools")
        statusItem.button?.image?.isTemplate = true
        statusItem.menu = menu
        menu.minimumWidth = MenuLayout.width
        menu.delegate = self
        requestAccessibilityPermissionIfNeeded()
    }

    func menuNeedsUpdate(_ menu: NSMenu) {
        rebuildMenu()
    }

    private func rebuildMenu() {
        menu.removeAllItems()

        guard AXIsProcessTrusted() else {
            let permissionItem = NSMenuItem(
                title: "请在系统提示中允许“辅助功能”权限",
                action: nil,
                keyEquivalent: ""
            )
            permissionItem.isEnabled = false
            menu.addItem(permissionItem)
            menu.addItem(withTitle: "退出", action: #selector(quit), keyEquivalent: "q").target = self
            return
        }

        var hasVisibleGroup = false
        for officeApp in officeApps {
            let applications = NSWorkspace.shared.runningApplications.filter {
                $0.bundleIdentifier == officeApp.bundleIdentifier
            }
            let windows = applications.flatMap(windows(for:))
            guard !windows.isEmpty else { continue }

            if hasVisibleGroup {
                menu.addItem(.separator())
            }
            addWindowGroup(
                name: officeApp.name,
                bundleIdentifier: officeApp.bundleIdentifier,
                windows: orderedWindows(windows)
            )
            hasVisibleGroup = true
        }

        if !hasVisibleGroup {
            let emptyItem = NSMenuItem(title: "没有打开的 Office 文件", action: nil, keyEquivalent: "")
            emptyItem.isEnabled = false
            menu.addItem(emptyItem)
        }

        menu.addItem(.separator())
        let newItem = NSMenuItem(title: "新建", action: nil, keyEquivalent: "")
        newItem.submenu = newDocumentMenu()
        menu.addItem(newItem)
        let refreshItem = menu.addItem(withTitle: "刷新", action: #selector(refresh), keyEquivalent: "r")
        refreshItem.target = self
        let quitItem = menu.addItem(withTitle: "退出", action: #selector(quit), keyEquivalent: "q")
        quitItem.target = self
    }

    private func addWindowGroup(
        name: String,
        bundleIdentifier: String,
        windows: [OfficeWindow]
    ) {
        if #available(macOS 14.0, *) {
            menu.addItem(.sectionHeader(title: name))
        } else {
            let header = NSMenuItem(title: name, action: nil, keyEquivalent: "")
            header.isEnabled = false
            menu.addItem(header)
        }

        for window in windows {
            let displayedTitle = displayTitle(for: window.title, bundleIdentifier: bundleIdentifier)
            let windowItem = NSMenuItem(title: displayedTitle, action: nil, keyEquivalent: "")
            if window.isFocused {
                windowItem.attributedTitle = NSAttributedString(
                    string: displayedTitle,
                    attributes: [
                        .font: NSFont.menuFont(ofSize: 0),
                        .foregroundColor: NSColor.secondaryLabelColor
                    ]
                )
            }
            windowItem.submenu = windowActionsMenu(for: window)
            menu.addItem(windowItem)
        }
    }

    private func newDocumentMenu() -> NSMenu {
        let newMenu = NSMenu()
        for officeApp in officeApps {
            let item = NSMenuItem(title: officeApp.name, action: #selector(createNewDocument(_:)), keyEquivalent: "")
            item.target = self
            item.representedObject = officeApp.bundleIdentifier
            newMenu.addItem(item)
        }
        return newMenu
    }

    private func windowActionsMenu(for window: OfficeWindow) -> NSMenu {
        let actionsMenu = NSMenu()

        let openItem = NSMenuItem(
            title: "打开",
            action: window.isFocused ? nil : #selector(activateWindow(_:)),
            keyEquivalent: ""
        )
        openItem.isEnabled = !window.isFocused
        if !window.isFocused {
            openItem.target = self
            openItem.representedObject = window
        }
        openItem.image = NSImage(systemSymbolName: "arrow.up.forward.app", accessibilityDescription: "打开")
        openItem.image?.isTemplate = true
        actionsMenu.addItem(openItem)

        let revealItem = NSMenuItem(title: "在 Finder 中显示", action: #selector(revealWindowFile(_:)), keyEquivalent: "")
        revealItem.target = self
        revealItem.representedObject = window
        revealItem.image = NSImage(systemSymbolName: "folder", accessibilityDescription: "在 Finder 中显示")
        revealItem.image?.isTemplate = true
        actionsMenu.addItem(revealItem)

        let copyItem = NSMenuItem(title: "复制", action: #selector(copyWindowFile(_:)), keyEquivalent: "")
        copyItem.target = self
        copyItem.representedObject = window
        copyItem.image = NSImage(systemSymbolName: "doc.on.doc", accessibilityDescription: "复制")
        copyItem.image?.isTemplate = true
        actionsMenu.addItem(copyItem)

        let closeItem = NSMenuItem(title: "关闭", action: #selector(closeWindow(_:)), keyEquivalent: "")
        closeItem.target = self
        closeItem.representedObject = window
        closeItem.image = NSImage(systemSymbolName: "xmark.circle", accessibilityDescription: "关闭")
        closeItem.image?.isTemplate = true
        actionsMenu.addItem(closeItem)

        return actionsMenu
    }

    private func windows(for application: NSRunningApplication) -> [OfficeWindow] {
        let appElement = AXUIElementCreateApplication(application.processIdentifier)
        var value: CFTypeRef?
        guard AXUIElementCopyAttributeValue(appElement, kAXWindowsAttribute as CFString, &value) == .success,
              let axWindows = value as? [AXUIElement] else {
            return []
        }

        var focusedWindow: AXUIElement?
        if NSWorkspace.shared.frontmostApplication?.processIdentifier == application.processIdentifier {
            var focusedValue: CFTypeRef?
            if AXUIElementCopyAttributeValue(appElement, kAXFocusedWindowAttribute as CFString, &focusedValue) == .success {
                focusedWindow = focusedValue as! AXUIElement
            }
        }

        return axWindows.compactMap { window in
            guard let title = stringAttribute(kAXTitleAttribute as CFString, from: window),
                  !title.trimmingCharacters(in: .whitespacesAndNewlines).isEmpty else {
                return nil
            }
            return OfficeWindow(
                pid: application.processIdentifier,
                element: window,
                title: title,
                isFocused: focusedWindow.map { CFEqual($0, window) } ?? false
            )
        }
    }

    private func orderedWindows(_ windows: [OfficeWindow]) -> [OfficeWindow] {
        for window in windows where windowOrder[window.orderKey] == nil {
            windowOrder[window.orderKey] = nextWindowOrder
            nextWindowOrder += 1
        }

        return windows.sorted { (windowOrder[$0.orderKey] ?? 0) < (windowOrder[$1.orderKey] ?? 0) }
    }

    private func displayTitle(for title: String, bundleIdentifier: String) -> String {
        let titleWithoutCompatibilityMode = titleWithoutWordCompatibilityMode(title, bundleIdentifier: bundleIdentifier)
        let titleForDisplay = (titleWithoutCompatibilityMode as NSString).deletingPathExtension
        let attributes: [NSAttributedString.Key: Any] = [.font: NSFont.menuFont(ofSize: 0)]
        let availableWidth = MenuLayout.windowTitleAvailableWidth
        guard (titleForDisplay as NSString).size(withAttributes: attributes).width > availableWidth else {
            return titleForDisplay
        }

        let ellipsis = "…"
        let ellipsisWidth = (ellipsis as NSString).size(withAttributes: attributes).width
        let characters = Array(titleForDisplay)
        var bestTitle: String?
        var bestWidth: CGFloat = 0
        var smallestCenterDifference = CGFloat.greatestFiniteMagnitude

        for leadingCount in 1..<characters.count {
            for trailingCount in 1..<(characters.count - leadingCount) {
                let leadingTitle = String(characters.prefix(leadingCount))
                let trailingTitle = String(characters.suffix(trailingCount))
                let candidate = leadingTitle + ellipsis + trailingTitle
                let candidateWidth = (candidate as NSString).size(withAttributes: attributes).width
                guard candidateWidth <= availableWidth else { continue }

                let ellipsisCenter = (leadingTitle as NSString).size(withAttributes: attributes).width + ellipsisWidth / 2
                let centerDifference = abs(ellipsisCenter - MenuLayout.windowTitleEllipsisCenter)
                if centerDifference < smallestCenterDifference ||
                    (centerDifference == smallestCenterDifference && candidateWidth > bestWidth) {
                    bestTitle = candidate
                    bestWidth = candidateWidth
                    smallestCenterDifference = centerDifference
                }
            }
        }

        return bestTitle ?? titleForDisplay
    }

    private func titleWithoutWordCompatibilityMode(_ title: String, bundleIdentifier: String) -> String {
        guard bundleIdentifier == "com.microsoft.Word" else { return title }

        let trimmedTitle = title.trimmingCharacters(in: .whitespacesAndNewlines)
        let compatibilityMode = "兼容性模式"
        guard trimmedTitle.hasSuffix(compatibilityMode) else { return title }

        let markerStart = trimmedTitle.index(trimmedTitle.endIndex, offsetBy: -compatibilityMode.count)
        let prefix = trimmedTitle[..<markerStart]
        let separatorCharacters = CharacterSet.whitespacesAndNewlines
            .union(CharacterSet(charactersIn: "-‐‑‒–—―－\u{00A0}"))
        guard let finalPrefixScalar = prefix.unicodeScalars.last,
              separatorCharacters.contains(finalPrefixScalar) else {
            return title
        }

        let fileName = String(prefix).trimmingCharacters(in: separatorCharacters)
        return fileName.isEmpty ? title : fileName
    }

    private func stringAttribute(_ attribute: CFString, from element: AXUIElement) -> String? {
        var value: CFTypeRef?
        guard AXUIElementCopyAttributeValue(element, attribute, &value) == .success else {
            return nil
        }
        return value as? String
    }

    private func requestAccessibilityPermissionIfNeeded() {
        guard !AXIsProcessTrusted() else { return }
        let options = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary
        AXIsProcessTrustedWithOptions(options)
    }

    @objc private func activateWindow(_ sender: NSMenuItem) {
        guard let reference = sender.representedObject as? OfficeWindow else { return }
        bringWindowToFront(reference)
    }

    private func bringWindowToFront(_ reference: OfficeWindow) {
        if let application = NSRunningApplication(processIdentifier: reference.pid) {
            application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        }
        AXUIElementSetAttributeValue(reference.element, kAXMinimizedAttribute as CFString, kCFBooleanFalse)
        AXUIElementPerformAction(reference.element, kAXRaiseAction as CFString)
    }

    @objc private func createNewDocument(_ sender: NSMenuItem) {
        guard let bundleIdentifier = sender.representedObject as? String else {
            return
        }
        let applicationName = sender.title

        if let application = NSWorkspace.shared.runningApplications.first(where: {
            $0.bundleIdentifier == bundleIdentifier
        }) {
            createNewDocument(in: application, after: 0.4)
            return
        }

        guard let applicationURL = NSWorkspace.shared.urlForApplication(withBundleIdentifier: bundleIdentifier) else {
            StatusToast.show(title: "未找到 \(applicationName)", symbolName: "exclamationmark.triangle", isError: true)
            return
        }

        let configuration = NSWorkspace.OpenConfiguration()
        configuration.activates = true
        NSWorkspace.shared.openApplication(at: applicationURL, configuration: configuration) { [weak self] application, _ in
            DispatchQueue.main.async {
                guard let application else {
                    StatusToast.show(title: "未能启动 \(applicationName)", symbolName: "exclamationmark.triangle", isError: true)
                    return
                }
                self?.createNewDocument(in: application, after: 1.0)
            }
        }
    }

    private func createNewDocument(in application: NSRunningApplication, after delay: TimeInterval) {
        application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        DispatchQueue.main.asyncAfter(deadline: .now() + delay) {
            let source = CGEventSource(stateID: .hidSystemState)
            let keyCode = CGKeyCode(45) // N
            let keyDown = CGEvent(keyboardEventSource: source, virtualKey: keyCode, keyDown: true)
            let keyUp = CGEvent(keyboardEventSource: source, virtualKey: keyCode, keyDown: false)
            keyDown?.flags = .maskCommand
            keyUp?.flags = .maskCommand
            keyDown?.postToPid(application.processIdentifier)
            keyUp?.postToPid(application.processIdentifier)
        }
        DispatchQueue.main.asyncAfter(deadline: .now() + delay + 0.3) { [weak self] in
            self?.rebuildMenu()
        }
    }

    @objc private func closeWindow(_ sender: NSMenuItem) {
        guard let reference = sender.representedObject as? OfficeWindow else { return }
        if let application = NSRunningApplication(processIdentifier: reference.pid) {
            application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        }
        var closeButtonValue: CFTypeRef?
        if AXUIElementCopyAttributeValue(reference.element, kAXCloseButtonAttribute as CFString, &closeButtonValue) == .success {
            AXUIElementPerformAction(closeButtonValue as! AXUIElement, kAXPressAction as CFString)
        }
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.2) { [weak self] in
            self?.rebuildMenu()
        }
    }

    @objc private func refresh() {
        rebuildMenu()
    }

    @objc private func copyWindowFile(_ sender: NSMenuItem) {
        guard let reference = sender.representedObject as? OfficeWindow else { return }
        if let fileURL = accessibleFileURL(for: reference.element) {
            copyToPasteboard(fileURL)
            return
        }

        bringWindowToFront(reference)
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.25) { [weak self] in
            self?.copyActiveDocument()
        }
    }

    @objc private func revealWindowFile(_ sender: NSMenuItem) {
        guard let reference = sender.representedObject as? OfficeWindow else { return }
        if let fileURL = accessibleFileURL(for: reference.element) {
            revealInFinder(fileURL)
            return
        }

        bringWindowToFront(reference)
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.25) { [weak self] in
            self?.revealActiveDocument()
        }
    }

    fileprivate func copyActiveDocument() {
        guard let application = NSWorkspace.shared.frontmostApplication else { return }

        DispatchQueue.global(qos: .userInitiated).async {
            guard let fileURL = self.activeDocumentURL(for: application, unsavedMessage: "请保存后再复制") else { return }
            self.copyToPasteboard(fileURL)
        }
    }

    private func revealActiveDocument() {
        guard let application = NSWorkspace.shared.frontmostApplication else { return }

        DispatchQueue.global(qos: .userInitiated).async {
            guard let fileURL = self.activeDocumentURL(for: application, unsavedMessage: "请先保存该文件") else { return }
            self.revealInFinder(fileURL)
        }
    }

    private func copyToPasteboard(_ fileURL: URL) {
        DispatchQueue.main.async {
            let pasteboard = NSPasteboard.general
            pasteboard.clearContents()
            guard pasteboard.writeObjects([fileURL as NSURL]) else {
                StatusToast.show(title: "未能复制该文件", symbolName: "exclamationmark.triangle", isError: true)
                return
            }
            StatusToast.show(title: "已复制：\(fileURL.lastPathComponent)", symbolName: "checkmark.circle.fill")
        }
    }

    private func revealInFinder(_ fileURL: URL) {
        DispatchQueue.main.async {
            NSWorkspace.shared.activateFileViewerSelecting([fileURL])
        }
    }

    private func activeDocumentURL(for application: NSRunningApplication, unsavedMessage: String) -> URL? {
        switch application.bundleIdentifier {
        case "com.microsoft.Excel":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Excel"
                    set documentPath to full name of active workbook
                end tell
                return documentPath
                """,
                unsavedMessage: unsavedMessage
            )
        case "com.microsoft.Word":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Word"
                    set documentPath to full name of active document
                end tell
                return documentPath
                """,
                unsavedMessage: unsavedMessage
            )
        case "com.microsoft.Powerpoint":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Powerpoint"
                    set documentPath to full name of active presentation
                end tell
                return documentPath
                """,
                unsavedMessage: unsavedMessage
            )
        default:
            return nil
        }
    }

    private func officeDocumentURL(source: String, unsavedMessage: String) -> URL? {
        var error: NSDictionary?
        guard let result = NSAppleScript(source: source)?.executeAndReturnError(&error), error == nil else {
            DispatchQueue.main.async {
                StatusToast.show(title: unsavedMessage, symbolName: "exclamationmark.triangle", isError: true)
            }
            return nil
        }

        let path = result.stringValue ?? ""
        guard let fileURL = officeFileURL(for: path),
              FileManager.default.fileExists(atPath: fileURL.path) else {
            DispatchQueue.main.async {
                StatusToast.show(title: unsavedMessage, symbolName: "exclamationmark.triangle", isError: true)
            }
            return nil
        }
        return fileURL
    }

    private func officeFileURL(for path: String) -> URL? {
        guard !path.isEmpty else { return nil }
        if path.hasPrefix("/") {
            return URL(fileURLWithPath: path)
        }
        return CFURLCreateWithFileSystemPath(
            kCFAllocatorDefault,
            path as CFString,
            CFURLPathStyle(rawValue: 1)!,
            false
        ) as URL?
    }

    private func accessibleFileURL(for element: AXUIElement) -> URL? {
        guard let documentPath = stringAttribute(kAXDocumentAttribute as CFString, from: element) else {
            return nil
        }
        let fileURL = URL(string: documentPath)?.isFileURL == true
            ? URL(string: documentPath)!
            : URL(fileURLWithPath: documentPath)
        return FileManager.default.fileExists(atPath: fileURL.path) ? fileURL : nil
    }

    @objc private func quit() {
        NSApp.terminate(nil)
    }
}

private final class StatusToast {
    private static var panel: NSPanel?
    private static var dismissWorkItem: DispatchWorkItem?

    static func show(title: String, symbolName: String, isError: Bool = false) {
        dismissWorkItem?.cancel()
        panel?.orderOut(nil)

        let font = NSFont.systemFont(ofSize: 13, weight: .medium)
        let height: CGFloat = 36
        let horizontalPadding: CGFloat = 14
        let iconSize: CGFloat = 16
        let iconTextSpacing: CGFloat = 6
        let label = NSTextField(labelWithString: title)
        label.font = font
        let textWidth = ceil(label.intrinsicContentSize.width) + 6
        let width = horizontalPadding * 2 + iconSize + iconTextSpacing + textWidth

        let panel = NSPanel(
            contentRect: NSRect(x: 0, y: 0, width: width, height: height),
            styleMask: [.borderless, .nonactivatingPanel],
            backing: .buffered,
            defer: false
        )
        panel.isOpaque = false
        panel.backgroundColor = .clear
        panel.hasShadow = true
        panel.level = .statusBar
        panel.collectionBehavior = [.canJoinAllSpaces, .transient, .fullScreenAuxiliary]

        let background = NSVisualEffectView(frame: panel.contentView!.bounds)
        background.autoresizingMask = [.width, .height]
        background.material = .hudWindow
        background.state = .active
        background.wantsLayer = true
        background.layer?.cornerRadius = 12
        background.layer?.masksToBounds = true

        let imageView = NSImageView(frame: NSRect(
            x: horizontalPadding,
            y: (height - iconSize) / 2,
            width: iconSize,
            height: iconSize
        ))
        imageView.image = NSImage(systemSymbolName: symbolName, accessibilityDescription: nil)
        imageView.contentTintColor = isError ? .systemOrange : .systemGreen
        background.addSubview(imageView)

        label.frame = NSRect(
            x: horizontalPadding + iconSize + iconTextSpacing,
            y: (height - 18) / 2,
            width: textWidth,
            height: 18
        )
        label.lineBreakMode = .byClipping
        background.addSubview(label)
        panel.contentView = background

        let screen = NSScreen.screens.first(where: { $0.frame.contains(NSEvent.mouseLocation) }) ?? NSScreen.main
        if let screen {
            let visibleFrame = screen.visibleFrame
            panel.setFrameOrigin(NSPoint(x: visibleFrame.midX - width / 2, y: visibleFrame.minY + 72))
        }
        panel.orderFrontRegardless()
        self.panel = panel

        let workItem = DispatchWorkItem {
            panel.orderOut(nil)
        }
        dismissWorkItem = workItem
        DispatchQueue.main.asyncAfter(deadline: .now() + 1.8, execute: workItem)
    }
}

private final class OfficeWindow: NSObject {
    let pid: pid_t
    let element: AXUIElement
    let title: String
    let isFocused: Bool
    var orderKey: String { "\(pid):\(title)" }

    init(pid: pid_t, element: AXUIElement, title: String, isFocused: Bool) {
        self.pid = pid
        self.element = element
        self.title = title
        self.isFocused = isFocused
    }
}

let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.run()
