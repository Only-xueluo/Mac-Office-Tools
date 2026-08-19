import AppKit
import ApplicationServices

private enum MenuLayout {
    static let width: CGFloat = 240
    static let trailingControlEdge: CGFloat = 226
}

final class AppDelegate: NSObject, NSApplicationDelegate, NSMenuDelegate {
    private let statusItem = NSStatusBar.system.statusItem(withLength: NSStatusItem.squareLength)
    private let menu = NSMenu()
    private let officeApps: [(name: String, bundleIdentifier: String, canCreate: Bool)] = [
        ("Word", "com.microsoft.Word", true),
        ("Excel", "com.microsoft.Excel", true),
        ("PowerPoint", "com.microsoft.Powerpoint", true)
    ]
    private var windowOrder: [String: Int] = [:]
    private var nextWindowOrder = 0
    private var menuIsVisible = false

    func applicationDidFinishLaunching(_ notification: Notification) {
        NSApp.setActivationPolicy(.accessory)
        statusItem.button?.image = NSImage(systemSymbolName: "doc.on.doc", accessibilityDescription: "Mac Office Tools")
        statusItem.button?.image?.isTemplate = true
        statusItem.menu = menu
        menu.delegate = self
        requestAccessibilityPermissionIfNeeded()
    }

    func menuNeedsUpdate(_ menu: NSMenu) {
        rebuildMenu()
    }

    func menuWillOpen(_ menu: NSMenu) {
        menuIsVisible = true
    }

    func menuDidClose(_ menu: NSMenu) {
        menuIsVisible = false
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
                canCreate: officeApp.canCreate,
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
        let refreshItem = menu.addItem(withTitle: "刷新", action: #selector(refresh), keyEquivalent: "r")
        refreshItem.target = self
        let quitItem = menu.addItem(withTitle: "退出", action: #selector(quit), keyEquivalent: "q")
        quitItem.target = self
    }

    private func addWindowGroup(
        name: String,
        bundleIdentifier: String,
        canCreate: Bool,
        windows: [OfficeWindow]
    ) {
        let header = NSMenuItem()
        header.view = GroupHeaderView(
            name: name,
            bundleIdentifier: bundleIdentifier,
            target: canCreate ? self : nil,
            action: canCreate ? #selector(createNewDocument(_:)) : nil
        )
        menu.addItem(header)

        for window in windows {
            let item = NSMenuItem()
            item.view = WindowItemView(
                window: window,
                target: self,
                activateAction: #selector(activateWindow(_:)),
                copyAction: #selector(copyWindowFile(_:)),
                closeAction: #selector(closeWindow(_:))
            )
            menu.addItem(item)
        }
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

    @objc private func activateWindow(_ sender: WindowButton) {
        let reference = sender.officeWindow

        menuIsVisible = false
        if let application = NSRunningApplication(processIdentifier: reference.pid) {
            application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        }
        AXUIElementSetAttributeValue(reference.element, kAXMinimizedAttribute as CFString, kCFBooleanFalse)
        AXUIElementPerformAction(reference.element, kAXRaiseAction as CFString)
        refreshMenuAfterWindowSwitch()
    }

    private func refreshMenuAfterWindowSwitch() {
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.2) { [weak self] in
            guard let self else { return }
            self.rebuildMenu()
            guard let button = self.statusItem.button else { return }
            self.menu.popUp(positioning: nil, at: NSPoint(x: 0, y: button.bounds.height), in: button)
        }
    }

    @objc private func createNewDocument(_ sender: NSButton) {
        guard let bundleIdentifier = sender.identifier?.rawValue,
              let application = NSWorkspace.shared.runningApplications.first(where: {
                  $0.bundleIdentifier == bundleIdentifier
              }) else {
            return
        }

        application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.4) {
            let source = CGEventSource(stateID: .hidSystemState)
            let keyCode = CGKeyCode(45) // N
            let keyDown = CGEvent(keyboardEventSource: source, virtualKey: keyCode, keyDown: true)
            let keyUp = CGEvent(keyboardEventSource: source, virtualKey: keyCode, keyDown: false)
            keyDown?.flags = .maskCommand
            keyUp?.flags = .maskCommand
            keyDown?.postToPid(application.processIdentifier)
            keyUp?.postToPid(application.processIdentifier)
        }
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.7) { [weak self] in
            self?.rebuildMenu()
        }
    }

    @objc private func closeWindow(_ sender: WindowCloseButton) {
        let reference = sender.officeWindow
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

    @objc private func copyWindowFile(_ sender: WindowCopyButton) {
        let reference = sender.officeWindow
        if let fileURL = accessibleFileURL(for: reference.element) {
            copyToPasteboard(fileURL)
            return
        }

        if let application = NSRunningApplication(processIdentifier: reference.pid) {
            application.activate(options: [.activateAllWindows, .activateIgnoringOtherApps])
        }
        AXUIElementSetAttributeValue(reference.element, kAXMinimizedAttribute as CFString, kCFBooleanFalse)
        AXUIElementPerformAction(reference.element, kAXRaiseAction as CFString)
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.25) { [weak self] in
            self?.copyActiveDocument()
        }
    }

    fileprivate func copyActiveDocument() {
        guard let application = NSWorkspace.shared.frontmostApplication else { return }

        DispatchQueue.global(qos: .userInitiated).async {
            guard let fileURL = self.activeDocumentURL(for: application) else { return }
            self.copyToPasteboard(fileURL)
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

    private func activeDocumentURL(for application: NSRunningApplication) -> URL? {
        switch application.bundleIdentifier {
        case "com.microsoft.Excel":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Excel"
                    set documentPath to full name of active workbook
                end tell
                return documentPath
                """,
                unsavedMessage: "请保存后再复制"
            )
        case "com.microsoft.Word":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Word"
                    set documentPath to full name of active document
                end tell
                return documentPath
                """,
                unsavedMessage: "请保存后再复制"
            )
        case "com.microsoft.Powerpoint":
            return officeDocumentURL(
                source: """
                tell application id "com.microsoft.Powerpoint"
                    set documentPath to full name of active presentation
                end tell
                return documentPath
                """,
                unsavedMessage: "请保存后再复制"
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

private final class GroupHeaderView: NSView {
    private var newButton: NSButton?

    init(name: String, bundleIdentifier: String, target: AnyObject?, action: Selector?) {
        super.init(frame: NSRect(x: 0, y: 0, width: MenuLayout.width, height: 25))

        let nameLabel = NSTextField(labelWithString: name)
        nameLabel.frame = NSRect(x: 14, y: 4, width: 123, height: 17)
        nameLabel.font = NSFont.menuFont(ofSize: 0)
        nameLabel.textColor = .secondaryLabelColor
        addSubview(nameLabel)

        if let target, let action {
            let newButton = NSButton(title: "+新建", target: target, action: action)
            newButton.frame = NSRect(x: MenuLayout.trailingControlEdge - 70, y: 2, width: 70, height: 21)
            newButton.isBordered = false
            newButton.font = NSFont.menuFont(ofSize: 0)
            newButton.alignment = .right
            newButton.identifier = NSUserInterfaceItemIdentifier(bundleIdentifier)
            addSubview(newButton)
            self.newButton = newButton
            updateNewButtonColor(isHovered: false)
        }

        addTrackingArea(NSTrackingArea(
            rect: bounds,
            options: [.activeAlways, .mouseEnteredAndExited, .mouseMoved],
            owner: self,
            userInfo: nil
        ))
    }

    required init?(coder: NSCoder) {
        fatalError("init(coder:) has not been implemented")
    }

    override func mouseMoved(with event: NSEvent) {
        guard let newButton else { return }
        updateNewButtonColor(isHovered: newButton.frame.contains(convert(event.locationInWindow, from: nil)))
    }

    override func mouseExited(with event: NSEvent) {
        updateNewButtonColor(isHovered: false)
    }

    private func updateNewButtonColor(isHovered: Bool) {
        newButton?.attributedTitle = NSAttributedString(
            string: "+新建",
            attributes: [
                .font: NSFont.menuFont(ofSize: 0),
                .foregroundColor: isHovered ? NSColor.labelColor : NSColor.secondaryLabelColor
            ]
        )
    }
}

private final class WindowItemView: NSView {
    private let officeWindow: OfficeWindow
    private let highlightView = NSView()
    private let windowButton: WindowButton
    private let copyButton: WindowCopyButton
    private let closeButton: WindowCloseButton

    init(
        window: OfficeWindow,
        target: AnyObject,
        activateAction: Selector,
        copyAction: Selector,
        closeAction: Selector
    ) {
        self.officeWindow = window
        self.windowButton = WindowButton(officeWindow: window, target: target, action: activateAction)
        self.copyButton = WindowCopyButton(officeWindow: window, target: target, action: copyAction)
        self.closeButton = WindowCloseButton(officeWindow: window, target: target, action: closeAction)
        super.init(frame: NSRect(x: 0, y: 0, width: MenuLayout.width, height: 23))

        highlightView.frame = NSRect(x: 4, y: 0, width: MenuLayout.width - 8, height: 23)
        highlightView.wantsLayer = true
        highlightView.layer?.backgroundColor = NSColor.selectedContentBackgroundColor.cgColor
        highlightView.layer?.cornerRadius = 4
        highlightView.isHidden = true
        addSubview(highlightView)

        windowButton.frame = NSRect(x: 20, y: 0, width: 153, height: 23)
        addSubview(windowButton)

        copyButton.frame = NSRect(x: MenuLayout.trailingControlEdge - 48, y: 1, width: 20, height: 21)
        addSubview(copyButton)

        closeButton.frame = NSRect(x: MenuLayout.trailingControlEdge - 22, y: 1, width: 22, height: 21)
        addSubview(closeButton)

        addTrackingArea(NSTrackingArea(
            rect: .zero,
            options: [.activeAlways, .inVisibleRect, .mouseEnteredAndExited],
            owner: self,
            userInfo: nil
        ))
    }

    required init?(coder: NSCoder) {
        fatalError("init(coder:) has not been implemented")
    }

    override func mouseEntered(with event: NSEvent) {
        updateHighlight(isHighlighted: true)
    }

    override func mouseExited(with event: NSEvent) {
        updateHighlight(isHighlighted: false)
    }

    private func updateHighlight(isHighlighted: Bool) {
        guard !officeWindow.isFocused else { return }
        highlightView.isHidden = !isHighlighted
        windowButton.setHighlighted(isHighlighted)
        copyButton.setHighlighted(isHighlighted)
        closeButton.setHighlighted(isHighlighted)
    }
}

private final class WindowButton: NSButton {
    let officeWindow: OfficeWindow
    private let titleFont: NSFont

    init(officeWindow: OfficeWindow, target: AnyObject, action: Selector) {
        self.officeWindow = officeWindow
        self.titleFont = officeWindow.isFocused
            ? NSFont.menuFont(ofSize: 0)
            : NSFont.systemFont(ofSize: NSFont.menuFont(ofSize: 0).pointSize, weight: .semibold)
        super.init(frame: .zero)
        title = officeWindow.title
        self.target = target
        self.action = action
        isBordered = false
        alignment = .left
        lineBreakMode = .byTruncatingTail
        font = titleFont
        isEnabled = !officeWindow.isFocused
        setHighlighted(false)
    }

    required init?(coder: NSCoder) {
        fatalError("init(coder:) has not been implemented")
    }

    func setHighlighted(_ isHighlighted: Bool) {
        let color: NSColor
        if officeWindow.isFocused {
            color = .secondaryLabelColor
        } else {
            color = isHighlighted ? .selectedMenuItemTextColor : .labelColor
        }
        attributedTitle = NSAttributedString(
            string: officeWindow.title,
            attributes: [
                .font: titleFont,
                .foregroundColor: color
            ]
        )
    }
}

private final class WindowCopyButton: NSButton {
    let officeWindow: OfficeWindow
    private var isRowHighlighted = false
    private var isPointerInside = false

    init(officeWindow: OfficeWindow, target: AnyObject, action: Selector) {
        self.officeWindow = officeWindow
        super.init(frame: .zero)
        self.target = target
        self.action = action
        isBordered = false
        image = NSImage(systemSymbolName: "doc.on.doc", accessibilityDescription: nil)?
            .withSymbolConfiguration(NSImage.SymbolConfiguration(pointSize: 11, weight: .regular))
        image?.isTemplate = true
        imagePosition = .imageOnly
        setAccessibilityLabel("复制 \(officeWindow.title)")
        updateTintColor()
        addTrackingArea(NSTrackingArea(
            rect: .zero,
            options: [.activeAlways, .inVisibleRect, .mouseEnteredAndExited],
            owner: self,
            userInfo: nil
        ))
    }

    required init?(coder: NSCoder) {
        fatalError("init(coder:) has not been implemented")
    }

    override func mouseEntered(with event: NSEvent) {
        isPointerInside = true
        updateTintColor()
    }

    override func mouseExited(with event: NSEvent) {
        isPointerInside = false
        updateTintColor()
    }

    func setHighlighted(_ isHighlighted: Bool) {
        isRowHighlighted = isHighlighted
        updateTintColor()
    }

    private func updateTintColor() {
        contentTintColor = isRowHighlighted
            ? .selectedMenuItemTextColor
            : (isPointerInside ? .labelColor : .secondaryLabelColor)
    }
}

private final class WindowCloseButton: NSButton {
    let officeWindow: OfficeWindow
    private var isRowHighlighted = false
    private var isPointerInside = false

    init(officeWindow: OfficeWindow, target: AnyObject, action: Selector) {
        self.officeWindow = officeWindow
        super.init(frame: .zero)
        self.target = target
        self.action = action
        isBordered = false
        image = NSImage(systemSymbolName: "xmark.circle", accessibilityDescription: nil)?
            .withSymbolConfiguration(NSImage.SymbolConfiguration(pointSize: 12, weight: .regular))
        image?.isTemplate = true
        imagePosition = .imageOnly
        setAccessibilityLabel("关闭 \(officeWindow.title)")
        updateTintColor()
        addTrackingArea(NSTrackingArea(
            rect: .zero,
            options: [.activeAlways, .inVisibleRect, .mouseEnteredAndExited],
            owner: self,
            userInfo: nil
        ))
    }

    required init?(coder: NSCoder) {
        fatalError("init(coder:) has not been implemented")
    }

    override func mouseEntered(with event: NSEvent) {
        isPointerInside = true
        updateTintColor()
    }

    override func mouseExited(with event: NSEvent) {
        isPointerInside = false
        updateTintColor()
    }

    func setHighlighted(_ isHighlighted: Bool) {
        isRowHighlighted = isHighlighted
        updateTintColor()
    }

    private func updateTintColor() {
        contentTintColor = isRowHighlighted
            ? .selectedMenuItemTextColor
            : (isPointerInside ? .labelColor : .secondaryLabelColor)
    }
}

let app = NSApplication.shared
let delegate = AppDelegate()
app.delegate = delegate
app.run()
