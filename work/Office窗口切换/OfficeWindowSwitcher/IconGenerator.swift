import AppKit

guard CommandLine.arguments.count == 3,
      let pixelSize = Int(CommandLine.arguments[2]),
      pixelSize > 0 else {
    fputs("Usage: IconGenerator <output-png> <pixel-size>\n", stderr)
    exit(1)
}

let size = NSSize(width: pixelSize, height: pixelSize)
let canvas = NSImage(size: size)
canvas.lockFocus()

let scale = CGFloat(pixelSize) / 1024
let background = NSBezierPath(
    roundedRect: NSRect(x: 40 * scale, y: 40 * scale, width: 944 * scale, height: 944 * scale),
    xRadius: 220 * scale,
    yRadius: 220 * scale
)
NSColor.systemIndigo.setFill()
background.fill()

let configuration = NSImage.SymbolConfiguration(pointSize: 680 * scale, weight: .medium)
let symbol = NSImage(systemSymbolName: "doc.on.doc", accessibilityDescription: nil)!
    .withSymbolConfiguration(configuration)!
symbol.isTemplate = true
symbol.draw(in: NSRect(x: 172 * scale, y: 172 * scale, width: 680 * scale, height: 680 * scale))
canvas.unlockFocus()

let bitmap = NSBitmapImageRep(data: canvas.tiffRepresentation!)!
try bitmap.representation(using: .png, properties: [:])!.write(to: URL(fileURLWithPath: CommandLine.arguments[1]))
