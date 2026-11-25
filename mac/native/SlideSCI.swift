import AppKit
import Foundation
import PDFKit
import CoreImage
import ImageIO
import ApplicationServices
import AVFoundation
import Darwin
import UniformTypeIdentifiers

// JSON is carried on stdin/stdout. Credentials never enter process arguments.
func fail(_ text: String) -> Never { FileHandle.standardError.write(Data(text.utf8)); exit(1) }
func json(_ value: Any) { do { let bytes = try JSONSerialization.data(withJSONObject: value, options: [.fragmentsAllowed, .sortedKeys]); FileHandle.standardOutput.write(bytes); FileHandle.standardOutput.write(Data([10])) } catch { fail(error.localizedDescription) } }
func required(_ input: [String: Any], _ key: String) throws -> String { guard let value = input[key] as? String, !value.isEmpty else { throw NSError(domain: "SlideSCI", code: 1, userInfo: [NSLocalizedDescriptionKey: "Missing \(key)"]) }; return value }
func snapshot(_ board: NSPasteboard) -> [[String: Data]] { (board.pasteboardItems ?? []).map { item in var result: [String: Data] = [:]; for type in item.types { if let data = item.data(forType: type) { result[type.rawValue] = data } }; return result } }
func restore(_ values: [[String: Data]], _ board: NSPasteboard) { board.clearContents(); let items = values.map { value -> NSPasteboardItem in let item = NSPasteboardItem(); for (type, data) in value { item.setData(data, forType: NSPasteboard.PasteboardType(type)) }; return item }; board.writeObjects(items) }
func fileURLs(_ board: NSPasteboard) -> [URL] { (board.readObjects(forClasses: [NSURL.self], options: [.urlReadingFileURLsOnly: true]) as? [URL]) ?? [] }
func files(_ board: NSPasteboard) -> [[String: Any]] { fileURLs(board).map { ["path": $0.path, "name": $0.lastPathComponent] } }
func png(_ data: Data, contrast: Bool = false) throws -> Data {
    guard var image = CIImage(data: data) else { throw NSError(domain: "SlideSCI", code: 2, userInfo: [NSLocalizedDescriptionKey: "Cannot decode image"]) }
    let context = CIContext()
    if contrast {
        let extent = CIVector(cgRect: image.extent)
        var minimum = [Float](repeating: 0, count: 4), maximum = minimum
        let low = image.applyingFilter("CIAreaMinimum", parameters: [kCIInputExtentKey: extent])
        let high = image.applyingFilter("CIAreaMaximum", parameters: [kCIInputExtentKey: extent])
        context.render(low, toBitmap: &minimum, rowBytes: 16, bounds: CGRect(x: 0, y: 0, width: 1, height: 1), format: .RGBAf, colorSpace: CGColorSpaceCreateDeviceRGB())
        context.render(high, toBitmap: &maximum, rowBytes: 16, bounds: CGRect(x: 0, y: 0, width: 1, height: 1), format: .RGBAf, colorSpace: CGColorSpaceCreateDeviceRGB())
        let lo = minimum.prefix(3).min()!, hi = maximum.prefix(3).max()!
        if hi > lo { let scale = CGFloat(1 / (hi - lo)), bias = -CGFloat(lo) * scale; image = image.applyingFilter("CIColorMatrix", parameters: ["inputRVector": CIVector(x: scale, y: 0, z: 0, w: 0), "inputGVector": CIVector(x: 0, y: scale, z: 0, w: 0), "inputBVector": CIVector(x: 0, y: 0, z: scale, w: 0), "inputBiasVector": CIVector(x: bias, y: bias, z: bias, w: 0)]) }
    }
    guard let cg = context.createCGImage(image, from: image.extent) else { throw NSError(domain: "SlideSCI", code: 3, userInfo: [NSLocalizedDescriptionKey: "Cannot render image"]) }
    let rep = NSBitmapImageRep(cgImage: cg)
    guard let output = rep.representation(using: .png, properties: [:]) else { throw NSError(domain: "SlideSCI", code: 4, userInfo: [NSLocalizedDescriptionKey: "Cannot encode PNG"]) }; return output
}
func pngInfo(_ data: Data) -> (width: Int, height: Int)? { NSBitmapImageRep(data: data).map { ($0.pixelsWide, $0.pixelsHigh) } }
let supportedHosts: Set<String> = ["com.microsoft.Powerpoint", "com.kingsoft.wpsoffice.mac", "com.kingsoft.wpsoffice"]
let mediaExtensions: Set<String> = ["png", "jpg", "jpeg", "gif", "tif", "tiff", "bmp", "svg", "emf", "wmf", "mp4", "m4v", "mov", "webm", "avi", "wmv", "asf", "mpg", "mpeg", "mpe"]
let hostPasteboardPrefixes = ["com.microsoft.", "com.kingsoft.", "org.openxmlformats.", "com.apple.iwork."]
func isHostInternalPasteboard(_ pb: NSPasteboard) -> Bool {
    let types = pb.types?.map { $0.rawValue.lowercased() } ?? []
    return types.contains { raw in hostPasteboardPrefixes.contains { raw.hasPrefix($0) } }
}
// PowerPoint/WPS copies and ordinary screenshots include an 8-bit TIFF fallback; only scientific
// 16-bit or calibrated TIFF clipboard images need background normalization.
func needsClipboardTIFF(_ data: Data) -> Bool {
    guard let source = CGImageSourceCreateWithData(data as CFData, nil),
          let props = CGImageSourceCopyPropertiesAtIndex(source, 0, nil) as? [CFString: Any] else { return false }
    if let depth = props[kCGImagePropertyDepth] as? Int, depth > 8 { return true }
    if let tiff = props[kCGImagePropertyTIFFDictionary] as? [CFString: Any],
       let desc = tiff[kCGImagePropertyTIFFImageDescription] as? String {
        let lower = desc.lowercased()
        if lower.contains("imagej") || lower.contains("<ome") { return true }
    }
    return false
}
var mediaEventTap: CFMachPort?
var pendingRestores: [String: (items: [[String: Data]], count: Int)] = [:]
var ignoredPasteboardCounts: [String: Int] = [:]
func interceptTIFF(_ pb: NSPasteboard, reason: String) {
    guard let host = NSWorkspace.shared.frontmostApplication?.bundleIdentifier, supportedHosts.contains(host) else { return }
    let urls = fileURLs(pb), tiffURLs = urls.filter { ["tif", "tiff"].contains($0.pathExtension.lowercased()) }
    let rawTiff = urls.isEmpty && !isHostInternalPasteboard(pb) ? pb.data(forType: .tiff) : nil
    guard (urls.count == 1 && tiffURLs.count == 1) || (rawTiff != nil && needsClipboardTIFF(rawTiff!)) else { return }
    do {
        let data = try tiffURLs.first.map { try Data(contentsOf: $0) } ?? rawTiff!
        let converted = try png(data, contrast: true), saved = snapshot(pb), item = NSPasteboardItem()
        item.setData(converted, forType: .png)
        if reason == "drop" {
            let directory = FileManager.default.homeDirectoryForCurrentUser.appendingPathComponent("Library/Containers/\(host)/Data/tmp/SlideSCI-Mac")
            try FileManager.default.createDirectory(at: directory, withIntermediateDirectories: true, attributes: [.posixPermissions: 0o700])
            let destination = directory.appendingPathComponent("\(UUID().uuidString).png"); try converted.write(to: destination, options: .atomic)
            item.setString(destination.absoluteString, forType: .fileURL)
        }
        pb.clearContents(); pb.writeObjects([item]); let replacedCount = pb.changeCount
        ignoredPasteboardCounts[pb.name.rawValue] = replacedCount
        pendingRestores[pb.name.rawValue] = (saved, replacedCount)
        json(["type": "mediaSource", "reason": reason, "files": tiffURLs.map { ["path": $0.path, "name": $0.lastPathComponent] }, "original": data.base64EncodedString(), "mime": "image/tiff", "intercepted": true, "changeCount": replacedCount, "pasteboard": pb.name.rawValue, "host": host, "timestamp": Date().timeIntervalSince1970 * 1000])
        DispatchQueue.main.asyncAfter(deadline: .now() + 0.8) { if pb.changeCount == replacedCount { restore(saved, pb); ignoredPasteboardCounts[pb.name.rawValue] = pb.changeCount }; if pendingRestores[pb.name.rawValue]?.count == replacedCount { pendingRestores.removeValue(forKey: pb.name.rawValue) } }
    } catch { json(["type": "nativeError", "message": error.localizedDescription]) }
}
let inputData = FileHandle.standardInput.readDataToEndOfFile()
let input: [String: Any]
do { input = try JSONSerialization.jsonObject(with: inputData) as? [String: Any] ?? [:] } catch { fail("Invalid native request") }
do {
    let action = try required(input, "action")
    let board = NSPasteboard.general
    switch action {
    case "pickFiles":
        let app = NSApplication.shared; app.setActivationPolicy(.accessory); app.activate(ignoringOtherApps: true)
        let panel = NSOpenPanel(); panel.canChooseFiles = true; panel.canChooseDirectories = false; panel.allowsMultipleSelection = input["multiple"] as? Bool ?? true
        switch input["kind"] as? String { case "images": panel.allowedContentTypes = [.image] + ["emf", "wmf"].compactMap { UTType(filenameExtension: $0) }; case "video": panel.allowedContentTypes = [.movie, .video]; case "presentations":
            let types = ["pptx", "ppt", "potx"].compactMap { UTType(filenameExtension: $0) }
            panel.allowedContentTypes = types.isEmpty ? [.data] : types
        default: break }
        json(panel.runModal() == .OK ? panel.urls.map { ["path": $0.path, "name": $0.lastPathComponent] } : [])
    case "chooseFolder":
        let app = NSApplication.shared; app.setActivationPolicy(.accessory); app.activate(ignoringOtherApps: true)
        let panel = NSOpenPanel(); panel.canChooseFiles = false; panel.canChooseDirectories = true; panel.canCreateDirectories = true
        if panel.runModal() == .OK { json(panel.url!.path) } else { json(NSNull()) }
    case "reveal": NSWorkspace.shared.activateFileViewerSelecting([URL(fileURLWithPath: try required(input, "path"))]); json([:])
    case "fonts": json(NSFontManager.shared.availableFontFamilies.sorted())
    case "readClipboardFiles": json(files(board))
    case "snapshotClipboard": json(snapshot(board).map { $0.mapValues { $0.base64EncodedString() } })
    case "restoreClipboard":
        let values = (input["items"] as? [[String: String]] ?? []).map { $0.compactMapValues { Data(base64Encoded: $0) } }; restore(values, board); json(["changeCount": board.changeCount])
    case "writeFileClipboard":
        let item = NSPasteboardItem(); item.setString(URL(fileURLWithPath: try required(input, "path")).absoluteString, forType: .fileURL)
        board.clearContents(); board.writeObjects([item]); json(["changeCount": board.changeCount])
    case "writeClipboard":
        guard let data = Data(base64Encoded: try required(input, "data")) else { throw NSError(domain: "SlideSCI", code: 5, userInfo: [NSLocalizedDescriptionKey: "Invalid base64"]) }
        let mime = try required(input, "mime"); let item = NSPasteboardItem()
        if mime == "image/svg+xml" { item.setData(data, forType: NSPasteboard.PasteboardType("public.svg-image")); item.setString(String(data: data, encoding: .utf8) ?? "", forType: .string) }
        else if mime.hasPrefix("image/") { let converted = try png(data); item.setData(converted, forType: .png); if let rep = NSImage(data: converted)?.tiffRepresentation { item.setData(rep, forType: .tiff) } }
        else if mime.hasPrefix("text/") { item.setData(data, forType: mime == "text/html" ? .html : .string) }
        else { item.setData(data, forType: NSPasteboard.PasteboardType(UTType(mimeType: mime)?.identifier ?? mime)) }
        board.clearContents(); board.writeObjects([item]); json(["changeCount": board.changeCount])
    case "mediaInfo":
        let url = URL(fileURLWithPath: try required(input, "path")), asset = AVURLAsset(url: url)
        var width = 0.0, height = 0.0, rotation = 0.0
        if let track = asset.tracks(withMediaType: .video).first {
            let size = track.naturalSize.applying(track.preferredTransform)
            width = abs(Double(size.width)); height = abs(Double(size.height))
            rotation = atan2(Double(track.preferredTransform.b), Double(track.preferredTransform.a)) * 180 / .pi
            if rotation < 0 { rotation += 360 }
            if rotation >= 359.99 { rotation = 0 }
        }
        let seconds = CMTimeGetSeconds(asset.duration)
        json(["width": width, "height": height, "duration": seconds.isFinite ? seconds : 0, "rotation": rotation])
    case "rasterPDF":
        let path = try required(input, "path"); guard let pdf = PDFDocument(url: URL(fileURLWithPath: path)) else { fail("Cannot open exported PDF") }
        let indices = input["indices"] as? [Int] ?? Array(0..<pdf.pageCount), dpi = input["dpi"] as? Double ?? 300, format = input["format"] as? String ?? "png", maximum = input["maximum"] as? Double
        guard dpi > 0 && dpi <= 2400 else { fail("DPI must be between 1 and 2400") }
        guard maximum == nil || (maximum! >= 1 && maximum! <= 32768) else { fail("Maximum raster dimension must be between 1 and 32768") }
        var pages: [[String: Any]] = []
        for index in indices { guard let page = pdf.page(at: index) else { fail("PDF page index out of range") }; let pageRect = page.bounds(for: .mediaBox)
            var bounds = pageRect
            if let crop = input["bounds"] as? [String: Double], let left = crop["left"], let top = crop["top"], let width = crop["width"], let height = crop["height"] { guard width > 0 && height > 0 else { fail("Selection bounds must have positive size") }; bounds = CGRect(x: left, y: pageRect.maxY - top - height, width: width, height: height) }
            let scale = maximum.map { min($0 / Double(bounds.width), $0 / Double(bounds.height)) } ?? dpi / 72
            let limit = maximum.map { Int($0.rounded(.down)) }
            var width = Int(ceil(bounds.width * scale)), height = Int(ceil(bounds.height * scale))
            if let limit = limit { width = min(width, limit); height = min(height, limit) }
            guard width > 0 && height > 0 && width * height <= 300_000_000 else { fail("Requested raster is too large") }
            let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: width, pixelsHigh: height, bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true, isPlanar: false, colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0)!
            NSGraphicsContext.saveGraphicsState(); let context = NSGraphicsContext(bitmapImageRep: rep)!; NSGraphicsContext.current = context; context.cgContext.setFillColor(NSColor.white.cgColor); context.cgContext.fill(CGRect(x: 0, y: 0, width: CGFloat(width), height: CGFloat(height))); context.cgContext.scaleBy(x: scale, y: scale); context.cgContext.translateBy(x: -bounds.minX, y: -bounds.minY); page.draw(with: .mediaBox, to: context.cgContext); NSGraphicsContext.restoreGraphicsState(); rep.size = bounds.size
            guard let data = rep.representation(using: format == "jpg" ? .jpeg : format == "bmp" ? .bmp : .png, properties: [.compressionFactor: 0.95]) else { fail("Cannot encode exported page") }; pages.append(["index": index, "data": data.base64EncodedString(), "mime": format == "jpg" ? "image/jpeg" : format == "bmp" ? "image/bmp" : "image/png", "width": width, "height": height])
        }; json(pages)
    case "slicePDF":
        guard let pdf = PDFDocument(url: URL(fileURLWithPath: try required(input, "path"))) else { fail("Cannot open PDF") }
        let indices = input["indices"] as? [Int] ?? Array(0..<pdf.pageCount)
        let separate = input["separate"] as? Bool ?? false
        let combined = separate ? nil : PDFDocument()
        var pages: [[String: Any]] = []
        if separate { pages.reserveCapacity(indices.count) }
        for index in indices {
            guard let page = pdf.page(at: index)?.copy() as? PDFPage else { fail("PDF page index out of range") }
            if let crop = input["bounds"] as? [String: Double], let left = crop["left"], let top = crop["top"], let width = crop["width"], let height = crop["height"] {
                let box = CGRect(x: left, y: page.bounds(for: .mediaBox).maxY - top - height, width: width, height: height)
                page.setBounds(box, for: .mediaBox); page.setBounds(box, for: .cropBox)
            }
            let output = combined ?? PDFDocument()
            output.insert(page, at: output.pageCount)
            if separate {
                guard let data = output.dataRepresentation() else { fail("Cannot encode PDF") }
                pages.append(["index": index, "data": data.base64EncodedString(), "mime": "application/pdf"])
            }
        }
        if let output = combined {
            guard let data = output.dataRepresentation() else { fail("Cannot encode PDF") }
            json(["data": data.base64EncodedString(), "mime": "application/pdf"])
        } else { json(pages) }
    case "normalizeTIFF":
        let path = try required(input, "path"), output = try png(Data(contentsOf: URL(fileURLWithPath: path)), contrast: true), size = pngInfo(output)
        json(["data": output.base64EncodedString(), "mime": "image/png", "width": size?.width ?? 0, "height": size?.height ?? 0])
    case "watch":
        let hosts = supportedHosts
        var counts: [String: Int] = [:]
        func capture(_ pb: NSPasteboard, _ reason: String) {
            guard let host = NSWorkspace.shared.frontmostApplication?.bundleIdentifier, hosts.contains(host), counts[pb.name.rawValue] != pb.changeCount, ignoredPasteboardCounts[pb.name.rawValue] != pb.changeCount else { return }
            let count = pb.changeCount
            counts[pb.name.rawValue] = count; ignoredPasteboardCounts.removeValue(forKey: pb.name.rawValue)
            let entries = files(pb).filter { entry in guard let path = entry["path"] as? String else { return false }; return mediaExtensions.contains(URL(fileURLWithPath: path).pathExtension.lowercased()) }
            if !entries.isEmpty { json(["type": "mediaSource", "reason": reason, "files": entries, "changeCount": count, "pasteboard": pb.name.rawValue, "host": host, "timestamp": Date().timeIntervalSince1970 * 1000]) }
            if entries.isEmpty, !isHostInternalPasteboard(pb), let data = pb.data(forType: .tiff), needsClipboardTIFF(data) {
                json(["type": "clipboardImage", "reason": reason, "data": data.base64EncodedString(), "mime": "image/tiff", "changeCount": count, "pasteboard": pb.name.rawValue, "host": host, "timestamp": Date().timeIntervalSince1970 * 1000])
            }
        }
        if input["interceptTIFF"] as? Bool == true {
            guard AXIsProcessTrusted() else { fail("TIFF paste/drop interception requires Accessibility access for the native helper. Enable it explicitly in System Settings; SlideSCI never requests it automatically.") }
            let mask = (CGEventMask(1) << CGEventType.keyDown.rawValue) | (CGEventMask(1) << CGEventType.leftMouseUp.rawValue)
            mediaEventTap = CGEvent.tapCreate(tap: .cgSessionEventTap, place: .headInsertEventTap, options: .defaultTap, eventsOfInterest: mask, callback: { _, type, event, _ in
                if type == .tapDisabledByTimeout || type == .tapDisabledByUserInput { if let tap = mediaEventTap { CGEvent.tapEnable(tap: tap, enable: true) }; return Unmanaged.passUnretained(event) }
                if type == .keyDown && event.flags.contains(.maskCommand) && event.getIntegerValueField(.keyboardEventKeycode) == 9 { interceptTIFF(NSPasteboard.general, reason: "paste") }
                else if type == .leftMouseUp { interceptTIFF(NSPasteboard(name: .drag), reason: "drop") }
                return Unmanaged.passUnretained(event)
            }, userInfo: nil)
            guard let tap = mediaEventTap, let source = CFMachPortCreateRunLoopSource(kCFAllocatorDefault, tap, 0) else { fail("macOS denied the requested TIFF paste/drop event tap. Grant Accessibility/Input Monitoring access explicitly or start capture without interception.") }
            CFRunLoopAddSource(CFRunLoopGetMain(), source, .commonModes); CGEvent.tapEnable(tap: tap, enable: true)
        }
        signal(SIGTERM, SIG_IGN)
        let stopSignal = DispatchSource.makeSignalSource(signal: SIGTERM, queue: .main)
        func finish() -> Never { for (name, saved) in pendingRestores { let pb = NSPasteboard(name: NSPasteboard.Name(name)); if pb.changeCount == saved.count { restore(saved.items, pb) } }; exit(0) }
        // The watcher must not outlive the service: a killed or crashed parent re-parents it to launchd.
        let parent = getppid()
        stopSignal.setEventHandler { finish() }; stopSignal.resume()
        let timer = Timer.scheduledTimer(withTimeInterval: 0.2, repeats: true) { _ in if getppid() != parent { finish() }; capture(board, "clipboard"); capture(NSPasteboard(name: .drag), "drag") }
        let monitor = NSEvent.addGlobalMonitorForEvents(matching: [.leftMouseUp, .rightMouseUp]) { _ in if let id = NSWorkspace.shared.frontmostApplication?.bundleIdentifier, hosts.contains(id) { capture(NSPasteboard(name: .drag), "drop"); json(["type": "hostChanged", "reason": "mouseUp", "host": id]) } }
        let observer = NSWorkspace.shared.notificationCenter.addObserver(forName: NSWorkspace.didActivateApplicationNotification, object: nil, queue: .main) { _ in if let id = NSWorkspace.shared.frontmostApplication?.bundleIdentifier, hosts.contains(id) { json(["type": "hostChanged", "reason": "activation", "host": id]) } }
        _ = (timer, monitor, observer, stopSignal)
        json(["type": "ready", "mouseEvents": monitor != nil, "accessibilityTrusted": AXIsProcessTrusted()]); RunLoop.main.run()
    default: fail("Unknown native helper action")
    }
} catch { fail(error.localizedDescription) }
