import Foundation
import PDFKit

struct FormulaBounds: Codable {
    let pageIndex: Int
    let text: String
    let minX: Double
    let minY: Double
    let maxX: Double
    let maxY: Double
    let width: Double
    let height: Double
    let centerX: Double
    let centerY: Double
    let relationshipXs: [Double]
}

struct GeometryReport: Codable {
    let pageWidth: Double
    let pageHeight: Double
    let unnumbered: FormulaBounds
    let numbered: FormulaBounds
    let equationNumber: FormulaBounds
    let aligned: [FormulaBounds]
    let pageText: String
}

struct RasterComponent: Codable {
    let minX: Double
    let minY: Double
    let maxX: Double
    let maxY: Double
    let width: Double
    let height: Double
    let centerX: Double
    let centerY: Double
}

struct RasterBand: Codable {
    let minY: Double
    let maxY: Double
    let height: Double
    let centerY: Double
    let inkWidth: Double
    let components: [RasterComponent]
}

struct NumberOnlyReport: Codable {
    let pageWidth: Double
    let pageHeight: Double
    let equationNumber: FormulaBounds
    let rasterBands: [RasterBand]
    let rasterComponents: [RasterComponent]
    let pageText: String
}

struct RasterOnlyReport: Codable {
    let pageWidth: Double
    let pageHeight: Double
    let rasterBands: [RasterBand]
    let rasterComponents: [RasterComponent]
    let pageText: String
}

enum GeometryError: Error, CustomStringConvertible {
    case usage
    case unreadablePdf(String)
    case missingToken(String)
    case missingEquationNumber

    var description: String {
        switch self {
        case .usage:
            return "Usage: swift pdf_formula_geometry.swift <pdf> <unnumbered-token> <numbered-token> [<align-token> <align-star-token>] | <pdf> --number-only | <pdf> --raster-only"
        case .unreadablePdf(let path):
            return "Unable to open PDF: \(path)"
        case .missingToken(let token):
            return "Unable to locate rendered formula token in PDF: \(token)"
        case .missingEquationNumber:
            return "Unable to locate the rendered equation number to the right of the numbered formula"
        }
    }
}

func formulaBounds(page: PDFPage, pageIndex: Int, text: String, range: NSRange) -> FormulaBounds? {
    guard let selection = page.selection(for: range) else { return nil }
    let bounds = selection.bounds(for: page)
    guard !bounds.isNull, !bounds.isEmpty else { return nil }
    var relationshipXs: [Double] = []
    if let pageText = page.string {
        let source = pageText as NSString
        let rangeEnd = min(source.length, range.location + range.length)
        if range.location < rangeEnd {
            for index in range.location..<rangeEnd where source.character(at: index) == 61 {
                if let relationSelection = page.selection(for: NSRange(location: index, length: 1)) {
                    let relationBounds = relationSelection.bounds(for: page)
                    if !relationBounds.isNull, !relationBounds.isEmpty {
                        relationshipXs.append(relationBounds.midX)
                    }
                }
            }
        }
    }
    return FormulaBounds(
        pageIndex: pageIndex,
        text: text,
        minX: bounds.minX,
        minY: bounds.minY,
        maxX: bounds.maxX,
        maxY: bounds.maxY,
        width: bounds.width,
        height: bounds.height,
        centerX: bounds.midX,
        centerY: bounds.midY,
        relationshipXs: relationshipXs
    )
}

func tokenRangeIgnoringWhitespace(token: String, in pageText: String) -> NSRange? {
    let source = pageText as NSString
    let target = token as NSString
    let whitespace = CharacterSet.whitespacesAndNewlines
    let targetUnits = (0..<target.length).compactMap { index -> unichar? in
        let unit = target.character(at: index)
        if let scalar = UnicodeScalar(unit), whitespace.contains(scalar) {
            return nil
        }
        return unit
    }
    guard !targetUnits.isEmpty else { return nil }

    for start in 0..<source.length {
        let firstScalar = UnicodeScalar(source.character(at: start))
        if let firstScalar, whitespace.contains(firstScalar) { continue }
        var sourceIndex = start
        var targetIndex = 0
        var lastMatched = start

        while sourceIndex < source.length && targetIndex < targetUnits.count {
            let sourceUnit = source.character(at: sourceIndex)
            if let scalar = UnicodeScalar(sourceUnit), whitespace.contains(scalar) {
                sourceIndex += 1
                continue
            }
            if sourceUnit != targetUnits[targetIndex] { break }
            lastMatched = sourceIndex
            sourceIndex += 1
            targetIndex += 1
        }
        if targetIndex == targetUnits.count {
            return NSRange(location: start, length: lastMatched - start + 1)
        }
    }
    return nil
}

func findToken(_ token: String, document: PDFDocument) throws -> FormulaBounds {
    var pageTexts: [String] = []
    for pageIndex in 0..<document.pageCount {
        guard let page = document.page(at: pageIndex), let pageText = page.string else { continue }
        pageTexts.append(pageText)
        let exactRange = (pageText as NSString).range(of: token)
        let range = exactRange.location != NSNotFound
            ? exactRange
            : tokenRangeIgnoringWhitespace(token: token, in: pageText)
        if let range,
           let bounds = formulaBounds(page: page, pageIndex: pageIndex, text: token, range: range) {
            return bounds
        }
    }
    throw GeometryError.missingToken("\(token); pages=\(pageTexts)")
}

func groupedIndices(_ active: [Bool], maximumGap: Int) -> [(Int, Int)] {
    var groups: [(Int, Int)] = []
    var start: Int?
    var lastActive = -1
    for index in active.indices {
        if active[index] {
            if start == nil || index - lastActive > maximumGap + 1 {
                if let start { groups.append((start, lastActive)) }
                start = index
            }
            lastActive = index
        }
    }
    if let start { groups.append((start, lastActive)) }
    return groups
}

func rasterInkComponents(page: PDFPage, scale: Int = 3) -> [RasterComponent] {
    let bounds = page.bounds(for: .cropBox)
    let width = max(1, Int(ceil(bounds.width * Double(scale))))
    let height = max(1, Int(ceil(bounds.height * Double(scale))))
    var pixels = [UInt8](repeating: 255, count: width * height)
    guard let context = CGContext(
        data: &pixels,
        width: width,
        height: height,
        bitsPerComponent: 8,
        bytesPerRow: width,
        space: CGColorSpaceCreateDeviceGray(),
        bitmapInfo: CGImageAlphaInfo.none.rawValue
    ) else { return [] }
    context.setFillColor(gray: 1, alpha: 1)
    context.fill(CGRect(x: 0, y: 0, width: width, height: height))
    context.saveGState()
    context.scaleBy(x: Double(scale), y: Double(scale))
    context.translateBy(x: -bounds.minX, y: -bounds.minY)
    page.draw(with: .cropBox, to: context)
    context.restoreGState()

    var visited = [Bool](repeating: false, count: width * height)
    var components: [RasterComponent] = []
    let neighborOffsets = [
        (-1, -1), (0, -1), (1, -1),
        (-1, 0),            (1, 0),
        (-1, 1),  (0, 1),   (1, 1),
    ]

    for startIndex in pixels.indices where
        pixels[startIndex] < 210 && !visited[startIndex] {
        visited[startIndex] = true
        var queue = [startIndex]
        var queueIndex = 0
        var minXpx = startIndex % width
        var maxXpx = minXpx
        var minYpx = startIndex / width
        var maxYpx = minYpx
        var darkPixelCount = 0

        while queueIndex < queue.count {
            let pixelIndex = queue[queueIndex]
            queueIndex += 1
            let x = pixelIndex % width
            let y = pixelIndex / width
            darkPixelCount += 1
            minXpx = min(minXpx, x)
            maxXpx = max(maxXpx, x)
            minYpx = min(minYpx, y)
            maxYpx = max(maxYpx, y)

            for (dx, dy) in neighborOffsets {
                let neighborX = x + dx
                let neighborY = y + dy
                guard neighborX >= 0, neighborX < width,
                      neighborY >= 0, neighborY < height else { continue }
                let neighborIndex = neighborY * width + neighborX
                guard !visited[neighborIndex],
                      pixels[neighborIndex] < 210 else { continue }
                visited[neighborIndex] = true
                queue.append(neighborIndex)
            }
        }

        guard darkPixelCount >= 2 else { continue }
        let minX = Double(minXpx) / Double(scale)
        let minY = Double(minYpx) / Double(scale)
        let maxX = Double(maxXpx + 1) / Double(scale)
        let maxY = Double(maxYpx + 1) / Double(scale)
        components.append(RasterComponent(
            minX: minX,
            minY: minY,
            maxX: maxX,
            maxY: maxY,
            width: maxX - minX,
            height: maxY - minY,
            centerX: (minX + maxX) / 2,
            centerY: (minY + maxY) / 2
        ))
    }
    return components
}

func rasterInkBands(page: PDFPage, scale: Int = 3) -> [RasterBand] {
    let bounds = page.bounds(for: .cropBox)
    let width = max(1, Int(ceil(bounds.width * Double(scale))))
    let height = max(1, Int(ceil(bounds.height * Double(scale))))
    var pixels = [UInt8](repeating: 255, count: width * height)
    guard let context = CGContext(
        data: &pixels,
        width: width,
        height: height,
        bitsPerComponent: 8,
        bytesPerRow: width,
        space: CGColorSpaceCreateDeviceGray(),
        bitmapInfo: CGImageAlphaInfo.none.rawValue
    ) else { return [] }
    context.setFillColor(gray: 1, alpha: 1)
    context.fill(CGRect(x: 0, y: 0, width: width, height: height))
    context.saveGState()
    context.scaleBy(x: Double(scale), y: Double(scale))
    context.translateBy(x: -bounds.minX, y: -bounds.minY)
    page.draw(with: .cropBox, to: context)
    context.restoreGState()

    var rowActive = [Bool](repeating: false, count: height)
    for y in 0..<height {
        var dark = 0
        for x in 0..<width where pixels[y * width + x] < 210 {
            dark += 1
            if dark >= 2 { break }
        }
        rowActive[y] = dark >= 2
    }

    return groupedIndices(rowActive, maximumGap: 3).compactMap { minYpx, maxYpx in
        guard maxYpx - minYpx + 1 >= 2 else { return nil }
        var columnActive = [Bool](repeating: false, count: width)
        var totalDark = 0
        for x in 0..<width {
            var dark = 0
            for y in minYpx...maxYpx where pixels[y * width + x] < 210 {
                dark += 1
            }
            columnActive[x] = dark > 0
            totalDark += dark
        }
        guard totalDark >= 8 else { return nil }
        let components = groupedIndices(columnActive, maximumGap: 2).compactMap { minXpx, maxXpx -> RasterComponent? in
            var componentMinYpx = maxYpx
            var componentMaxYpx = minYpx
            var componentHasInk = false
            for y in minYpx...maxYpx {
                for x in minXpx...maxXpx where pixels[y * width + x] < 210 {
                    componentHasInk = true
                    componentMinYpx = min(componentMinYpx, y)
                    componentMaxYpx = max(componentMaxYpx, y)
                }
            }
            guard componentHasInk else { return nil }
            let minX = Double(minXpx) / Double(scale)
            let minY = Double(componentMinYpx) / Double(scale)
            let maxX = Double(maxXpx + 1) / Double(scale)
            let maxY = Double(componentMaxYpx + 1) / Double(scale)
            return RasterComponent(
                minX: minX,
                minY: minY,
                maxX: maxX,
                maxY: maxY,
                width: maxX - minX,
                height: maxY - minY,
                centerX: (minX + maxX) / 2,
                centerY: (minY + maxY) / 2
            )
        }.filter { $0.width >= 0.5 }
        guard !components.isEmpty else { return nil }
        let minY = Double(minYpx) / Double(scale)
        let maxY = Double(maxYpx + 1) / Double(scale)
        return RasterBand(
            minY: minY,
            maxY: maxY,
            height: maxY - minY,
            centerY: (minY + maxY) / 2,
            inkWidth: components.reduce(0) { $0 + $1.width },
            components: components
        )
    }
}

func findRightmostEquationNumber(document: PDFDocument) throws -> FormulaBounds {
    var candidates: [FormulaBounds] = []
    for pageIndex in 0..<document.pageCount {
        guard let page = document.page(at: pageIndex), let pageText = page.string else { continue }
        let source = pageText as NSString
        var searchRange = NSRange(location: 0, length: source.length)
        while searchRange.length > 0 {
            let found = source.range(of: "1", options: [], range: searchRange)
            if found.location == NSNotFound { break }
            if let bounds = formulaBounds(page: page, pageIndex: pageIndex, text: "1", range: found) {
                candidates.append(bounds)
            }
            let nextLocation = found.location + found.length
            if nextLocation >= source.length { break }
            searchRange = NSRange(location: nextLocation, length: source.length - nextLocation)
        }
    }
    guard let best = candidates.max(by: { $0.centerX < $1.centerX }) else {
        throw GeometryError.missingEquationNumber
    }
    return best
}

func findEquationNumber(document: PDFDocument, numbered: FormulaBounds) throws -> FormulaBounds {
    guard let page = document.page(at: numbered.pageIndex), let pageText = page.string else {
        throw GeometryError.missingEquationNumber
    }
    let source = pageText as NSString
    var searchRange = NSRange(location: 0, length: source.length)
    var candidates: [FormulaBounds] = []

    while searchRange.length > 0 {
        let found = source.range(of: "1", options: [], range: searchRange)
        if found.location == NSNotFound { break }
        if let bounds = formulaBounds(page: page, pageIndex: numbered.pageIndex, text: "1", range: found),
           bounds.minX > numbered.maxX + 4,
           abs(bounds.centerY - numbered.centerY) <= max(12, numbered.height) {
            candidates.append(bounds)
        }
        let nextLocation = found.location + found.length
        if nextLocation >= source.length { break }
        searchRange = NSRange(location: nextLocation, length: source.length - nextLocation)
    }

    guard let best = candidates.min(by: {
        let lhsVertical = abs($0.centerY - numbered.centerY)
        let rhsVertical = abs($1.centerY - numbered.centerY)
        if abs(lhsVertical - rhsVertical) > 0.01 { return lhsVertical < rhsVertical }
        return $0.minX > $1.minX
    }) else {
        throw GeometryError.missingEquationNumber
    }
    return best
}

do {
    guard CommandLine.arguments.count == 3 ||
          CommandLine.arguments.count == 4 ||
          CommandLine.arguments.count == 6 else {
        throw GeometryError.usage
    }
    let pdfPath = CommandLine.arguments[1]
    guard let document = PDFDocument(url: URL(fileURLWithPath: pdfPath)) else {
        throw GeometryError.unreadablePdf(pdfPath)
    }
    if CommandLine.arguments.count == 3 {
        let mode = CommandLine.arguments[2]
        guard mode == "--number-only" || mode == "--raster-only" else {
            throw GeometryError.usage
        }
        let encoder = JSONEncoder()
        encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
        if mode == "--raster-only" {
            guard let page = document.page(at: 0) else {
                throw GeometryError.unreadablePdf(pdfPath)
            }
            let pageBounds = page.bounds(for: .cropBox)
            let report = RasterOnlyReport(
                pageWidth: pageBounds.width,
                pageHeight: pageBounds.height,
                rasterBands: rasterInkBands(page: page),
                rasterComponents: rasterInkComponents(page: page),
                pageText: page.string ?? ""
            )
            FileHandle.standardOutput.write(try encoder.encode(report))
        } else {
            let equationNumber = try findRightmostEquationNumber(document: document)
            guard let page = document.page(at: equationNumber.pageIndex) else {
                throw GeometryError.unreadablePdf(pdfPath)
            }
            let pageBounds = page.bounds(for: .cropBox)
            let report = NumberOnlyReport(
                pageWidth: pageBounds.width,
                pageHeight: pageBounds.height,
                equationNumber: equationNumber,
                rasterBands: rasterInkBands(page: page),
                rasterComponents: rasterInkComponents(page: page),
                pageText: page.string ?? ""
            )
            FileHandle.standardOutput.write(try encoder.encode(report))
        }
        FileHandle.standardOutput.write(Data("\n".utf8))
        exit(0)
    }

    let unnumberedToken = CommandLine.arguments[2]
    let numberedToken = CommandLine.arguments[3]
    let unnumbered = try findToken(unnumberedToken, document: document)
    let numbered = try findToken(numberedToken, document: document)
    let aligned = CommandLine.arguments.count == 6
        ? [
            try findToken(CommandLine.arguments[4], document: document),
            try findToken(CommandLine.arguments[5], document: document),
        ]
        : []
    let equationNumber = try findEquationNumber(document: document, numbered: numbered)
    guard let page = document.page(at: numbered.pageIndex) else {
        throw GeometryError.unreadablePdf(pdfPath)
    }
    let pageBounds = page.bounds(for: .cropBox)
    let report = GeometryReport(
        pageWidth: pageBounds.width,
        pageHeight: pageBounds.height,
        unnumbered: unnumbered,
        numbered: numbered,
        equationNumber: equationNumber,
        aligned: aligned,
        pageText: page.string ?? ""
    )
    let encoder = JSONEncoder()
    encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
    FileHandle.standardOutput.write(try encoder.encode(report))
    FileHandle.standardOutput.write(Data("\n".utf8))
} catch {
    FileHandle.standardError.write(Data("\(error)\n".utf8))
    exit(1)
}
