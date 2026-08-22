import CoreGraphics
import Foundation
import ImageIO

struct RasterComponent: Codable {
    let minX: Int
    let minY: Int
    let maxX: Int
    let maxY: Int
    let width: Int
    let height: Int
    let centerX: Double
    let centerY: Double
    let pixelCount: Int
}

guard CommandLine.arguments.count == 6,
      let cropMinX = Int(CommandLine.arguments[2]),
      let cropMinY = Int(CommandLine.arguments[3]),
      let cropMaxX = Int(CommandLine.arguments[4]),
      let cropMaxY = Int(CommandLine.arguments[5]) else {
    fputs("Usage: swift image_raster_geometry.swift <image> <minX> <minY> <maxX> <maxY>\n", stderr)
    exit(2)
}

let imageURL = URL(fileURLWithPath: CommandLine.arguments[1]) as CFURL
guard let source = CGImageSourceCreateWithURL(imageURL, nil),
      let image = CGImageSourceCreateImageAtIndex(source, 0, nil) else {
    throw NSError(domain: "VisualTeXImageGeometry", code: 1)
}
let width = image.width
let height = image.height
guard cropMinX >= 0, cropMinY >= 0, cropMaxX <= width, cropMaxY <= height,
      cropMaxX > cropMinX, cropMaxY > cropMinY else {
    throw NSError(domain: "VisualTeXImageGeometry", code: 2)
}

var pixels = [UInt8](repeating: 255, count: width * height)
guard let context = CGContext(
    data: &pixels,
    width: width,
    height: height,
    bitsPerComponent: 8,
    bytesPerRow: width,
    space: CGColorSpaceCreateDeviceGray(),
    bitmapInfo: CGImageAlphaInfo.none.rawValue
) else {
    throw NSError(domain: "VisualTeXImageGeometry", code: 3)
}
context.setFillColor(gray: 1, alpha: 1)
context.fill(CGRect(x: 0, y: 0, width: width, height: height))
context.translateBy(x: 0, y: CGFloat(height))
context.scaleBy(x: 1, y: -1)
context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))

var visited = [Bool](repeating: false, count: width * height)
let neighbors = [
    (-1, -1), (0, -1), (1, -1),
    (-1, 0),            (1, 0),
    (-1, 1),  (0, 1),   (1, 1),
]
var components: [RasterComponent] = []

for y in cropMinY..<cropMaxY {
    for x in cropMinX..<cropMaxX {
        let start = y * width + x
        guard pixels[start] < 190, !visited[start] else { continue }
        visited[start] = true
        var queue = [start]
        var queueIndex = 0
        var minX = x
        var maxX = x
        var minY = y
        var maxY = y
        while queueIndex < queue.count {
            let index = queue[queueIndex]
            queueIndex += 1
            let currentX = index % width
            let currentY = index / width
            minX = min(minX, currentX)
            maxX = max(maxX, currentX)
            minY = min(minY, currentY)
            maxY = max(maxY, currentY)
            for (dx, dy) in neighbors {
                let nextX = currentX + dx
                let nextY = currentY + dy
                guard nextX >= cropMinX, nextX < cropMaxX,
                      nextY >= cropMinY, nextY < cropMaxY else { continue }
                let next = nextY * width + nextX
                guard pixels[next] < 190, !visited[next] else { continue }
                visited[next] = true
                queue.append(next)
            }
        }
        guard queue.count >= 3 else { continue }
        components.append(RasterComponent(
            minX: minX,
            minY: minY,
            maxX: maxX + 1,
            maxY: maxY + 1,
            width: maxX - minX + 1,
            height: maxY - minY + 1,
            centerX: Double(minX + maxX + 1) / 2,
            centerY: Double(minY + maxY + 1) / 2,
            pixelCount: queue.count
        ))
    }
}

let encoder = JSONEncoder()
encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
FileHandle.standardOutput.write(try encoder.encode(components))
