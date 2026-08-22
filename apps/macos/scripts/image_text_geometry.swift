import Foundation
import Vision

struct TextObservation: Codable {
    let text: String
    let confidence: Float
    let minX: Double
    let minY: Double
    let width: Double
    let height: Double
}

guard CommandLine.arguments.count == 2 else {
    fputs("Usage: swift image_text_geometry.swift <image>\n", stderr)
    exit(2)
}

let imageURL = URL(fileURLWithPath: CommandLine.arguments[1])
let request = VNRecognizeTextRequest()
request.recognitionLevel = .accurate
request.recognitionLanguages = ["zh-Hans", "en-US"]
request.usesLanguageCorrection = false
try VNImageRequestHandler(url: imageURL).perform([request])

let observations = (request.results ?? []).compactMap { observation -> TextObservation? in
    guard let candidate = observation.topCandidates(1).first else { return nil }
    let box = observation.boundingBox
    return TextObservation(
        text: candidate.string,
        confidence: candidate.confidence,
        minX: box.minX,
        minY: box.minY,
        width: box.width,
        height: box.height
    )
}
let encoder = JSONEncoder()
encoder.outputFormatting = [.prettyPrinted, .sortedKeys]
FileHandle.standardOutput.write(try encoder.encode(observations))
