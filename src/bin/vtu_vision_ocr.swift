import Foundation
import Vision
import AppKit

func recognize(imagePath: String) -> String? {
    let url = URL(fileURLWithPath: imagePath)
    guard let image = NSImage(contentsOf: url),
          let tiffData = image.tiffRepresentation,
          let bitmap = NSBitmapImageRep(data: tiffData),
          let cgImage = bitmap.cgImage else {
        return nil
    }

    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = false
    request.recognitionLanguages = ["en-US"]

    let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])
    do {
        try handler.perform([request])
        guard let results = request.results, !results.isEmpty else {
            return nil
        }
        var recognized = ""
        for observation in results {
            if let top = observation.topCandidates(1).first {
                recognized += top.string
            }
        }
        return recognized.trimmingCharacters(in: .whitespacesAndNewlines)
    } catch {
        return nil
    }
}

let args = CommandLine.arguments
if args.count > 1 {
    if let result = recognize(imagePath: args[1]) {
        print(result)
        exit(0)
    } else {
        exit(1)
    }
} else {
    print("Usage: vtu_vision_ocr <image_path>")
    exit(1)
}
