import JSZip from "jszip";

export type AIAssetStatus = "pending" | "extracted" | "partial" | "unsupported" | "error";

export type AIUploadedAsset = {
  id: string;
  filename: string;
  mimeType: string;
  size: number;
  status: AIAssetStatus;
  extractedText: string;
  message?: string;
};

type ExtractionResult = Omit<AIUploadedAsset, "id">;

const textExtensions = [".txt", ".md"];
const imageExtensions = [".png", ".jpg", ".jpeg"];

export function isAISupportedUpload(filename: string) {
  const extension = getExtension(filename);
  return [...textExtensions, ...imageExtensions, ".docx", ".pptx", ".pdf"].includes(extension);
}

export async function extractTextFromFile(file: File): Promise<ExtractionResult> {
  const extension = getExtension(file.name);
  const base = {
    filename: file.name,
    mimeType: file.type || extension.replace(".", "").toUpperCase(),
    size: file.size,
  };

  try {
    if (textExtensions.includes(extension)) {
      const text = await file.text();
      return {
        ...base,
        status: text.trim() ? "extracted" : "partial",
        extractedText: text,
        message: text.trim() ? "Plain text extracted." : "File was readable, but no text was found.",
      };
    }

    if (extension === ".docx") {
      return {
        ...base,
        ...(await extractDocxText(file)),
      };
    }

    if (extension === ".pptx") {
      return {
        ...base,
        ...(await extractPptxText(file)),
      };
    }

    if (extension === ".pdf") {
      return {
        ...base,
        status: "unsupported",
        extractedText: "",
        message: "PDF text extraction is not available in this browser flow yet.",
      };
    }

    if (imageExtensions.includes(extension)) {
      return {
        ...base,
        status: "partial",
        extractedText: "",
        message: "Image uploaded; OCR not implemented yet.",
      };
    }

    return {
      ...base,
      status: "unsupported",
      extractedText: "",
      message: "Unsupported file type for AI extraction.",
    };
  } catch (error) {
    return {
      ...base,
      status: "error",
      extractedText: "",
      message: error instanceof Error ? error.message : "Text extraction failed.",
    };
  }
}

async function extractDocxText(file: File) {
  const zip = await JSZip.loadAsync(file);
  const documentXml = await zip.file("word/document.xml")?.async("text");

  if (!documentXml) {
    return {
      status: "error" as const,
      extractedText: "",
      message: "DOCX document XML was not found.",
    };
  }

  const text = extractTextNodes(documentXml, "w:t").join(" ").replace(/\s+/g, " ").trim();

  return {
    status: text ? "extracted" as const : "partial" as const,
    extractedText: text,
    message: text ? "DOCX text extracted from document XML." : "DOCX opened, but no readable text was found.",
  };
}

async function extractPptxText(file: File) {
  const zip = await JSZip.loadAsync(file);
  const slideFiles = Object.keys(zip.files)
    .filter((path) => /^ppt\/slides\/slide\d+\.xml$/i.test(path))
    .sort((a, b) => getSlideNumber(a) - getSlideNumber(b));

  if (!slideFiles.length) {
    return {
      status: "error" as const,
      extractedText: "",
      message: "PPTX slide XML was not found.",
    };
  }

  const slideTexts = await Promise.all(
    slideFiles.map(async (path, index) => {
      const xml = await zip.file(path)?.async("text");
      const text = xml ? extractTextNodes(xml, "a:t").join(" ").replace(/\s+/g, " ").trim() : "";
      return text ? `Slide ${index + 1}: ${text}` : "";
    }),
  );
  const text = slideTexts.filter(Boolean).join("\n\n");

  return {
    status: text ? "extracted" as const : "partial" as const,
    extractedText: text,
    message: text ? "PPTX text extracted from slide XML." : "PPTX opened, but no readable slide text was found.",
  };
}

function extractTextNodes(xml: string, tagName: string) {
  const parser = new DOMParser();
  const documentXml = parser.parseFromString(xml, "application/xml");
  const parserError = documentXml.querySelector("parsererror");

  if (parserError) {
    throw new Error("Could not parse document XML.");
  }

  return Array.from(documentXml.getElementsByTagName(tagName))
    .map((node) => node.textContent?.trim() ?? "")
    .filter(Boolean);
}

function getExtension(filename: string) {
  const dotIndex = filename.lastIndexOf(".");
  return dotIndex >= 0 ? filename.slice(dotIndex).toLowerCase() : "";
}

function getSlideNumber(path: string) {
  return Number(path.match(/slide(\d+)\.xml$/i)?.[1] ?? 0);
}
