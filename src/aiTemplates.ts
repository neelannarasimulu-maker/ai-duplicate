export type AIOutputLength =
  | "Very short"
  | "Short"
  | "Medium"
  | "Detailed"
  | "Maximum 1 page"
  | "Maximum 2 pages"
  | "Executive summary only"
  | "Comprehensive";

export type AIOutputFormat =
  | "Markdown"
  | "DOCX-ready content"
  | "Email"
  | "Table"
  | "Presentation-ready Markdown"
  | "Image prompt"
  | "JSON";

export type AITemplate = {
  id: string;
  name: string;
  description: string;
  outputType: string;
  defaultAudience: string;
  defaultTone: string;
  defaultFormat: AIOutputFormat;
  defaultLength: AIOutputLength;
  sections: string[];
};

export const AIOutputLengths: AIOutputLength[] = [
  "Very short",
  "Short",
  "Medium",
  "Detailed",
  "Maximum 1 page",
  "Maximum 2 pages",
  "Executive summary only",
  "Comprehensive",
];

export const AIOutputFormats: AIOutputFormat[] = [
  "Markdown",
  "DOCX-ready content",
  "Email",
  "Table",
  "Presentation-ready Markdown",
  "Image prompt",
  "JSON",
];

export const AITemplates: AITemplate[] = [
  {
    id: "adhoc-generation",
    name: "Adhoc Generation",
    description: "Start with a blank structure and add only the sections needed for this output.",
    outputType: "Adhoc generated output",
    defaultAudience: "The intended reader or user",
    defaultTone: "Clear, useful, and context-aware",
    defaultFormat: "Markdown",
    defaultLength: "Medium",
    sections: [],
  },
  {
    id: "business-case",
    name: "Business Case Document",
    description: "A decision-ready business case with rationale, options, impact, risks, and required decisions.",
    outputType: "Business case document",
    defaultAudience: "Decision makers and business stakeholders",
    defaultTone: "Professional, evidence-led, and practical",
    defaultFormat: "Markdown",
    defaultLength: "Detailed",
    sections: [
      "Title",
      "Executive summary",
      "Background",
      "Problem or opportunity",
      "Strategic rationale",
      "Options considered",
      "Recommended approach",
      "Financial or operational impact",
      "Risks and mitigations",
      "Assumptions",
      "Implementation plan",
      "Success measures",
      "Decisions required",
    ],
  },
  {
    id: "customer-next-steps",
    name: "Customer Information and Next Steps",
    description: "A structured customer note that captures context, requirements, risks, open questions, and actions.",
    outputType: "Customer information and next steps document",
    defaultAudience: "Account team and customer stakeholders",
    defaultTone: "Clear, accountable, and action-oriented",
    defaultFormat: "Markdown",
    defaultLength: "Medium",
    sections: [
      "Customer overview",
      "Current context",
      "Key information received",
      "Needs or requirements",
      "Opportunities",
      "Risks or concerns",
      "Decisions made",
      "Open questions",
      "Next steps",
      "Owner and due date",
    ],
  },
  {
    id: "general-business",
    name: "General Business Document",
    description: "A reusable business document structure for turning rough source material into polished output.",
    outputType: "General business document",
    defaultAudience: "Internal team",
    defaultTone: "Professional, clear, and concise",
    defaultFormat: "Markdown",
    defaultLength: "Medium",
    sections: [
      "Title",
      "Purpose",
      "Background",
      "Key points",
      "Main content",
      "Risks or gaps",
      "Recommendations",
      "Remaining decisions",
      "Next steps",
    ],
  },
  {
    id: "proposal",
    name: "Proposal Document",
    description: "A client-ready proposal outline with approach, deliverables, value, timeline, and assumptions.",
    outputType: "Proposal document",
    defaultAudience: "Client or sponsor",
    defaultTone: "Confident, useful, and commercially grounded",
    defaultFormat: "DOCX-ready content",
    defaultLength: "Detailed",
    sections: [
      "Overview",
      "Client need",
      "Proposed approach",
      "Deliverables",
      "Value or benefit",
      "Timeline",
      "Assumptions",
      "Commercial considerations",
      "Next steps",
    ],
  },
  {
    id: "meeting-brief",
    name: "Meeting Brief",
    description: "A practical pre-read for meetings where outcomes, agenda, status, decisions, and actions matter.",
    outputType: "Meeting brief",
    defaultAudience: "Meeting attendees and decision makers",
    defaultTone: "Concise, structured, and decision-focused",
    defaultFormat: "Markdown",
    defaultLength: "Short",
    sections: [
      "Purpose",
      "Desired outcomes",
      "Agenda",
      "Current status",
      "Key discussion points",
      "Decisions needed",
      "Actions",
    ],
  },
  {
    id: "executive-summary",
    name: "Executive Summary",
    description: "A compact executive summary that emphasizes facts, implications, risks, and decisions.",
    outputType: "Executive summary",
    defaultAudience: "Executives and senior stakeholders",
    defaultTone: "Brief, direct, and decision-focused",
    defaultFormat: "Markdown",
    defaultLength: "Executive summary only",
    sections: [
      "Executive summary",
      "Key facts",
      "Implications",
      "Risks or gaps",
      "Decisions required",
    ],
  },
  {
    id: "image-prompt-generator",
    name: "Image Prompt Generator",
    description: "Turns rough visual ideas or business context into a polished prompt you can paste into DALL-E.",
    outputType: "Image generation prompt",
    defaultAudience: "DALL-E image generation model",
    defaultTone: "Specific, visual, and production-ready",
    defaultFormat: "Image prompt",
    defaultLength: "Detailed",
    sections: [
      "Primary image prompt",
      "Subject and focal point",
      "Scene and environment",
      "Composition and camera framing",
      "Style and medium",
      "Lighting and color palette",
      "Text or logo handling",
      "Negative prompt",
      "Optional variations",
    ],
  },
];
