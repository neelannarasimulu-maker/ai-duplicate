import { ArrowDown, ArrowUp, Clipboard, FileText, Plus, RefreshCw, Trash2, Upload, WandSparkles, X } from "lucide-react";
import { ChangeEvent, useMemo, useState } from "react";
import { extractTextFromFile, isAISupportedUpload, type AIUploadedAsset } from "../fileTextExtractors";
import { AIOutputFormats, AIOutputLengths, AITemplates, type AIOutputFormat, type AIOutputLength } from "../aiTemplates";
import type { SavedOutput, WorkTask } from "../types";

const maxSourceWarningCharacters = 100_000;

const defaultTemplate = AITemplates[0];

type AIFormState = {
  sourceText: string;
  outputTitle: string;
  outputType: string;
  audience: string;
  tone: string;
  outputLength: AIOutputLength;
  outputFormat: AIOutputFormat;
  selectedTemplateId: string;
  sections: string[];
  additionalInstructions: string;
};

const initialFormState: AIFormState = {
  sourceText: "",
  outputTitle: "",
  outputType: defaultTemplate.outputType,
  audience: defaultTemplate.defaultAudience,
  tone: defaultTemplate.defaultTone,
  outputLength: defaultTemplate.defaultLength,
  outputFormat: defaultTemplate.defaultFormat,
  selectedTemplateId: defaultTemplate.id,
  sections: defaultTemplate.sections,
  additionalInstructions: "",
};

type AIPromptBuilderProps = {
  activeTask?: WorkTask;
  onOutputChange?: (value: string) => void;
  onPromptGenerated?: (prompt: string, sourceText: string) => void;
  onSaveActivity?: (activity: { prompt: string; output: string; sourceText: string; form: AIFormState }) => void;
  onTaskSelect?: (taskId: string) => void;
  taskHistory?: SavedOutput[];
  tasks?: WorkTask[];
};

export function AIPromptBuilder({
  activeTask,
  onOutputChange,
  onPromptGenerated,
  onSaveActivity,
  onTaskSelect,
  taskHistory = [],
  tasks = [],
}: AIPromptBuilderProps) {
  const [form, setForm] = useState<AIFormState>(initialFormState);
  const [assets, setAssets] = useState<AIUploadedAsset[]>([]);
  const [newSection, setNewSection] = useState("");
  const [generatedPrompt, setGeneratedPrompt] = useState("");
  const [chatGptOutput, setChatGptOutput] = useState("");
  const [error, setError] = useState("");
  const [copyMessage, setCopyMessage] = useState("");

  const selectedTemplate = AITemplates.find((template) => template.id === form.selectedTemplateId) ?? defaultTemplate;
  const extractedSource = useMemo(
    () =>
      assets
        .filter((asset) => asset.extractedText.trim())
        .map((asset) => `SOURCE FILE: ${asset.filename}\n${asset.extractedText.trim()}`)
        .join("\n\n"),
    [assets],
  );
  const linkedTaskSource = useMemo(() => buildLinkedTaskSource(activeTask), [activeTask]);
  const typedSourceIncludesLinkedTask = form.sourceText.includes("LINKED TASK CONTEXT");
  const combinedSourceMaterial = [
    typedSourceIncludesLinkedTask ? "" : linkedTaskSource,
    form.sourceText.trim(),
    extractedSource,
  ].filter(Boolean).join("\n\n");
  const combinedCharacterCount = combinedSourceMaterial.length;
  const sourceTooLarge = combinedCharacterCount > maxSourceWarningCharacters;

  async function handleFilesSelected(event: ChangeEvent<HTMLInputElement>) {
    const files = Array.from(event.target.files ?? []);
    if (!files.length) return;

    const pendingAssets: AIUploadedAsset[] = files.map((file) => ({
      id: createId(),
      filename: file.name,
      mimeType: file.type || "Unknown",
      size: file.size,
      status: "pending",
      extractedText: "",
      message: isAISupportedUpload(file.name) ? "Queued for extraction." : "Unsupported file type for AI extraction.",
    }));

    setAssets((current) => [...current, ...pendingAssets]);
    event.target.value = "";

    await Promise.all(
      files.map(async (file, index) => {
        const result = await extractTextFromFile(file);
        setAssets((current) =>
          current.map((asset) => (asset.id === pendingAssets[index].id ? { id: asset.id, ...result } : asset)),
        );
      }),
    );
  }

  function selectTemplate(templateId: string) {
    const template = AITemplates.find((item) => item.id === templateId) ?? defaultTemplate;
    setForm((current) => ({
      ...current,
      selectedTemplateId: template.id,
      outputType: template.outputType,
      audience: template.defaultAudience,
      tone: template.defaultTone,
      outputFormat: template.defaultFormat,
      outputLength: template.defaultLength,
      sections: [...template.sections],
    }));
    setGeneratedPrompt("");
    setError("");
  }

  function updateSection(index: number, value: string) {
    setForm((current) => ({
      ...current,
      sections: current.sections.map((section, sectionIndex) => (sectionIndex === index ? value : section)),
    }));
  }

  function removeSection(index: number) {
    setForm((current) => ({
      ...current,
      sections: current.sections.filter((_, sectionIndex) => sectionIndex !== index),
    }));
  }

  function moveSection(index: number, direction: -1 | 1) {
    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= form.sections.length) return;

    setForm((current) => {
      const sections = [...current.sections];
      const [section] = sections.splice(index, 1);
      sections.splice(targetIndex, 0, section);
      return { ...current, sections };
    });
  }

  function addSection() {
    const section = newSection.trim();
    if (!section) return;
    setForm((current) => ({ ...current, sections: [...current.sections, section] }));
    setNewSection("");
  }

  function generatePrompt() {
    const usableSections = form.sections.map((section) => section.trim()).filter(Boolean);
    const sourceMaterial = combinedSourceMaterial.trim();

    if (!sourceMaterial) {
      setError("Select a task, add typed source material, or upload a file with extractable text before generating the master prompt.");
      setGeneratedPrompt("");
      return;
    }

    setError("");
    setCopyMessage("");
    const promptForm = {
      ...form,
      outputTitle: form.outputTitle.trim() || activeTask?.title || "",
      sections: usableSections,
    };
    const prompt = buildMasterPrompt(promptForm, sourceMaterial);
    setGeneratedPrompt(prompt);
    onPromptGenerated?.(prompt, sourceMaterial);
  }

  async function copyPrompt() {
    if (!generatedPrompt) return;

    try {
      await navigator.clipboard.writeText(generatedPrompt);
      setCopyMessage("Prompt copied to clipboard.");
    } catch {
      setCopyMessage("Clipboard copy failed. Select the prompt text and copy it manually.");
    }
  }

  function clearForm() {
    setForm(initialFormState);
    setAssets([]);
    setNewSection("");
    setGeneratedPrompt("");
    setChatGptOutput("");
    onOutputChange?.("");
    setError("");
    setCopyMessage("");
  }

  function updateChatGptOutput(value: string) {
    setChatGptOutput(value);
    onOutputChange?.(value);
  }

  function saveActivity() {
    if (!activeTask) {
      setError("Select or create a task before saving AI activity.");
      return;
    }
    if (!generatedPrompt.trim() && !chatGptOutput.trim()) {
      setError("Generate a prompt or paste ChatGPT output before saving activity.");
      return;
    }
    setError("");
    onSaveActivity?.({
      prompt: generatedPrompt,
      output: chatGptOutput,
      sourceText: combinedSourceMaterial,
      form,
    });
    setCopyMessage("AI activity saved to the linked task.");
  }

  function loadHistoryItem(item: SavedOutput) {
    setForm((current) => ({ ...current, sourceText: item.input ?? "" }));
    setGeneratedPrompt(item.gptPrompt ?? "");
    setChatGptOutput(item.result ?? "");
    onPromptGenerated?.(item.gptPrompt ?? "", item.input ?? "");
    onOutputChange?.(item.result ?? "");
    setCopyMessage("Loaded AI activity from task history.");
  }

  return (
    <div className="ai-engine-builder">
      <section className="ai-engine-hero panel">
        <div>
          <p className="eyebrow">AI generation engine</p>
          <h2>Prompt Builder</h2>
          <p>
            Build a high-context master prompt from source material, templates, editable sections, and quality rules.
          </p>
        </div>
        <button className="ai-action" onClick={generatePrompt} type="button">
          <WandSparkles size={16} />
          Generate master prompt
        </button>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={0} title="Linked Task" />
        <div className="form-grid">
          <label>
            AI activity task
            <select
              onChange={(event) => onTaskSelect?.(event.target.value)}
              value={activeTask?.id ?? ""}
            >
              <option value="">Select a task</option>
              {tasks.map((task) => (
                <option key={task.id} value={task.id}>
                  {task.title}
                </option>
              ))}
            </select>
            <small className="field-helper">Prompts and pasted ChatGPT outputs are saved against the selected task.</small>
          </label>
          <div className="linked-template-field">
            <span>Current link</span>
            <strong>{activeTask ? activeTask.title : "No task selected"}</strong>
          </div>
        </div>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={1} title="Source Input" />
        <label>
          Paste or type source material
          <textarea
            className="ai-engine-source-textarea"
            onChange={(event) => setForm((current) => ({ ...current, sourceText: event.target.value }))}
            placeholder="Paste notes, transcript text, customer information, draft content, or rough source material..."
            value={form.sourceText}
          />
        </label>
        <label className="upload-control ai-engine-upload-control">
          <Upload size={16} />
          Upload documents or images
          <input
            accept=".txt,.md,.docx,.pptx,.pdf,.png,.jpg,.jpeg"
            multiple
            onChange={handleFilesSelected}
            type="file"
          />
        </label>
        <div className="ai-engine-asset-list" aria-label="Uploaded files">
          {assets.map((asset) => (
            <div className="ai-engine-asset-row" key={asset.id}>
              <FileText size={18} />
              <div>
                <strong>{asset.filename}</strong>
                <span>{asset.mimeType || "Unknown type"} - {formatBytes(asset.size)}</span>
              </div>
              <StatusPill status={asset.status} />
              <span>{asset.extractedText.length.toLocaleString()} chars</span>
              <span>{asset.message}</span>
              <button className="ghost-button icon-button" onClick={() => setAssets((current) => current.filter((item) => item.id !== asset.id))} type="button" title="Remove file">
                <X size={15} />
              </button>
            </div>
          ))}
          {assets.length === 0 && <p className="empty compact-empty">No files uploaded yet.</p>}
        </div>
        <div className="ai-engine-count-row">
          <strong>Combined source character count:</strong>
          <span>{combinedCharacterCount.toLocaleString()}</span>
          {sourceTooLarge && <span className="ai-engine-warning">Large source warning: over 100,000 characters.</span>}
        </div>
        <label>
          Consolidated source material preview
          <textarea className="small-textarea ai-engine-preview-textarea" readOnly value={combinedSourceMaterial} />
        </label>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={2} title="Output Definition" />
        <div className="form-grid">
          <label>
            Output title
            <input
              onChange={(event) => setForm((current) => ({ ...current, outputTitle: event.target.value }))}
              placeholder="Example: May customer update"
              value={form.outputTitle}
            />
          </label>
          <label>
            Output type
            <input
              onChange={(event) => setForm((current) => ({ ...current, outputType: event.target.value }))}
              value={form.outputType}
            />
          </label>
          <label>
            Audience
            <input
              onChange={(event) => setForm((current) => ({ ...current, audience: event.target.value }))}
              value={form.audience}
            />
          </label>
          <label>
            Tone
            <input
              onChange={(event) => setForm((current) => ({ ...current, tone: event.target.value }))}
              value={form.tone}
            />
          </label>
          <label>
            Output length
            <select
              onChange={(event) => setForm((current) => ({ ...current, outputLength: event.target.value as AIOutputLength }))}
              value={form.outputLength}
            >
              {AIOutputLengths.map((length) => (
                <option key={length} value={length}>{length}</option>
              ))}
            </select>
          </label>
          <label>
            Output format
            <select
              onChange={(event) => setForm((current) => ({ ...current, outputFormat: event.target.value as AIOutputFormat }))}
              value={form.outputFormat}
            >
              {AIOutputFormats.map((format) => (
                <option key={format} value={format}>{format}</option>
              ))}
            </select>
            <small className="field-helper">
              ChatGPT will return formatted text. The app must export the final content to .docx if a real Word file is required.
            </small>
          </label>
        </div>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={3} title="Template Selection" />
        <div className="ai-engine-template-grid">
          {AITemplates.map((template) => (
            <button
              className={template.id === selectedTemplate.id ? "ai-engine-template-button active" : "ai-engine-template-button"}
              key={template.id}
              onClick={() => selectTemplate(template.id)}
              type="button"
            >
              <strong>{template.name}</strong>
              <span>{template.description}</span>
            </button>
          ))}
        </div>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={4} title="Section Override" />
        <div className="ai-engine-section-list">
          {form.sections.map((section, index) => (
            <div className="ai-engine-section-row" key={`${section}-${index}`}>
              <input
                aria-label={`Section ${index + 1}`}
                onChange={(event) => updateSection(index, event.target.value)}
                value={section}
              />
              <button className="ghost-button icon-button" disabled={index === 0} onClick={() => moveSection(index, -1)} type="button" title="Move section up">
                <ArrowUp size={15} />
              </button>
              <button className="ghost-button icon-button" disabled={index === form.sections.length - 1} onClick={() => moveSection(index, 1)} type="button" title="Move section down">
                <ArrowDown size={15} />
              </button>
              <button className="ghost-button icon-button" onClick={() => removeSection(index)} type="button" title="Remove section">
                <Trash2 size={15} />
              </button>
            </div>
          ))}
          {form.sections.length === 0 && <p className="empty compact-empty">No sections defined. Add the sections you need below.</p>}
        </div>
        <div className="ai-engine-add-section">
          <input
            aria-label="New section name"
            onChange={(event) => setNewSection(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") addSection();
            }}
            placeholder="Add another section"
            value={newSection}
          />
          <button className="ghost-button" onClick={addSection} type="button">
            <Plus size={16} />
            Add section
          </button>
        </div>
        <label>
          Additional instructions or constraints
          <textarea
            className="small-textarea"
            onChange={(event) => setForm((current) => ({ ...current, additionalInstructions: event.target.value }))}
            placeholder="Example: Use South African English. Keep financial assumptions separate from confirmed figures."
            value={form.additionalInstructions}
          />
        </label>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={5} title="Master Prompt" />
        {error && <div className="sync-banner ai-engine-error" role="alert">{error}</div>}
        {copyMessage && <div className="sync-banner" role="status">{copyMessage}</div>}
        <textarea className="output-textarea ai-engine-master-prompt" readOnly value={generatedPrompt} />
        <div className="ai-engine-actions">
          <button className="primary-button" disabled={!generatedPrompt} onClick={copyPrompt} type="button">
            <Clipboard size={16} />
            Copy prompt
          </button>
          <button className="ai-action" onClick={generatePrompt} type="button">
            <RefreshCw size={16} />
            Regenerate prompt
          </button>
          <button className="ghost-button" onClick={clearForm} type="button">
            <Trash2 size={16} />
            Clear form
          </button>
        </div>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={6} title="ChatGPT Output" />
        <label>
          Paste output generated in ChatGPT
          <textarea
            className="output-textarea"
            onChange={(event) => updateChatGptOutput(event.target.value)}
            placeholder="Paste the answer from ChatGPT here. Saving will link this output and the prompt to the selected task."
            value={chatGptOutput}
          />
        </label>
        <div className="ai-engine-actions">
          <button className="primary-button" disabled={!activeTask || (!generatedPrompt && !chatGptOutput)} onClick={saveActivity} type="button">
            <FileText size={16} />
            Save AI activity to task
          </button>
        </div>
      </section>

      <section className="panel ai-engine-step">
        <StepHeading number={7} title="Task AI History" />
        <div className="history-list">
          {taskHistory.map((item) => (
            <button className="history-item" key={item.id} onClick={() => loadHistoryItem(item)} type="button">
              <strong>{item.title}</strong>
              <span>{new Date(item.createdAt).toLocaleString()}</span>
              <span>{item.input ? "Source saved" : "No source saved"} - {item.gptPrompt ? "Prompt saved" : "No prompt saved"} - {item.result ? "Output saved" : "No output saved"}</span>
            </button>
          ))}
          {taskHistory.length === 0 && <p className="empty compact-empty">No AI activity saved for this task yet.</p>}
        </div>
      </section>
    </div>
  );
}

function StepHeading({ number, title }: { number: number; title: string }) {
  return (
    <div className="ai-engine-step-heading">
      <span>{number}</span>
      <h3>{title}</h3>
    </div>
  );
}

function StatusPill({ status }: { status: AIUploadedAsset["status"] }) {
  return <span className={`ai-engine-status ai-engine-status-${status}`}>{status}</span>;
}

function buildLinkedTaskSource(task?: WorkTask) {
  if (!task) return "";

  const checklist = task.checklist.length
    ? task.checklist.map((item) => `- [${item.done ? "x" : " "}] ${item.text}`).join("\n")
    : "No checklist items.";
  const sections = task.requirements?.sections?.trim() || "No task sections defined.";
  const constraints = task.requirements?.constraints?.trim() || "No task constraints defined.";

  return [
    "LINKED TASK CONTEXT",
    `Task title: ${task.title}`,
    `Task status: ${task.status}`,
    `Task category: ${task.category}`,
    `Priority: ${task.priority}`,
    `Due date: ${task.dueDate || "None"}`,
    `Task details:\n${task.details.trim() || "No task details provided."}`,
    `Task content goal: ${task.requirements?.outputType || "Not specified"}`,
    `Task audience: ${task.requirements?.audience || "Not specified"}`,
    `Task tone: ${task.requirements?.tone || "Not specified"}`,
    `Task required sections: ${sections}`,
    `Task constraints: ${constraints}`,
    `Task checklist:\n${checklist}`,
  ].join("\n");
}

function buildMasterPrompt(form: AIFormState, sourceMaterial: string) {
  const outputTitle = form.outputTitle.trim() || "Untitled output";
  const additionalInstructions = form.additionalInstructions.trim() || "None provided.";
  const requiredSections = form.sections.length ? form.sections.join(", ") : "Use the best structure for the requested output.";
  const isImagePrompt = form.outputFormat === "Image prompt" || /image generation prompt/i.test(form.outputType);
  const exportInstruction = aiEngineExportInstruction(form.outputFormat);

  return `You are a senior AI generation engine. Your job is to turn messy source material into a high-quality, context-aware final output.

INTENT BRIEF
Objective: Create ${form.outputType.trim() || "the requested output"} for "${outputTitle}".
Desired outcome: A polished, qualitative output that can be used immediately with minimal editing.
Audience: ${form.audience.trim() || "Not specified"}
Tone: ${form.tone.trim() || "Clear and practical"}
${exportInstruction}
Output length: ${form.outputLength}
Required sections: ${requiredSections}
Additional instructions: ${additionalInstructions}

CONTEXT INTELLIGENCE
Before writing, infer the real purpose of the request from the source material, title, audience, and selected template.
Identify the subject, stakeholders, business or personal context, desired decision or action, and any constraints that change the answer.
Prioritize the details that matter to the audience. Remove noise, repetition, and weak filler.
Preserve nuance: distinguish confirmed facts, likely implications, assumptions, risks, gaps, and open questions.
If the source material is thin, produce the best possible output and clearly mark missing information without stalling.

SOURCE MATERIAL RULES
Use only the information provided in the source material.
Do not invent facts, figures, names, dates, commitments, or document contents.
Clearly separate confirmed facts from assumptions where relevant.
If information is missing, state the gap clearly.
Avoid duplication.
Make the output structured, specific, concise, and practical.
Use headings that match the required sections.
Where decisions are needed, list them clearly.
Where next steps are required, make them specific and action-oriented.

QUALITY BAR
The final answer must be useful, not generic.
Every section must earn its place: include concrete content, not template filler.
Use clear headings, strong information hierarchy, and precise wording.
Adapt the level of detail to the requested length.
Check that the output answers the objective, fits the audience, follows the requested format, and respects the constraints.
Do the quality check silently; do not describe the check.

${isImagePrompt ? `IMAGE PROMPT RULES
Create a prompt that can be pasted directly into DALL-E.
Describe the subject, environment, composition, camera/framing, visual style, lighting, color palette, mood, level of realism, and important details.
State any text/logo requirements explicitly. If text must appear in the image, keep it short and quote it exactly.
Include a negative prompt that lists what to avoid.
Do not ask DALL-E to infer business context; translate context into visible details.
Return the image prompt as the primary artifact, plus optional variations if requested by the sections.
` : ""}

SOURCE MATERIAL
${sourceMaterial}

FINAL OUTPUT REQUIRED
Produce the final answer directly.
Do not explain your process.
Do not include generic advice unless requested.
Do not include placeholders unless information is missing or the user requested placeholders.
${isImagePrompt ? "For the primary image prompt, return clean prompt text that is ready to paste into DALL-E." : ""}`;
}

function aiEngineExportInstruction(format: AIOutputFormat) {
  if (format === "DOCX-ready content") {
    return [
      "Formatting/export target: DOCX-ready Markdown.",
      "Return the final answer directly in this chat.",
      "Do not claim to create, attach, export or download a DOCX file.",
      "Structure the content so it can be copied into Word or exported by my app with minimal editing.",
    ].join("\n");
  }

  if (format === "Presentation-ready Markdown") return "Formatting/export target: PPTX-ready slide content.";
  if (format === "Table") return "Formatting/export target: Markdown table or structured table text.";
  if (format === "Email") return "Formatting/export target: Email-ready text.";
  if (format === "Image prompt") return "Formatting/export target: DALL-E-ready image prompt text.";
  if (format === "JSON") return "Formatting/export target: JSON text.";
  return "Formatting/export target: Markdown.";
}

function createId() {
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

function formatBytes(size: number) {
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}
