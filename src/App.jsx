import React, { useEffect, useMemo, useRef, useState } from "react";

const FALLBACK_PROMPTS = [
  {
    id: "0001",
    title: "Daily Priority Map",
    category: "01. Work Planning and Task Prioritization",
    structure: "Role-Based RTCF",
    output: "numbered action plan",
    department: "Operations",
    level: "Intermediate",
    useCase: "Use when organizing daily priorities into a clear action plan.",
    placeholders: "[Source Material], [Department Context], [Target Audience], [Deadline]",
    prompt:
      "Role: You are Dadi Coach Corporation's daily work planning specialist. Help the user convert raw task notes into a prioritized action plan with objective, task sequence, owner, deadline, dependency, and evidence of completion. Do not invent missing details."
  }
];

const THEME_OPTIONS = ["green", "yellow", "clean"];
const VIEW_OPTIONS = [
  { id: "guide", label: "Quick Guide" },
  { id: "library", label: "Prompt Library" },
  { id: "assistant", label: "Prompt Assistant" },
  { id: "improve", label: "Improve Prompt" },
  { id: "upload", label: "Upload Prompts" }
];

const STRUCTURE_OPTIONS = [
  "All structures",
  "Role-Based RTCF",
  "Role-Task-Output-Constraint",
  "CRAFT",
  "RISEN",
  "TAG",
  "Context-Action-Format",
  "Diagnostic-Plan-Deliver",
  "Audit-Insight-Action",
  "SOP Builder",
  "Project Control Prompt"
];

function safeText(value) {
  return String(value ?? "").trim();
}

function escapeHtml(value) {
  return safeText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function getPromptId(item, index, prefix = "") {
  const raw = safeText(item.id || item.promptId || item.number || item["Prompt ID"] || item.ID || item.Id);
  if (raw) return raw;
  return `${prefix}${String(index + 1).padStart(4, "0")}`;
}

function normalizePrompt(item, index = 0, prefix = "") {
  const title = safeText(item.title || item.name || item["Prompt Title"] || item.Title || `Prompt ${index + 1}`);
  const promptText = safeText(item.prompt || item.content || item.template || item.Prompt || item["Full Prompt"] || item["System Prompt"]);
  return {
    id: getPromptId(item, index, prefix),
    title,
    category: safeText(item.category || item.group || item.Category || "User Uploaded Prompts"),
    structure: safeText(item.structure || item.framework || item.Structure || item["Prompt Structure"] || "General"),
    output: safeText(item.output || item.expectedOutput || item.outputFormat || item["Expected Output"] || "ready-to-use prompt"),
    department: safeText(item.department || item.Department || "General"),
    level: safeText(item.level || item.Level || "Intermediate"),
    useCase: safeText(item.useCase || item.bestUseCase || item.description || item["Best Use Case"] || "Use this prompt for a Dadi internal AI task."),
    placeholders: safeText(item.placeholders || item["Customizable Placeholders"] || "[Source Material], [Target Audience], [Deadline], [Required Output]"),
    prompt: promptText.startsWith("Role:") ? promptText : promptText ? `Role: ${promptText}` : "Role: Add your system prompt here."
  };
}

function scorePrompt(prompt, query) {
  const q = query.toLowerCase().trim();
  if (!q) return 1;
  const words = q.split(/\s+/).filter(Boolean);
  const fields = [
    [prompt.title, 9],
    [prompt.category, 7],
    [prompt.structure, 5],
    [prompt.output, 5],
    [prompt.department, 3],
    [prompt.useCase, 3],
    [prompt.placeholders, 2],
    [prompt.prompt, 1]
  ];
  return fields.reduce((score, [field, weight]) => {
    const text = safeText(field).toLowerCase();
    if (!text) return score;
    if (text.includes(q)) return score + weight * 20;
    return score + words.filter((word) => text.includes(word)).length * weight;
  }, 0);
}

function downloadFile(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function downloadExcel(prompts) {
  const columns = ["ID", "Title", "Category", "Structure", "Expected Output", "Department", "Level", "Best Use Case", "Placeholders", "System Prompt"];
  const rows = prompts.map((prompt) => [
    prompt.id,
    prompt.title,
    prompt.category,
    prompt.structure,
    prompt.output,
    prompt.department,
    prompt.level,
    prompt.useCase,
    prompt.placeholders,
    prompt.prompt
  ]);
  const tableRows = [columns, ...rows]
    .map((row) => `<tr>${row.map((cell) => `<td style="mso-number-format:'\\@';">${escapeHtml(cell)}</td>`).join("")}</tr>`)
    .join("");
  const workbook = `<!doctype html><html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="UTF-8"><meta name="ProgId" content="Excel.Sheet"></head><body><table border="1">${tableRows}</table></body></html>`;
  downloadFile("Dadi_Prompt_Library_All_Prompts.xls", workbook, "application/vnd.ms-excel;charset=utf-8");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];
    if (quoted) {
      if (char === '"' && next === '"') {
        cell += '"';
        i += 1;
      } else if (char === '"') {
        quoted = false;
      } else {
        cell += char;
      }
    } else if (char === '"') {
      quoted = true;
    } else if (char === ",") {
      row.push(cell);
      cell = "";
    } else if (char === "\n") {
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
    } else if (char !== "\r") {
      cell += char;
    }
  }
  row.push(cell);
  rows.push(row);
  const cleanRows = rows.filter((item) => item.some((value) => safeText(value)));
  if (cleanRows.length < 2) return [];
  const headers = cleanRows[0].map((header) => safeText(header));
  return cleanRows.slice(1).map((values) => {
    const item = {};
    headers.forEach((header, index) => {
      item[header] = values[index] ?? "";
    });
    return item;
  });
}

function promptScore(text) {
  const content = safeText(text);
  if (!content) return 0;
  let score = 42;
  if (/role:/i.test(content)) score += 10;
  if (/task:/i.test(content)) score += 10;
  if (/context:/i.test(content) || /source/i.test(content)) score += 10;
  if (/format:/i.test(content) || /output/i.test(content)) score += 10;
  if (/constraint|do not|avoid|verify|review/i.test(content)) score += 8;
  if (content.length > 180) score += 5;
  if (content.length > 380) score += 5;
  return Math.min(100, score);
}

function buildImprovedPrompt({ rawPrompt, taskContext, targetAudience, preferredStructure, tone, focus }) {
  const cleanPrompt = safeText(rawPrompt);
  const context = safeText(taskContext) || "[Add source material, project brief, meeting notes, screenshots, data, or reference documents here]";
  const audience = safeText(targetAudience) || "[Target Audience]";
  const style = safeText(tone) || "professional, clear, specific, and suitable for Dadi Coach Corporation";
  const structure = safeText(preferredStructure) || "RTCF";
  const focusLine = safeText(focus) || "clarity, completeness, accurate source alignment, and actionable output";
  if (!cleanPrompt) return "";
  return `Role: You are a senior Dadi Coach Corporation prompt engineering specialist.

Task: Improve the user's rough prompt into a precise, reusable, role-based system prompt using the ${structure} structure.

Original Prompt:
${cleanPrompt}

Context / Source Material:
${context}

Target Audience: ${audience}

Focus: ${focusLine}

Requirements:
1. Preserve the user's original intention.
2. Make the task specific, actionable, and complete.
3. Add clear input requirements, output format, constraints, and review criteria.
4. Use placeholders where information may change, such as [Source Material], [Target Audience], [Program Name], [Country], [Deadline], and [Required Output].
5. Do not invent company facts, prices, names, policies, figures, links, or commitments.
6. Separate internal notes from client-facing wording when needed.
7. Ask up to three clarification questions only if the task cannot be completed accurately.

Output Format:
Return one copy-ready improved prompt with these sections: Role, Task, Context, Constraints, Output Format, Quality Check.

Tone / Style: ${style}`;
}

function CopyIcon({ size = 16 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" aria-hidden="true">
      <path d="M8 8h11v13H8z" stroke="currentColor" strokeWidth="2" />
      <path d="M5 16H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" stroke="currentColor" strokeWidth="2" />
    </svg>
  );
}

function StarIcon({ filled }) {
  return (
    <svg width="18" height="18" viewBox="0 0 24 24" fill={filled ? "currentColor" : "none"} aria-hidden="true">
      <path d="m12 2.5 2.9 5.9 6.5.9-4.7 4.6 1.1 6.5L12 17.3l-5.8 3.1 1.1-6.5-4.7-4.6 6.5-.9L12 2.5Z" stroke="currentColor" strokeWidth="1.8" strokeLinejoin="round" />
    </svg>
  );
}

function SparkIcon() {
  return <span aria-hidden="true">✦</span>;
}

function SidebarContent({ categories, selectedCategory, setSelectedCategory }) {
  return (
    <>
      <div className="sidebarHeader">
        <h2>Categories</h2>
        <p>Browse by task area</p>
      </div>
      <div className="categoryList">
        {categories.map((category) => (
          <button key={category.name} type="button" className={`categoryItem ${selectedCategory === category.name ? "active" : ""}`} onClick={() => setSelectedCategory(category.name)}>
            <span>{category.name.replace(/^\d{2}\.\s*/, "")}</span>
            <strong>{category.count}</strong>
          </button>
        ))}
      </div>
    </>
  );
}

function PromptCard({ prompt, favorites, toggleFavorite, copyPrompt, openPrompt, copyFullCard }) {
  return (
    <article className="promptCard">
      <div className="cardHeader">
        <div>
          <p className="promptNumber">PROMPT {prompt.id}</p>
          <h2>{prompt.title}</h2>
        </div>
        <button type="button" className={`favoriteButton ${favorites.includes(prompt.id) ? "active" : ""}`} onClick={() => toggleFavorite(prompt)} aria-label={favorites.includes(prompt.id) ? "Remove from favorites" : "Add to favorites"}>
          <StarIcon filled={favorites.includes(prompt.id)} />
        </button>
      </div>
      <div className="badgeWrap">
        <span className="badge categoryBadge">{prompt.category.replace(/^\d{2}\.\s*/, "")}</span>
        <span className="badge structureBadge">{prompt.structure}</span>
        <span className="badge outputBadge">{prompt.output}</span>
      </div>
      <p className="useCase">{prompt.useCase}</p>
      <pre className="promptPreview">{prompt.prompt}</pre>
      <div className="cardActions">
        <button type="button" className="primaryAction" onClick={() => copyPrompt(prompt)}><CopyIcon /> Copy Prompt</button>
        <button type="button" className="secondaryAction" onClick={() => openPrompt(prompt)}>View Full</button>
        <button type="button" className="secondaryAction" onClick={() => copyFullCard(prompt)}>Copy Card</button>
      </div>
    </article>
  );
}

export default function App() {
  const [view, setView] = useState("library");
  const [basePrompts, setBasePrompts] = useState(FALLBACK_PROMPTS.map((item, index) => normalizePrompt(item, index)));
  const [uploadedPrompts, setUploadedPrompts] = useState(() => {
    try {
      const stored = JSON.parse(localStorage.getItem("dadi_uploaded_prompts") || "[]");
      return Array.isArray(stored) ? stored.map((item, index) => normalizePrompt(item, index, "U")) : [];
    } catch {
      return [];
    }
  });
  const [loadStatus, setLoadStatus] = useState("Loading official prompt data...");
  const [query, setQuery] = useState("");
  const [assistantQuery, setAssistantQuery] = useState("");
  const [selectedCategory, setSelectedCategory] = useState("All Prompts");
  const [selectedStructure, setSelectedStructure] = useState("All structures");
  const [selectedDepartment, setSelectedDepartment] = useState("All departments");
  const [favoritesOnly, setFavoritesOnly] = useState(false);
  const [selectedPrompt, setSelectedPrompt] = useState(null);
  const [theme, setTheme] = useState("green");
  const [toast, setToast] = useState("");
  const [favorites, setFavorites] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("dadi_favorite_prompts") || "[]");
    } catch {
      return [];
    }
  });
  const [recentlyUsed, setRecentlyUsed] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem("dadi_recent_prompts") || "[]");
    } catch {
      return [];
    }
  });
  const [improver, setImprover] = useState({
    rawPrompt: "Create a work plan about employees' productivity",
    taskContext: "",
    targetAudience: "Dadi Coach employees",
    preferredStructure: "RTCF",
    tone: "Professional, clear, and management-ready",
    focus: "priorities, KPIs, timelines, action steps"
  });
  const [improvedPrompt, setImprovedPrompt] = useState(() => buildImprovedPrompt({
    rawPrompt: "Create a work plan about employees' productivity",
    taskContext: "",
    targetAudience: "Dadi Coach employees",
    preferredStructure: "RTCF",
    tone: "Professional, clear, and management-ready",
    focus: "priorities, KPIs, timelines, action steps"
  }));
  const fileInputRef = useRef(null);

  const prompts = useMemo(() => [...basePrompts, ...uploadedPrompts], [basePrompts, uploadedPrompts]);

  useEffect(() => {
    let active = true;
    const version = Date.now();
    fetch(`/prompts.json?v=${version}`, { cache: "no-store" })
      .then((response) => {
        if (!response.ok) throw new Error("Prompt file not found");
        return response.json();
      })
      .then((data) => {
        if (!active) return;
        const list = Array.isArray(data) ? data : Array.isArray(data.prompts) ? data.prompts : [];
        if (list.length > 0) {
          setBasePrompts(list.map((item, index) => normalizePrompt(item, index)));
          setLoadStatus(`${list.length.toLocaleString()} official prompts loaded`);
        } else {
          setLoadStatus("Prompt file loaded, but no valid prompt records were found.");
        }
      })
      .catch((error) => setLoadStatus(`Using fallback prompt data. ${error.message}`));
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => localStorage.setItem("dadi_favorite_prompts", JSON.stringify(favorites)), [favorites]);
  useEffect(() => localStorage.setItem("dadi_recent_prompts", JSON.stringify(recentlyUsed)), [recentlyUsed]);
  useEffect(() => localStorage.setItem("dadi_uploaded_prompts", JSON.stringify(uploadedPrompts)), [uploadedPrompts]);
  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 2300);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const categories = useMemo(() => {
    const map = new Map();
    prompts.forEach((prompt) => map.set(prompt.category, (map.get(prompt.category) || 0) + 1));
    return [
      { name: "All Prompts", count: prompts.length },
      ...Array.from(map.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([name, count]) => ({ name, count }))
    ];
  }, [prompts]);

  const structures = useMemo(() => {
    const values = new Set(STRUCTURE_OPTIONS);
    prompts.forEach((prompt) => values.add(prompt.structure));
    return Array.from(values).filter(Boolean);
  }, [prompts]);

  const departments = useMemo(() => {
    const values = new Set(["All departments"]);
    prompts.forEach((prompt) => values.add(prompt.department));
    return Array.from(values).filter(Boolean).sort((a, b) => {
      if (a === "All departments") return -1;
      if (b === "All departments") return 1;
      return a.localeCompare(b);
    });
  }, [prompts]);

  const filteredPrompts = useMemo(() => {
    return prompts
      .filter((prompt) => {
        const matchesQuery = !query || scorePrompt(prompt, query) > 0;
        const matchesCategory = selectedCategory === "All Prompts" || prompt.category === selectedCategory;
        const matchesStructure = selectedStructure === "All structures" || prompt.structure === selectedStructure;
        const matchesDepartment = selectedDepartment === "All departments" || prompt.department === selectedDepartment;
        const matchesFavorite = !favoritesOnly || favorites.includes(prompt.id);
        return matchesQuery && matchesCategory && matchesStructure && matchesDepartment && matchesFavorite;
      })
      .sort((a, b) => scorePrompt(b, query) - scorePrompt(a, query));
  }, [prompts, query, selectedCategory, selectedStructure, selectedDepartment, favoritesOnly, favorites]);

  const assistantMatches = useMemo(() => {
    const searchTerm = assistantQuery || query;
    return prompts
      .map((prompt) => ({ prompt, score: scorePrompt(prompt, searchTerm) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 8)
      .map((item) => item.prompt);
  }, [prompts, assistantQuery, query]);

  const recentPromptCards = recentlyUsed.map((id) => prompts.find((prompt) => prompt.id === id)).filter(Boolean).slice(0, 4);
  const improvedScore = promptScore(improvedPrompt);

  function showToast(message) {
    setToast(message);
  }

  async function copyText(text, label = "Copied") {
    try {
      await navigator.clipboard.writeText(text);
      showToast(label);
    } catch {
      const textArea = document.createElement("textarea");
      textArea.value = text;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      textArea.remove();
      showToast(label);
    }
  }

  function rememberPrompt(id) {
    setRecentlyUsed((current) => [id, ...current.filter((item) => item !== id)].slice(0, 8));
  }

  function copyPrompt(prompt) {
    rememberPrompt(prompt.id);
    copyText(prompt.prompt, "Prompt copied successfully");
  }

  function copyFullCard(prompt) {
    const content = [
      `PROMPT ${prompt.id}`,
      prompt.title,
      `Category: ${prompt.category}`,
      `Structure: ${prompt.structure}`,
      `Expected Output: ${prompt.output}`,
      `Department: ${prompt.department}`,
      `Level: ${prompt.level}`,
      `Best Use Case: ${prompt.useCase}`,
      `Placeholders: ${prompt.placeholders}`,
      "",
      prompt.prompt
    ].join("\n");
    rememberPrompt(prompt.id);
    copyText(content, "Full card copied successfully");
  }

  function openPrompt(prompt) {
    rememberPrompt(prompt.id);
    setSelectedPrompt(prompt);
  }

  function toggleFavorite(prompt) {
    setFavorites((current) => current.includes(prompt.id) ? current.filter((item) => item !== prompt.id) : [...current, prompt.id]);
  }

  function resetFilters() {
    setQuery("");
    setSelectedCategory("All Prompts");
    setSelectedStructure("All structures");
    setSelectedDepartment("All departments");
    setFavoritesOnly(false);
  }

  function handleDownloadAllPrompts() {
    downloadExcel(prompts);
    showToast("Excel file downloaded");
  }

  function handleDownloadJson() {
    downloadFile("Dadi_Prompt_Library_Backup.json", JSON.stringify({ exportedAt: new Date().toISOString(), totalPrompts: prompts.length, prompts }, null, 2), "application/json;charset=utf-8");
    showToast("JSON backup downloaded");
  }

  function handleImprovePrompt() {
    const result = buildImprovedPrompt(improver);
    setImprovedPrompt(result);
    if (result) showToast("Improved prompt generated");
  }

  function handleSaveImprovedPrompt() {
    if (!improvedPrompt) return;
    const title = safeText(improver.rawPrompt).slice(0, 55) || "Improved Prompt";
    const item = normalizePrompt({
      id: `U${String(uploadedPrompts.length + 1).padStart(4, "0")}`,
      title: `${title}${title.length >= 55 ? "..." : ""}`,
      category: "User Improved Prompts",
      structure: improver.preferredStructure || "RTCF",
      output: "copy-ready improved system prompt",
      department: "General",
      level: "Intermediate",
      useCase: "Saved from Prompt Improvement Studio.",
      placeholders: "[Source Material], [Target Audience], [Program Name], [Deadline], [Required Output]",
      prompt: improvedPrompt
    }, uploadedPrompts.length, "U");
    setUploadedPrompts((current) => [item, ...current]);
    setView("library");
    showToast("Improved prompt saved to library");
  }

  function handleUploadClick() {
    fileInputRef.current?.click();
  }

  function handleFileUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = String(reader.result || "");
        let rawItems = [];
        if (file.name.toLowerCase().endsWith(".json")) {
          const data = JSON.parse(text);
          rawItems = Array.isArray(data) ? data : Array.isArray(data.prompts) ? data.prompts : [];
        } else {
          rawItems = parseCsv(text);
        }
        const normalized = rawItems
          .map((item, index) => normalizePrompt(item, uploadedPrompts.length + index, "U"))
          .filter((item) => safeText(item.title) && safeText(item.prompt));
        if (!normalized.length) {
          showToast("No valid prompts found in the uploaded file");
          return;
        }
        setUploadedPrompts((current) => [...current, ...normalized]);
        setView("library");
        showToast(`${normalized.length} uploaded prompt(s) added`);
      } catch {
        showToast("Upload failed. Use JSON or CSV with prompt columns.");
      } finally {
        event.target.value = "";
      }
    };
    reader.readAsText(file);
  }

  function clearUploadedPrompts() {
    if (!uploadedPrompts.length) return;
    const confirmed = window.confirm("Remove all prompts uploaded in this browser?");
    if (!confirmed) return;
    setUploadedPrompts([]);
    showToast("Uploaded prompts cleared");
  }

  function clearLocalData() {
    localStorage.removeItem("dadi_uploaded_prompts");
    localStorage.removeItem("dadi_favorite_prompts");
    localStorage.removeItem("dadi_recent_prompts");
    setUploadedPrompts([]);
    setFavorites([]);
    setRecentlyUsed([]);
    showToast("Local saved data cleared");
  }

  return (
    <div className={`dadiApp theme-${theme}`}>
      <style>{styles}</style>
      <header className="topbar">
        <button type="button" className="brandGroup" onClick={() => setView("library")} aria-label="Dadi Prompt Library home">
          <div className="brandLogo"><img src="/dadi-coach-logo.png" alt="Dadi Coach Logo" /></div>
          <div>
            <h1>Dadi Prompt Library</h1>
            <p>Internal AI prompt templates</p>
          </div>
        </button>
        <nav className="navPills" aria-label="Main navigation">
          {VIEW_OPTIONS.map((item) => (
            <button key={item.id} type="button" className={view === item.id ? "active" : ""} onClick={() => setView(item.id)}>{item.label}</button>
          ))}
          <button type="button" className="downloadButton" onClick={handleDownloadAllPrompts}>Download All Prompts</button>
          <input ref={fileInputRef} type="file" accept=".json,.csv,text/csv,application/json" onChange={handleFileUpload} hidden />
        </nav>
        <div className="themeSwitch" aria-label="Theme selector">
          {THEME_OPTIONS.map((item) => (
            <button key={item} type="button" className={theme === item ? "active" : ""} onClick={() => setTheme(item)}>{item[0].toUpperCase() + item.slice(1)}</button>
          ))}
        </div>
      </header>

      <main className="pageShell">
        {view !== "library" && <button type="button" className="backLink" onClick={() => setView("library")}>← Back to Prompt Library</button>}
        {view === "guide" && (
          <GuideView prompts={prompts} loadStatus={loadStatus} setView={setView} clearLocalData={clearLocalData} />
        )}
        {view === "library" && (
          <LibraryView
            prompts={prompts}
            basePrompts={basePrompts}
            uploadedPrompts={uploadedPrompts}
            loadStatus={loadStatus}
            query={query}
            setQuery={setQuery}
            selectedCategory={selectedCategory}
            setSelectedCategory={setSelectedCategory}
            selectedStructure={selectedStructure}
            setSelectedStructure={setSelectedStructure}
            selectedDepartment={selectedDepartment}
            setSelectedDepartment={setSelectedDepartment}
            favoritesOnly={favoritesOnly}
            setFavoritesOnly={setFavoritesOnly}
            categories={categories}
            structures={structures}
            departments={departments}
            filteredPrompts={filteredPrompts}
            recentPromptCards={recentPromptCards}
            favorites={favorites}
            toggleFavorite={toggleFavorite}
            copyPrompt={copyPrompt}
            openPrompt={openPrompt}
            copyFullCard={copyFullCard}
            resetFilters={resetFilters}
            setView={setView}
          />
        )}
        {view === "assistant" && (
          <AssistantView assistantQuery={assistantQuery} setAssistantQuery={setAssistantQuery} assistantMatches={assistantMatches} copyPrompt={copyPrompt} openPrompt={openPrompt} setView={setView} />
        )}
        {view === "improve" && (
          <ImproveView
            improver={improver}
            setImprover={setImprover}
            improvedPrompt={improvedPrompt}
            setImprovedPrompt={setImprovedPrompt}
            improvedScore={improvedScore}
            handleImprovePrompt={handleImprovePrompt}
            copyText={copyText}
            handleSaveImprovedPrompt={handleSaveImprovedPrompt}
          />
        )}
        {view === "upload" && (
          <UploadView uploadedPrompts={uploadedPrompts} handleUploadClick={handleUploadClick} handleDownloadJson={handleDownloadJson} clearUploadedPrompts={clearUploadedPrompts} clearLocalData={clearLocalData} />
        )}
      </main>

      {selectedPrompt && (
        <div className="modalOverlay" role="dialog" aria-modal="true">
          <div className="promptModal">
            <div className="modalHeader">
              <div>
                <p className="promptNumber">PROMPT {selectedPrompt.id}</p>
                <h2>{selectedPrompt.title}</h2>
              </div>
              <button type="button" className="iconClose" onClick={() => setSelectedPrompt(null)}>×</button>
            </div>
            <div className="modalMeta">
              <span>{selectedPrompt.category}</span>
              <span>{selectedPrompt.structure}</span>
              <span>{selectedPrompt.output}</span>
              <span>{selectedPrompt.department}</span>
            </div>
            <p className="modalUseCase"><strong>Best use case:</strong> {selectedPrompt.useCase}</p>
            <p className="modalPlaceholders"><strong>Placeholders:</strong> {selectedPrompt.placeholders}</p>
            <pre className="modalPromptText">{selectedPrompt.prompt}</pre>
            <div className="modalActions">
              <button type="button" className="primaryAction" onClick={() => copyPrompt(selectedPrompt)}><CopyIcon /> Copy Prompt</button>
              <button type="button" className="secondaryAction" onClick={() => copyFullCard(selectedPrompt)}>Copy Full Card</button>
            </div>
          </div>
        </div>
      )}
      <button type="button" className="floatingAssistant" onClick={() => setView("assistant")} aria-label="Open Prompt Assistant">♟</button>
      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}

function GuideView({ prompts, loadStatus, setView, clearLocalData }) {
  return (
    <section className="guideView">
      <div className="sectionTitleRow">
        <div>
          <p className="eyebrow dark">Quick Prompt Selection Guide</p>
          <h2>Choose prompts by task type, not by department.</h2>
          <p>Employees can search by keyword, filter by category, copy a prompt, replace placeholders, and attach source materials before using ChatGPT.</p>
        </div>
        <div className="scoreBubble">{prompts.length}<span>prompts available</span></div>
      </div>
      <div className="guideGrid">
        <div className="guideCard"><h3>1. Search</h3><p>Use task keywords such as meeting summary, work plan, proposal, poster, CRM, or teacher evaluation.</p></div>
        <div className="guideCard"><h3>2. Copy</h3><p>Copy the full Role-based prompt and replace placeholders like [Source Material], [Country], and [Deadline].</p></div>
        <div className="guideCard"><h3>3. Improve</h3><p>Paste a rough prompt in the Improvement Studio to convert it into a stronger system prompt.</p></div>
        <div className="guideCard"><h3>4. Upload</h3><p>Upload your own JSON or CSV prompt list. Uploaded prompts are saved only in this browser.</p></div>
      </div>
      <div className="statusCard"><strong>Prompt data status:</strong> {loadStatus}<button type="button" className="secondaryAction" onClick={clearLocalData}>Clear Browser Saved Data</button><button type="button" className="primaryAction" onClick={() => setView("library")}>Browse Library</button></div>
    </section>
  );
}

function LibraryView(props) {
  const {
    prompts,
    basePrompts,
    uploadedPrompts,
    loadStatus,
    query,
    setQuery,
    selectedCategory,
    setSelectedCategory,
    selectedStructure,
    setSelectedStructure,
    selectedDepartment,
    setSelectedDepartment,
    favoritesOnly,
    setFavoritesOnly,
    categories,
    structures,
    departments,
    filteredPrompts,
    recentPromptCards,
    favorites,
    toggleFavorite,
    copyPrompt,
    openPrompt,
    copyFullCard,
    resetFilters,
    setView
  } = props;
  return (
    <>
      <section className="libraryHero">
        <div>
          <p className="eyebrow dark">Prompt Library</p>
          <h2>Search, copy, improve, and organize Dadi system prompts.</h2>
          <p>{loadStatus}. The library includes official prompts and any prompts uploaded in this browser.</p>
        </div>
        <div className="verificationCard">
          <p>Library Verification</p>
          <strong>{basePrompts.length.toLocaleString()}</strong>
          <span>official prompts</span>
          <small>{uploadedPrompts.length} uploaded • {prompts.length.toLocaleString()} total</small>
        </div>
      </section>

      <section className="searchPanel" aria-label="Prompt search and filters">
        <div className="searchRow primarySearchRow">
          <label className="searchInputWrap">
            <span className="srOnly">Search prompts</span>
            <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder='Search by task, role, output, or keyword - e.g. "meeting summary", "sales reply", "RTCF"' />
          </label>
          <button type="button" className="askButton" onClick={() => setView("assistant")}>Ask Assistant</button>
        </div>
        <div className="searchRow filterRow">
          <select value={selectedCategory} onChange={(event) => setSelectedCategory(event.target.value)}>{categories.map((category) => <option key={category.name} value={category.name}>{category.name}</option>)}</select>
          <select value={selectedStructure} onChange={(event) => setSelectedStructure(event.target.value)}>{structures.map((structure) => <option key={structure} value={structure}>{structure}</option>)}</select>
          <select value={selectedDepartment} onChange={(event) => setSelectedDepartment(event.target.value)}>{departments.map((department) => <option key={department} value={department}>{department}</option>)}</select>
          <button type="button" className={`toggleButton ${favoritesOnly ? "active" : ""}`} onClick={() => setFavoritesOnly((value) => !value)}>Favorites Only</button>
          <button type="button" className="clearButton" onClick={resetFilters}>Reset</button>
        </div>
        <div className="resultSummary"><strong>{filteredPrompts.length}</strong> prompts found <span>{prompts.length.toLocaleString()} total prompts</span><span>{Math.max(categories.length - 1, 0)} categories</span><span>{uploadedPrompts.length} uploaded</span></div>
      </section>

      <div className="contentLayout">
        <aside className="sidebar"><SidebarContent categories={categories} selectedCategory={selectedCategory} setSelectedCategory={setSelectedCategory} /></aside>
        <section className="promptArea">
          {recentPromptCards.length > 0 && (
            <div className="recentPanel">
              <div><h2>Recently Used</h2><p>Fast access to prompts you copied or opened recently.</p></div>
              <div className="recentList">{recentPromptCards.map((prompt) => <button key={prompt.id} type="button" onClick={() => openPrompt(prompt)}>{prompt.title}</button>)}</div>
            </div>
          )}
          <div className="promptGrid">
            {filteredPrompts.map((prompt) => <PromptCard key={`${prompt.id}-${prompt.title}`} prompt={prompt} favorites={favorites} toggleFavorite={toggleFavorite} copyPrompt={copyPrompt} openPrompt={openPrompt} copyFullCard={copyFullCard} />)}
          </div>
          {!filteredPrompts.length && <div className="emptyState"><h2>No prompts found</h2><p>Try another keyword, category, or structure filter.</p><button type="button" onClick={resetFilters}>Reset Filters</button></div>}
        </section>
      </div>
    </>
  );
}

function AssistantView({ assistantQuery, setAssistantQuery, assistantMatches, copyPrompt, openPrompt, setView }) {
  return (
    <section className="assistantView">
      <div className="sectionTitleRow compact">
        <div>
          <p className="eyebrow dark">Prompt Assistant</p>
          <h2>Tell the assistant what you need, then choose the closest prompt.</h2>
          <p>Example: “I need a prompt to evaluate a teacher's recorded class video with strengths and areas to improve.”</p>
        </div>
      </div>
      <div className="assistantWorkspace">
        <div className="inputPanel">
          <h3>What task do you need help with?</h3>
          <textarea value={assistantQuery} onChange={(event) => setAssistantQuery(event.target.value)} placeholder="Describe your task, audience, output, and source material..." rows={10} />
          <div className="toolActions"><button type="button" className="primaryAction">Search Prompt Library</button><button type="button" className="secondaryAction" onClick={() => setView("improve")}>Improve My Own Prompt</button></div>
        </div>
        <div className="resultPanel">
          <h3>Recommended Prompts</h3>
          {!assistantQuery && <p className="helperText">Type a task description to see the most relevant prompts.</p>}
          {assistantMatches.map((prompt) => (
            <div className="assistantResult" key={prompt.id}>
              <p>PROMPT {prompt.id}</p>
              <h4>{prompt.title}</h4>
              <span>{prompt.category} • {prompt.structure}</span>
              <div className="toolActions"><button type="button" className="primaryAction" onClick={() => copyPrompt(prompt)}>Copy</button><button type="button" className="secondaryAction" onClick={() => openPrompt(prompt)}>View</button></div>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}

function ImproveView({ improver, setImprover, improvedPrompt, setImprovedPrompt, improvedScore, handleImprovePrompt, copyText, handleSaveImprovedPrompt }) {
  return (
    <section className="improveView">
      <div className="sectionTitleRow compact">
        <div>
          <h2>Improve Your Prompt</h2>
          <p>Refine and enhance your prompt to make it clearer, more specific, and more effective.</p>
        </div>
        <div className="titleActions"><button type="button" className="outlineAction">💡 Prompt Improvement Guide</button><button type="button" className="outlineAction">↻ History</button></div>
      </div>
      <div className="improveGrid">
        <div className="improvePanel">
          <h3>1. Enter Your Original Prompt</h3>
          <label>Paste or type the prompt you want to improve.<span>{improver.rawPrompt.length}/4000</span></label>
          <textarea value={improver.rawPrompt} onChange={(event) => setImprover((current) => ({ ...current, rawPrompt: event.target.value }))} rows={6} placeholder="Paste your rough prompt here..." />
          <div className="miniFormGrid">
            <label>Task Context / Source Material <input value={improver.taskContext} onChange={(event) => setImprover((current) => ({ ...current, taskContext: event.target.value }))} placeholder="e.g., Project brief, meeting notes, documents..." /></label>
            <label>Target Audience <input value={improver.targetAudience} onChange={(event) => setImprover((current) => ({ ...current, targetAudience: event.target.value }))} /></label>
            <label>Output Format <select value={improver.preferredStructure} onChange={(event) => setImprover((current) => ({ ...current, preferredStructure: event.target.value }))}><option>RTCF</option><option>Role-Task-Output-Constraint</option><option>CRAFT</option><option>RISEN</option><option>TAG</option><option>SOP Builder</option></select></label>
            <label>Tone / Style <input value={improver.tone} onChange={(event) => setImprover((current) => ({ ...current, tone: event.target.value }))} /></label>
            <label>What should the prompt focus on? <textarea value={improver.focus} onChange={(event) => setImprover((current) => ({ ...current, focus: event.target.value }))} rows={4} placeholder="e.g., priorities, KPIs, timelines, action steps..." /></label>
            <label>Additional Instructions <textarea rows={4} placeholder="e.g., include examples, use metrics..." /></label>
          </div>
          <div className="panelFooter"><button type="button" className="secondaryAction" onClick={() => setImprover({ rawPrompt: "", taskContext: "", targetAudience: "Dadi Coach employees", preferredStructure: "RTCF", tone: "Professional, clear, and management-ready", focus: "" })}>🗑 Clear All</button><button type="button" className="primaryAction" onClick={handleImprovePrompt}>✨ Improve Prompt</button></div>
        </div>
        <div className="improvePanel resultImprovePanel">
          <div className="resultHeader"><div><h3>2. Improved Prompt <span>Optimized</span></h3><p>Review the improved prompt below. You can edit, copy, or save it.</p></div><div className="promptScore"><strong>{improvedScore}</strong><span>/100</span></div></div>
          <div className="codeBox">
            <button type="button" onClick={() => copyText(improvedPrompt, "Improved prompt copied")}><CopyIcon /> Copy</button>
            <textarea value={improvedPrompt} onChange={(event) => setImprovedPrompt(event.target.value)} />
          </div>
          <p className="tipText">ⓘ Tip: Review the improved prompt and ensure it aligns with your intent and context.</p>
          <div className="panelFooter"><button type="button" className="secondaryAction" onClick={() => copyText(improvedPrompt, "Improved prompt copied")}>✎ Copy Improved Prompt</button><button type="button" className="primaryAction" onClick={handleSaveImprovedPrompt}>🔖 Save to Library</button></div>
        </div>
      </div>
      <div className="bottomGrid">
        <div className="breakdownPanel"><h3>Improvement Breakdown</h3><div className="breakdownList"><span>Clarity<br/><small>Improved intent and specificity</small></span><span>Structure<br/><small>Added clear sections</small></span><span>Context<br/><small>Included details and considerations</small></span><span>Actionability<br/><small>Added steps and requirements</small></span><span>Format<br/><small>Specified output format</small></span></div></div>
        <div className="breakdownPanel"><h3>Related Suggestions</h3><div className="suggestionTags"><span>Employee Productivity Plan</span><span>Performance Improvement Strategy</span><span>Work Plan Template</span><span>KPI Tracking Plan</span><span>Productivity Action Plan</span></div></div>
      </div>
    </section>
  );
}

function UploadView({ uploadedPrompts, handleUploadClick, handleDownloadJson, clearUploadedPrompts, clearLocalData }) {
  return (
    <section className="uploadView">
      <div className="sectionTitleRow compact"><div><p className="eyebrow dark">Prompt Upload</p><h2>Add your own prompt templates</h2><p>Upload a JSON or CSV file with columns such as Title, Category, Structure, Expected Output, Best Use Case, and Prompt.</p></div></div>
      <div className="uploadGrid">
        <div className="uploadDrop"><div className="uploadIcon">⬆</div><h3>Upload JSON / CSV</h3><p>Uploaded prompts are saved only in this browser and can be exported as backup.</p><button type="button" className="primaryAction" onClick={handleUploadClick}>Upload File</button></div>
        <div className="uploadInfo"><h3>Upload format</h3><p>Recommended CSV columns:</p><pre>ID,Title,Category,Structure,Expected Output,Department,Level,Best Use Case,Tags,Prompt</pre><div className="toolActions"><button type="button" className="secondaryAction" onClick={handleDownloadJson}>Download JSON Backup</button><button type="button" className="secondaryAction" onClick={clearUploadedPrompts}>Clear Uploaded</button><button type="button" className="secondaryAction" onClick={clearLocalData}>Clear Browser Data</button></div></div>
      </div>
      <div className="statusCard"><strong>{uploadedPrompts.length}</strong> uploaded prompt(s) saved in this browser.</div>
    </section>
  );
}

const styles = `
:root {
  --dadi-green: #08743f;
  --dadi-green-dark: #00391f;
  --dadi-green-deep: #002f1c;
  --dadi-green-soft: #e9f6ee;
  --dadi-yellow: #ffba08;
  --dadi-yellow-soft: #fff3c7;
  --dadi-orange: #f48a00;
  --text-main: #0b2f22;
  --text-muted: #61766e;
  --border-soft: #dce8df;
  --bg-page: #fbfdf8;
  --shadow-soft: 0 18px 48px rgba(6, 43, 24, .08);
  --shadow-card: 0 24px 65px rgba(6, 43, 24, .12);
}
* { box-sizing: border-box; }
body { margin: 0; font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; background: var(--bg-page); color: var(--text-main); }
button, input, select, textarea { font: inherit; }
button { cursor: pointer; }
.dadiApp { min-height: 100vh; background: radial-gradient(circle at top right, rgba(255,186,8,.10), transparent 34%), linear-gradient(180deg,#ffffff 0%,#fbfdf8 38%,#f7fbf2 100%); }
.theme-yellow { --dadi-green: #6d8d00; --dadi-green-dark: #4d4a00; --dadi-green-deep: #3d3b00; --dadi-green-soft: #fff8d8; --dadi-yellow: #f5a900; --dadi-orange: #f07800; }
.theme-clean { --dadi-green: #3f684f; --dadi-green-dark: #123223; --dadi-green-deep: #13251d; --dadi-green-soft: #f2f6f2; --dadi-yellow: #ece6cf; --dadi-orange: #9d7a25; --bg-page: #ffffff; }
.topbar { position: sticky; top: 0; z-index: 50; min-height: 86px; padding: 14px 28px; display: grid; grid-template-columns: auto 1fr auto; gap: 18px; align-items: center; background: rgba(255,255,255,.88); border-bottom: 1px solid var(--border-soft); backdrop-filter: blur(18px); box-shadow: 0 8px 26px rgba(4, 42, 22, .05); }
.brandGroup { border: 0; background: transparent; display: flex; align-items: center; gap: 14px; color: inherit; text-align: left; }
.brandLogo { width: 132px; height: 54px; border: 1px solid var(--border-soft); border-radius: 999px; background: #fff; display: grid; place-items: center; box-shadow: 0 12px 34px rgba(6,43,24,.09); overflow: hidden; }
.brandLogo img { width: 118px; height: auto; display: block; }
.brandGroup h1 { margin: 0; font-size: 20px; line-height: 1.1; color: var(--dadi-green-dark); font-weight: 950; letter-spacing: -.02em; }
.brandGroup p { margin: 5px 0 0; color: var(--text-muted); font-size: 13px; }
.navPills { display: flex; align-items: center; justify-content: center; gap: 10px; flex-wrap: wrap; }
.navPills button, .themeSwitch button { border: 1px solid var(--border-soft); background: #fff; color: var(--dadi-green-dark); border-radius: 999px; padding: 12px 20px; font-weight: 900; box-shadow: 0 8px 22px rgba(6,43,24,.04); }
.navPills button.active, .themeSwitch button.active { background: var(--dadi-green); color: #fff; border-color: var(--dadi-green); }
.navPills .downloadButton { background: var(--dadi-yellow); color: #17200a; border-color: var(--dadi-yellow); }
.themeSwitch { display: flex; gap: 8px; }
.pageShell { width: min(1680px, calc(100% - 56px)); margin: 0 auto; padding: 22px 0 64px; }
.backLink { border: 0; background: transparent; color: var(--dadi-green-dark); font-weight: 850; margin: 0 0 22px; }
.eyebrow { margin: 0 0 10px; color: var(--dadi-yellow); font-weight: 950; letter-spacing: .12em; text-transform: uppercase; font-size: 12px; }
.eyebrow.dark { color: var(--dadi-green); }
.libraryHero, .sectionTitleRow, .searchPanel, .sidebar, .promptCard, .recentPanel, .improvePanel, .breakdownPanel, .uploadDrop, .uploadInfo, .statusCard, .guideCard, .inputPanel, .resultPanel { background: rgba(255,255,255,.96); border: 1px solid var(--border-soft); border-radius: 24px; box-shadow: var(--shadow-soft); }
.libraryHero { display: grid; grid-template-columns: 1fr 320px; gap: 24px; align-items: center; padding: 34px; margin-bottom: 22px; background: linear-gradient(135deg,#fff 0%,#f8fcf5 70%,rgba(255,186,8,.12)); }
.libraryHero h2, .sectionTitleRow h2 { margin: 0; color: var(--dadi-green-dark); font-size: clamp(32px,4vw,58px); line-height: 1.03; letter-spacing: -.04em; }
.libraryHero p, .sectionTitleRow p { color: var(--text-muted); font-size: 16px; line-height: 1.7; max-width: 900px; }
.verificationCard { border-radius: 22px; padding: 26px; background: linear-gradient(135deg,var(--dadi-green),#5a9349); color: #fff; box-shadow: var(--shadow-card); }
.verificationCard p { color: var(--dadi-yellow); margin: 0 0 12px; text-transform: uppercase; letter-spacing: .08em; font-weight: 950; font-size: 12px; }
.verificationCard strong { display: block; font-size: 60px; line-height: 1; }
.verificationCard span, .verificationCard small { display: block; margin-top: 9px; color: rgba(255,255,255,.88); }
.searchPanel { padding: 22px; margin-bottom: 22px; }
.searchRow { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.primarySearchRow { display: grid; grid-template-columns: 1fr auto; }
.searchInputWrap input, .filterRow select, .inputPanel textarea, .improvePanel textarea, .improvePanel input, .improvePanel select { width: 100%; border: 1px solid #d9e4dc; background: #fff; color: var(--text-main); border-radius: 16px; outline: none; transition: border-color 160ms ease, box-shadow 160ms ease; }
.searchInputWrap input { min-height: 58px; padding: 0 18px; font-size: 16px; }
.searchInputWrap input:focus, .filterRow select:focus, textarea:focus, input:focus, select:focus { border-color: var(--dadi-green); box-shadow: 0 0 0 4px rgba(13,107,54,.12); }
.askButton, .primaryAction { border: 0; background: var(--dadi-yellow); color: #102015; border-radius: 16px; padding: 15px 22px; font-weight: 950; box-shadow: 0 12px 24px rgba(255,184,0,.22); display: inline-flex; align-items: center; justify-content: center; gap: 8px; }
.filterRow { margin-top: 14px; }
.filterRow select { flex: 1 1 220px; min-height: 48px; padding: 0 14px; }
.toggleButton, .clearButton, .secondaryAction, .outlineAction { border: 1px solid var(--border-soft); background: #fff; color: var(--text-main); border-radius: 13px; padding: 12px 15px; font-weight: 850; display: inline-flex; align-items: center; justify-content: center; gap: 8px; }
.toggleButton.active { color: #fff; background: var(--dadi-green); border-color: var(--dadi-green); }
.resultSummary { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-top: 16px; color: var(--text-muted); font-size: 14px; }
.resultSummary strong { display: inline-grid; place-items: center; min-width: 40px; min-height: 30px; padding: 2px 10px; background: var(--dadi-green); color: #fff; border-radius: 999px; }
.resultSummary span::before { content: ""; display: inline-block; width: 5px; height: 5px; border-radius: 50%; background: #9aab9f; margin-right: 10px; vertical-align: 2px; }
.contentLayout { display: grid; grid-template-columns: 310px 1fr; gap: 26px; align-items: start; }
.sidebar { position: sticky; top: 108px; max-height: calc(100vh - 132px); overflow: auto; padding: 18px; }
.sidebarHeader h2 { margin: 0 0 4px; font-size: 20px; }
.sidebarHeader p { margin: 0 0 14px; color: var(--text-muted); font-size: 13px; }
.categoryList { display: grid; gap: 8px; }
.categoryItem { border: 0; background: transparent; color: var(--text-main); border-radius: 14px; padding: 12px; text-align: left; display: grid; grid-template-columns: 1fr auto; gap: 10px; align-items: center; font-weight: 850; }
.categoryItem strong { min-width: 28px; height: 24px; display: inline-grid; place-items: center; background: #edf3ed; border-radius: 999px; color: var(--text-muted); font-size: 12px; }
.categoryItem.active { background: var(--dadi-green-soft); box-shadow: inset 4px 0 0 var(--dadi-green); }
.promptArea { min-width: 0; }
.recentPanel { display: flex; align-items: center; justify-content: space-between; gap: 18px; padding: 16px 20px; margin-bottom: 20px; }
.recentPanel h2 { margin: 0; font-size: 18px; }
.recentPanel p { margin: 4px 0 0; color: var(--text-muted); font-size: 13px; }
.recentList { display: flex; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
.recentList button, .suggestionTags span { border: 1px solid var(--border-soft); background: var(--dadi-green-soft); color: var(--dadi-green-dark); border-radius: 999px; padding: 9px 12px; font-weight: 850; }
.promptGrid { display: grid; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); gap: 20px; }
.promptCard { padding: 22px; display: flex; flex-direction: column; min-height: 560px; }
.cardHeader { display: flex; justify-content: space-between; gap: 16px; }
.promptNumber { margin: 0 0 8px; color: var(--dadi-orange); font-size: 12px; font-weight: 950; letter-spacing: .08em; }
.promptCard h2 { margin: 0; color: var(--dadi-green-dark); font-size: 24px; line-height: 1.15; }
.favoriteButton { flex: 0 0 auto; border: 0; color: #b98900; background: var(--dadi-yellow-soft); width: 46px; height: 46px; border-radius: 16px; display: grid; place-items: center; }
.badgeWrap { display: flex; gap: 8px; flex-wrap: wrap; margin: 18px 0 12px; }
.badge { display: inline-flex; align-items: center; border-radius: 999px; padding: 6px 10px; font-size: 12px; font-weight: 850; }
.categoryBadge { background: var(--dadi-green-soft); color: var(--dadi-green-dark); }
.structureBadge { background: var(--dadi-yellow-soft); color: #5c4100; }
.outputBadge { background: #f2f6f2; color: var(--text-muted); }
.useCase { color: var(--text-muted); line-height: 1.55; font-size: 14px; min-height: 66px; }
.promptPreview, .modalPromptText { flex: 1; background: radial-gradient(circle at top right,rgba(255,255,255,.08),transparent 32%), #063d22; color: #fff; border-radius: 16px; padding: 16px; max-height: 230px; overflow: auto; line-height: 1.65; font-size: 13px; white-space: pre-wrap; }
.cardActions, .toolActions, .panelFooter { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 14px; }
.cardActions .primaryAction { flex: 1 1 140px; }
.sectionTitleRow { padding: 28px; display: flex; justify-content: space-between; align-items: center; gap: 24px; margin-bottom: 22px; }
.sectionTitleRow.compact h2 { font-size: clamp(28px, 3vw, 40px); }
.titleActions { display: flex; gap: 12px; flex-wrap: wrap; }
.improveGrid { display: grid; grid-template-columns: 1fr 1fr; gap: 26px; align-items: start; }
.improvePanel { padding: 26px; }
.improvePanel h3 { margin: 0 0 14px; color: var(--dadi-green-dark); }
.improvePanel label { display: block; color: var(--dadi-green-dark); font-weight: 780; margin-bottom: 10px; }
.improvePanel label span { float: right; color: var(--text-muted); font-weight: 500; }
.improvePanel textarea { min-height: 112px; padding: 14px 16px; resize: vertical; }
.improvePanel input, .improvePanel select { min-height: 46px; padding: 0 14px; margin-top: 8px; }
.miniFormGrid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin-top: 18px; }
.resultHeader { display: flex; justify-content: space-between; gap: 16px; align-items: start; }
.resultHeader h3 span { background: #d8f5df; color: var(--dadi-green); border-radius: 999px; padding: 6px 12px; font-size: 13px; margin-left: 8px; }
.promptScore { width: 64px; height: 64px; border-radius: 50%; border: 7px solid var(--dadi-green); display: grid; place-items: center; color: var(--dadi-green-dark); font-weight: 950; }
.promptScore span { font-weight: 700; color: var(--text-muted); font-size: 12px; margin-left: 1px; }
.codeBox { position: relative; margin-top: 14px; }
.codeBox button { position: absolute; right: 14px; top: 14px; z-index: 2; border: 1px solid rgba(255,255,255,.35); color: #fff; background: rgba(255,255,255,.09); border-radius: 10px; padding: 8px 12px; display: flex; align-items: center; gap: 7px; font-weight: 850; }
.codeBox textarea { width: 100%; min-height: 420px; background: radial-gradient(circle at top right,rgba(255,255,255,.10),transparent 34%), #063d22; color: #fff; border-color: #0b6a3a; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; line-height: 1.65; padding: 24px; }
.tipText, .helperText { color: var(--text-muted); }
.bottomGrid { display: grid; grid-template-columns: 1.2fr .8fr; gap: 24px; margin-top: 24px; }
.breakdownPanel { padding: 22px; }
.breakdownList { display: grid; grid-template-columns: repeat(5,1fr); gap: 14px; }
.breakdownList span { border: 1px solid var(--border-soft); border-radius: 16px; padding: 16px; color: var(--dadi-green-dark); font-weight: 850; }
.breakdownList small { color: var(--text-muted); font-weight: 500; }
.suggestionTags { display: flex; flex-wrap: wrap; gap: 12px; }
.assistantWorkspace { display: grid; grid-template-columns: .9fr 1.1fr; gap: 24px; }
.inputPanel, .resultPanel { padding: 24px; }
.inputPanel textarea { width: 100%; padding: 16px; resize: vertical; }
.assistantResult { border: 1px solid var(--border-soft); border-radius: 18px; padding: 16px; margin-bottom: 12px; background: #fbfdf9; }
.assistantResult p { margin: 0 0 6px; color: var(--dadi-orange); font-size: 12px; font-weight: 950; }
.assistantResult h4 { margin: 0 0 8px; font-size: 18px; color: var(--dadi-green-dark); }
.assistantResult span { color: var(--text-muted); font-size: 13px; }
.uploadGrid { display: grid; grid-template-columns: .8fr 1.2fr; gap: 24px; }
.uploadDrop, .uploadInfo, .statusCard { padding: 24px; }
.uploadIcon { width: 78px; height: 78px; border-radius: 28px; background: var(--dadi-yellow); display: grid; place-items: center; font-size: 34px; box-shadow: 0 14px 30px rgba(255,184,0,.25); }
.uploadInfo pre { background: #063d22; color: #fff; border-radius: 16px; padding: 18px; overflow: auto; }
.guideGrid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 18px; }
.guideCard { padding: 22px; }
.scoreBubble { min-width: 150px; height: 150px; border-radius: 50%; display: grid; place-items: center; background: var(--dadi-green); color: #fff; font-size: 32px; font-weight: 950; text-align: center; }
.scoreBubble span { display: block; font-size: 12px; font-weight: 700; opacity: .9; }
.emptyState { padding: 38px; text-align: center; background: #fff; border-radius: 24px; border: 1px solid var(--border-soft); }
.modalOverlay { position: fixed; inset: 0; z-index: 80; background: rgba(7,32,17,.46); display: grid; place-items: center; padding: 24px; }
.promptModal { width: min(920px, 96vw); max-height: 90vh; overflow: auto; background: #fff; border-radius: 28px; padding: 26px; box-shadow: 0 30px 80px rgba(0,0,0,.26); }
.modalHeader { display: flex; justify-content: space-between; gap: 16px; align-items: start; }
.modalHeader h2 { margin: 0; font-size: 32px; color: var(--dadi-green-dark); }
.iconClose { border: 0; background: #f1f5f0; color: var(--text-main); width: 40px; height: 40px; border-radius: 50%; font-size: 22px; font-weight: 900; }
.modalMeta { display: flex; gap: 8px; flex-wrap: wrap; margin: 18px 0; }
.modalMeta span { background: var(--dadi-green-soft); color: var(--dadi-green-dark); padding: 8px 11px; border-radius: 999px; font-size: 13px; font-weight: 850; }
.modalUseCase, .modalPlaceholders { color: var(--text-muted); line-height: 1.6; }
.modalPromptText { max-height: 370px; }
.floatingAssistant { position: fixed; right: 28px; bottom: 28px; z-index: 40; width: 70px; height: 70px; border: 0; border-radius: 50%; display: grid; place-items: center; background: var(--dadi-yellow); color: #102015; box-shadow: 0 18px 42px rgba(255,184,0,.34); font-size: 28px; }
.toast { position: fixed; left: 50%; bottom: 28px; transform: translateX(-50%); z-index: 100; background: #062d1a; color: #fff; padding: 12px 18px; border-radius: 999px; box-shadow: 0 18px 38px rgba(0,0,0,.22); }
.srOnly { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
@media (max-width: 1200px) {
  .topbar { grid-template-columns: 1fr; }
  .navPills { justify-content: flex-start; }
  .themeSwitch { justify-content: flex-start; }
  .libraryHero, .improveGrid, .bottomGrid, .assistantWorkspace, .uploadGrid { grid-template-columns: 1fr; }
  .contentLayout { grid-template-columns: 1fr; }
  .sidebar { position: static; max-height: 380px; }
  .guideGrid, .breakdownList { grid-template-columns: repeat(2, 1fr); }
}
@media (max-width: 720px) {
  .pageShell { width: min(100% - 28px, 1680px); }
  .topbar { padding: 12px 14px; }
  .brandLogo { width: 96px; height: 44px; }
  .brandLogo img { width: 86px; }
  .primarySearchRow, .miniFormGrid, .guideGrid, .breakdownList { grid-template-columns: 1fr; }
  .navPills button, .themeSwitch button { padding: 10px 13px; }
  .sectionTitleRow { display: block; }
}
`;
