(function () {
  const { sections, metadata } = window.ASSESSMENT_DATA;
  const STORAGE_KEY = "scdc_dropping_anchor_session_v1";
  const SECTION_TO_HASH = {
    strengths: "assessing",
    support: "support"
  };
  const HASH_TO_SECTION = Object.fromEntries(
    Object.entries(SECTION_TO_HASH).map(([sectionId, hash]) => [hash, sectionId])
  );
  const state = {
    activeSectionId: sections[0].id,
    summaryCollapsed: true,
    session: {
      metadata: { orgName: "", completedBy: "", completedDate: "" },
      answers: {},
      reflection: {}
    }
  };

  const dom = {
    subtitle: document.getElementById("subtitle"),
    sectionTabs: document.getElementById("sectionTabs"),
    sectionTitle: document.getElementById("sectionTitle"),
    sectionInstructions: document.getElementById("sectionInstructions"),
    scaleLegend: document.getElementById("scaleLegend"),
    questionsContainer: document.getElementById("questionsContainer"),
    scoredCount: document.getElementById("scoredCount"),
    averageScore: document.getElementById("averageScore"),
    percentScore: document.getElementById("percentScore"),
    resultNarrative: document.getElementById("resultNarrative"),
    calculateBtn: document.getElementById("calculateBtn"),
    importBtn: document.getElementById("importBtn"),
    importFileInput: document.getElementById("importFileInput"),
    importStatus: document.getElementById("importStatus"),
    exportWordBtn: document.getElementById("exportWordBtn"),
    exportPdfBtn: document.getElementById("exportPdfBtn"),
    newSessionBtn: document.getElementById("newSessionBtn"),
    summaryPanel: document.getElementById("summaryPanel"),
    summaryToggle: document.getElementById("summaryToggle"),
    summaryContent: document.getElementById("summaryContent"),
    missingHint: document.getElementById("missingHint"),
    progressFill: document.getElementById("progressFill"),
    progressLabel: document.getElementById("progressLabel"),
    orgName: document.getElementById("orgName"),
    completedBy: document.getElementById("completedBy"),
    completedDate: document.getElementById("completedDate")
  };

  if (window.pdfjsLib) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  function loadSession() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") return;
      state.activeSectionId = parsed.activeSectionId || state.activeSectionId;
      state.summaryCollapsed = parsed.summaryCollapsed ?? state.summaryCollapsed;
      state.session.metadata = parsed.metadata || state.session.metadata;
      state.session.answers = parsed.answers || {};
      state.session.reflection = parsed.reflection || {};
    } catch (error) {
      console.warn("Failed to load previous session:", error);
    }
  }

  function saveSession() {
    const payload = {
      activeSectionId: state.activeSectionId,
      summaryCollapsed: state.summaryCollapsed,
      metadata: state.session.metadata,
      answers: state.session.answers,
      reflection: state.session.reflection
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  }

  function ensureSectionStores(sectionId) {
    if (!state.session.answers[sectionId]) state.session.answers[sectionId] = {};
    if (!state.session.reflection[sectionId]) state.session.reflection[sectionId] = {};
  }

  function getStoredAnswer(sectionId, categoryIndex, questionIndex) {
    ensureSectionStores(sectionId);
    return state.session.answers[sectionId][`${categoryIndex}-${questionIndex}`] || {};
  }

  function setStoredAnswer(sectionId, categoryIndex, questionIndex, patch) {
    ensureSectionStores(sectionId);
    const key = `${categoryIndex}-${questionIndex}`;
    const existing = state.session.answers[sectionId][key] || {};
    state.session.answers[sectionId][key] = { ...existing, ...patch };
  }

  function getStoredReflection(sectionId, idx) {
    ensureSectionStores(sectionId);
    return state.session.reflection[sectionId][String(idx)] || "";
  }

  function setStoredReflection(sectionId, idx, value) {
    ensureSectionStores(sectionId);
    state.session.reflection[sectionId][String(idx)] = value;
  }

  function syncMetadataFromInputs() {
    state.session.metadata.orgName = dom.orgName.value.trim();
    state.session.metadata.completedBy = dom.completedBy.value.trim();
    state.session.metadata.completedDate = dom.completedDate.value;
  }

  function applyMetadataToInputs() {
    dom.orgName.value = state.session.metadata.orgName || "";
    dom.completedBy.value = state.session.metadata.completedBy || "";
    if (state.session.metadata.completedDate) {
      dom.completedDate.value = state.session.metadata.completedDate;
    } else {
      dom.completedDate.valueAsDate = new Date();
      state.session.metadata.completedDate = dom.completedDate.value;
    }
  }

  function normalizeText(value) {
    return (value || "")
      .toLowerCase()
      .replace(/&/g, " and ")
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201c\u201d]/g, '"')
      .replace(/[^a-z0-9]+/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function buildQuestionLookup() {
    const map = new Map();
    sections.forEach((section) => {
      section.categories.forEach((category, cIdx) => {
        category.questions.forEach((question, qIdx) => {
          map.set(normalizeText(question), {
            sectionId: section.id,
            categoryIndex: cIdx,
            questionIndex: qIdx
          });
        });
      });
    });
    return map;
  }

  const questionLookup = buildQuestionLookup();

  function sectionIdFromHash(hashValue) {
    const clean = (hashValue || "").replace(/^#/, "").trim().toLowerCase();
    return HASH_TO_SECTION[clean] || null;
  }

  function hashForSectionId(sectionId) {
    return SECTION_TO_HASH[sectionId] || sectionId;
  }

  function syncUrlHashToActiveSection() {
    const target = `#${hashForSectionId(state.activeSectionId)}`;
    if (window.location.hash !== target) {
      window.history.replaceState(null, "", target);
    }
  }

  function getActiveSection() {
    return sections.find((section) => section.id === state.activeSectionId);
  }

  function renderTabs() {
    dom.sectionTabs.innerHTML = "";
    sections.forEach((section) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = `tab-btn${section.id === state.activeSectionId ? " active" : ""}`;
      btn.setAttribute("role", "tab");
      btn.setAttribute("aria-selected", section.id === state.activeSectionId ? "true" : "false");
      btn.textContent = section.label;
      btn.addEventListener("click", () => {
        state.activeSectionId = section.id;
        saveSession();
        syncUrlHashToActiveSection();
        render();
      });
      dom.sectionTabs.appendChild(btn);
    });
  }

  function renderScale(section) {
    dom.scaleLegend.innerHTML = "";
    section.scale.forEach((item) => {
      const chip = document.createElement("div");
      chip.className = "scale-chip";
      chip.textContent = `${item.value} = ${item.label}`;
      dom.scaleLegend.appendChild(chip);
    });
  }

  function createQuestionRow(sectionId, categoryIndex, questionIndex, questionText, scale) {
    const wrapper = document.createElement("article");
    wrapper.className = "question";

    const qId = `${sectionId}-${categoryIndex}-${questionIndex}`;

    const title = document.createElement("p");
    title.className = "q-title";
    title.textContent = questionText;
    wrapper.appendChild(title);

    const options = document.createElement("div");
    options.className = "score-options";

    scale.forEach((item) => {
      const label = document.createElement("label");
      const input = document.createElement("input");
      input.type = "radio";
      input.name = `${qId}-score`;
      input.value = String(item.value);
      input.dataset.sectionId = sectionId;
      input.dataset.categoryIndex = String(categoryIndex);
      input.dataset.questionIndex = String(questionIndex);
      label.appendChild(input);
      if (item.value === 1) {
        label.append(document.createTextNode("1 - Not at all"));
      } else if (item.value === 5) {
        label.append(document.createTextNode("5 - Fully"));
      } else {
        label.append(document.createTextNode(String(item.value)));
      }
      options.appendChild(label);
    });

    const savedAnswer = getStoredAnswer(sectionId, categoryIndex, questionIndex);
    if (savedAnswer.score) {
      const radioToCheck = options.querySelector(`input[value="${savedAnswer.score}"]`);
      if (radioToCheck) radioToCheck.checked = true;
    }

    wrapper.appendChild(options);

    const commentWrap = document.createElement("div");
    commentWrap.className = "comment-wrap";

    const commentLabel = document.createElement("label");
    commentLabel.setAttribute("for", `${qId}-comment`);
    commentLabel.textContent = "Comment";

    const textarea = document.createElement("textarea");
    textarea.id = `${qId}-comment`;
    textarea.dataset.sectionId = sectionId;
    textarea.dataset.categoryIndex = String(categoryIndex);
    textarea.dataset.questionIndex = String(questionIndex);
    textarea.placeholder = "Add notes for improvement or support needed";
    textarea.value = savedAnswer.comment || "";

    commentWrap.appendChild(commentLabel);
    commentWrap.appendChild(textarea);
    wrapper.appendChild(commentWrap);

    return wrapper;
  }

  function renderQuestions(section) {
    dom.questionsContainer.innerHTML = "";

    section.categories.forEach((category, cIdx) => {
      const categoryEl = document.createElement("section");
      categoryEl.className = "category";

      const title = document.createElement("h3");
      title.textContent = category.title;
      categoryEl.appendChild(title);

      category.questions.forEach((question, qIdx) => {
        categoryEl.appendChild(createQuestionRow(section.id, cIdx, qIdx, question, section.scale));
      });

      dom.questionsContainer.appendChild(categoryEl);
    });

    const reflectionWrap = document.createElement("section");
    reflectionWrap.className = "category reflection-wrap";

    const reflectionTitle = document.createElement("h4");
    reflectionTitle.textContent = "Final Reflection";
    reflectionWrap.appendChild(reflectionTitle);

    section.reflection.forEach((prompt, idx) => {
      const field = document.createElement("div");
      field.className = "comment-wrap";

      const label = document.createElement("label");
      label.setAttribute("for", `${section.id}-reflection-${idx}`);
      label.textContent = prompt;

      const textarea = document.createElement("textarea");
      textarea.id = `${section.id}-reflection-${idx}`;
      textarea.dataset.reflectionPrompt = prompt;
      textarea.value = getStoredReflection(section.id, idx);

      field.appendChild(label);
      field.appendChild(textarea);
      reflectionWrap.appendChild(field);
    });

    dom.questionsContainer.appendChild(reflectionWrap);
  }

  function getScoring(sectionId) {
    const section = sections.find((item) => item.id === sectionId);
    let total = 0;
    let count = 0;

    section.categories.forEach((category, cIdx) => {
      category.questions.forEach((_, qIdx) => {
        const savedAnswer = getStoredAnswer(sectionId, cIdx, qIdx);
        if (savedAnswer.score) {
          count += 1;
          total += Number(savedAnswer.score);
        }
      });
    });

    const max = section.categories.reduce((sum, category) => sum + category.questions.length, 0) * 5;
    const avg = count ? total / count : 0;
    const percent = max ? (total / max) * 100 : 0;

    return { total, count, avg, percent, max };
  }

  function getMissingQuestionNumbers(sectionId) {
    const section = sections.find((item) => item.id === sectionId);
    let runningQuestionNumber = 0;
    const missing = [];

    section.categories.forEach((category, cIdx) => {
      category.questions.forEach((_, qIdx) => {
        runningQuestionNumber += 1;
        const savedAnswer = getStoredAnswer(sectionId, cIdx, qIdx);
        if (!savedAnswer.score) {
          missing.push(runningQuestionNumber);
        }
      });
    });

    return missing;
  }

  function getQuestionTotal(section) {
    return section.categories.reduce((sum, c) => sum + c.questions.length, 0);
  }

  function setSummaryCollapsed(collapsed) {
    state.summaryCollapsed = collapsed;
    dom.summaryPanel.classList.toggle("collapsed", collapsed);
    dom.summaryToggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
    saveSession();
  }

  function getScoreMessage(avg) {
    if (avg >= 4.5) return "Very strong overall position.";
    if (avg >= 3.5) return "Solid progress with targeted improvements to make.";
    if (avg >= 2.5) return "Mixed picture; focus on a few priority areas next.";
    if (avg > 0) return "Early-stage performance; stronger support and planning likely needed.";
    return "No scores entered yet.";
  }

  function getScoredQuestionItems(sectionId) {
    const section = sections.find((item) => item.id === sectionId);
    const scored = [];
    let questionNumber = 0;

    section.categories.forEach((category, cIdx) => {
      category.questions.forEach((question, qIdx) => {
        questionNumber += 1;
        const savedAnswer = getStoredAnswer(sectionId, cIdx, qIdx);
        if (savedAnswer.score) {
          scored.push({
            questionNumber,
            question,
            score: Number(savedAnswer.score),
            category: category.title
          });
        }
      });
    });

    return scored;
  }

  function buildSectionAnalysis(sectionId) {
    const section = sections.find((item) => item.id === sectionId);
    const score = getScoring(sectionId);
    const totalQuestions = getQuestionTotal(section);
    const completionPercent = totalQuestions ? Math.round((score.count / totalQuestions) * 100) : 0;
    const missingNumbers = getMissingQuestionNumbers(sectionId);
    const scoredItems = getScoredQuestionItems(sectionId);

    const strongest = [...scoredItems]
      .sort((a, b) => b.score - a.score || a.questionNumber - b.questionNumber)
      .slice(0, 4);
    const focusAreas = [...scoredItems]
      .sort((a, b) => a.score - b.score || a.questionNumber - b.questionNumber)
      .slice(0, 6);

    const categoryAnalysis = section.categories
      .map((category, cIdx) => {
        let answered = 0;
        let total = 0;
        category.questions.forEach((_, qIdx) => {
          const saved = getStoredAnswer(sectionId, cIdx, qIdx);
          if (saved.score) {
            answered += 1;
            total += Number(saved.score);
          }
        });
        return {
          title: category.title,
          answered,
          totalQuestions: category.questions.length,
          average: answered ? total / answered : 0
        };
      })
      .filter((entry) => entry.answered > 0);

    const weakestCategories = [...categoryAnalysis]
      .sort((a, b) => a.average - b.average)
      .slice(0, 2);

    return {
      sectionLabel: section.label,
      totalQuestions,
      completionPercent,
      score,
      headline: getScoreMessage(score.avg),
      missingNumbers,
      strongest,
      focusAreas,
      weakestCategories
    };
  }

  function setImportStatus(message, type) {
    dom.importStatus.textContent = message || "";
    if (type === "error") {
      dom.importStatus.style.color = "#8a2d2d";
      return;
    }
    if (type === "success") {
      dom.importStatus.style.color = "#2f5f3a";
      return;
    }
    dom.importStatus.style.color = "#405462";
  }

  function stripHtmlToText(raw) {
    const htmlWithBreaks = raw
      .replace(/<\s*br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|h1|h2|h3|h4|h5|h6|li|tr|section|article)>/gi, "$&\n");
    const doc = new DOMParser().parseFromString(htmlWithBreaks, "text/html");
    return (doc.body ? doc.body.textContent : htmlWithBreaks) || "";
  }

  function textToLines(text) {
    return text
      .replace(/\r\n/g, "\n")
      .replace(/\r/g, "\n")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);
  }

  async function readPdfText(file) {
    if (!window.pdfjsLib) {
      throw new Error("PDF import is unavailable because the PDF library could not be loaded.");
    }
    const data = new Uint8Array(await file.arrayBuffer());
    const pdf = await window.pdfjsLib.getDocument({ data }).promise;
    const pageText = [];
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const line = content.items.map((item) => item.str).join(" ");
      pageText.push(line);
    }
    return pageText.join("\n");
  }

  async function readDocxText(file) {
    if (!window.mammoth) {
      throw new Error("Word import is unavailable because the DOCX library could not be loaded.");
    }
    const arrayBuffer = await file.arrayBuffer();
    const result = await window.mammoth.extractRawText({ arrayBuffer });
    return result.value || "";
  }

  function applyParsedImport(lines) {
    const imported = {
      metadata: {},
      answers: {},
      reflection: {}
    };

    function ensureImportedSection(sectionId) {
      if (!imported.answers[sectionId]) imported.answers[sectionId] = {};
      if (!imported.reflection[sectionId]) imported.reflection[sectionId] = {};
    }

    lines.forEach((line, idx) => {
      if (line.startsWith("Organisation / Group:")) {
        imported.metadata.orgName = line.replace("Organisation / Group:", "").trim();
      } else if (line.startsWith("Completed by:")) {
        imported.metadata.completedBy = line.replace("Completed by:", "").trim();
      } else if (line.startsWith("Date:")) {
        imported.metadata.completedDate = line.replace("Date:", "").trim();
      }

      if (line.startsWith("Score:") || line.startsWith("Score :") || line.startsWith("Score")) {
        const scoreMatch = line.match(/([1-5])\s*\/\s*5/);
        const previous = lines[idx - 1] || "";
        const questionCandidate = previous.replace(/^-+\s*/, "").trim();
        const lookup = questionLookup.get(normalizeText(questionCandidate));
        if (lookup && scoreMatch) {
          ensureImportedSection(lookup.sectionId);
          const key = `${lookup.categoryIndex}-${lookup.questionIndex}`;
          const existing = imported.answers[lookup.sectionId][key] || {};
          imported.answers[lookup.sectionId][key] = {
            ...existing,
            score: Number(scoreMatch[1])
          };
        }
      }

      if (line.startsWith("Comment:")) {
        const previousScoreLine = lines[idx - 1] || "";
        const questionLine = lines[idx - 2] || "";
        const questionCandidate = questionLine.replace(/^-+\s*/, "").trim();
        const lookup = questionLookup.get(normalizeText(questionCandidate));
        const commentValue = line.replace("Comment:", "").trim();
        if (lookup) {
          ensureImportedSection(lookup.sectionId);
          const key = `${lookup.categoryIndex}-${lookup.questionIndex}`;
          const existing = imported.answers[lookup.sectionId][key] || {};
          imported.answers[lookup.sectionId][key] = {
            ...existing,
            comment: commentValue === "-" ? "" : commentValue
          };
        } else if (previousScoreLine) {
          // no-op, retained for readability in mixed imports
        }
      }

      if (line.startsWith("- ")) {
        sections.forEach((section) => {
          section.reflection.forEach((prompt, promptIndex) => {
            if (normalizeText(line.slice(2)) === normalizeText(prompt)) {
              const nextLine = (lines[idx + 1] || "").trim();
              ensureImportedSection(section.id);
              imported.reflection[section.id][String(promptIndex)] = nextLine === "-" ? "" : nextLine;
            }
          });
        });
      }
    });

    let importedScoreCount = 0;
    const importedSectionIdsWithScores = new Set();
    Object.values(imported.answers).forEach((sectionAnswers) => {
      Object.values(sectionAnswers).forEach((entry) => {
        if (entry && entry.score) importedScoreCount += 1;
      });
    });
    Object.entries(imported.answers).forEach(([sectionId, sectionAnswers]) => {
      const hasScores = Object.values(sectionAnswers).some((entry) => entry && entry.score);
      if (hasScores) importedSectionIdsWithScores.add(sectionId);
    });

    if (!importedScoreCount) {
      return { importedCount: 0, importedSectionIds: [] };
    }

    Object.entries(imported.answers).forEach(([sectionId, sectionAnswers]) => {
      if (!state.session.answers[sectionId]) state.session.answers[sectionId] = {};
      state.session.answers[sectionId] = {
        ...state.session.answers[sectionId],
        ...sectionAnswers
      };
    });

    Object.entries(imported.reflection).forEach(([sectionId, reflectionValues]) => {
      if (!state.session.reflection[sectionId]) state.session.reflection[sectionId] = {};
      state.session.reflection[sectionId] = {
        ...state.session.reflection[sectionId],
        ...reflectionValues
      };
    });

    state.session.metadata = {
      ...state.session.metadata,
      ...Object.fromEntries(
        Object.entries(imported.metadata).filter(([, value]) => typeof value === "string" && value.length > 0)
      )
    };
    saveSession();
    applyMetadataToInputs();
    return {
      importedCount: importedScoreCount,
      importedSectionIds: Array.from(importedSectionIdsWithScores)
    };
  }

  async function importFromSelectedFile(file) {
    const fileName = file.name || "selected file";
    const ext = fileName.includes(".") ? fileName.split(".").pop().toLowerCase() : "";
    let rawText = "";

    if (ext === "pdf") {
      rawText = await readPdfText(file);
    } else if (ext === "docx") {
      rawText = await readDocxText(file);
    } else {
      rawText = await file.text();
      const looksLikeHtml = /<\s*html|<\s*body|<\s*p|<\s*div/i.test(rawText);
      if (looksLikeHtml || ext === "doc" || ext === "htm" || ext === "html") {
        rawText = stripHtmlToText(rawText);
      }
    }

    const lines = textToLines(rawText);
    return applyParsedImport(lines);
  }

  function calculateAndRenderSummary() {
    const active = getActiveSection();
    const analysis = buildSectionAnalysis(active.id);

    dom.scoredCount.textContent = `${analysis.score.count} / ${analysis.totalQuestions}`;
    dom.averageScore.textContent = analysis.score.avg.toFixed(1);
    dom.percentScore.textContent = `${Math.round(analysis.score.percent)}%`;
    dom.progressFill.style.width = `${analysis.completionPercent}%`;
    dom.progressLabel.textContent = `${analysis.completionPercent}%`;

    if (analysis.missingNumbers.length > 0 && analysis.missingNumbers.length <= 3) {
      dom.missingHint.textContent = `Left to complete ${analysis.missingNumbers.map((n) => `Q${n}`).join(", ")}`;
      dom.missingHint.classList.add("visible");
    } else {
      dom.missingHint.textContent = "";
      dom.missingHint.classList.remove("visible");
    }

    const container = dom.resultNarrative;
    container.innerHTML = "";

    const p1 = document.createElement("p");
    p1.textContent = analysis.headline;
    container.appendChild(p1);

    const p2 = document.createElement("p");
    p2.textContent = `Completion is ${analysis.completionPercent}% with ${analysis.score.count} of ${analysis.totalQuestions} questions scored.`;
    container.appendChild(p2);

    if (analysis.focusAreas.length > 0) {
      const p3 = document.createElement("p");
      p3.textContent = "Focus areas (lowest scores):";
      container.appendChild(p3);

      const focusList = document.createElement("ul");
      analysis.focusAreas.slice(0, 5).forEach((item) => {
        const li = document.createElement("li");
        li.textContent = `Q${item.questionNumber}: ${item.score}/5 - ${item.question}`;
        focusList.appendChild(li);
      });
      container.appendChild(focusList);
    }

    if (analysis.strongest.length > 0) {
      const p4 = document.createElement("p");
      p4.textContent = "Current strengths:";
      container.appendChild(p4);

      const strengths = document.createElement("ul");
      analysis.strongest.slice(0, 3).forEach((item) => {
        const li = document.createElement("li");
        li.textContent = `Q${item.questionNumber}: ${item.score}/5 - ${item.question}`;
        strengths.appendChild(li);
      });
      container.appendChild(strengths);
    }

    if (analysis.weakestCategories.length > 0) {
      const p5 = document.createElement("p");
      p5.textContent = "Category focus:";
      container.appendChild(p5);

      const categories = document.createElement("ul");
      analysis.weakestCategories.forEach((cat) => {
        const li = document.createElement("li");
        li.textContent = `${cat.title} (avg ${cat.average.toFixed(1)}/5 across ${cat.answered}/${cat.totalQuestions} answered)`;
        categories.appendChild(li);
      });
      container.appendChild(categories);
    }

    if (analysis.missingNumbers.length > 0) {
      const p6 = document.createElement("p");
      p6.textContent = `Unanswered questions: ${analysis.missingNumbers.map((n) => `Q${n}`).join(", ")}.`;
      container.appendChild(p6);
    }

    if (analysis.totalQuestions > 0 && analysis.score.count === analysis.totalQuestions) {
      setSummaryCollapsed(false);
    }
  }

  function gatherSectionResponses(section) {
    const categories = section.categories.map((category, cIdx) => {
      const questions = category.questions.map((question, qIdx) => {
        const savedAnswer = getStoredAnswer(section.id, cIdx, qIdx);

        return {
          question,
          score: savedAnswer.score ? Number(savedAnswer.score) : null,
          comment: (savedAnswer.comment || "").trim()
        };
      });

      return {
        title: category.title,
        questions
      };
    });

    const reflection = section.reflection.map((prompt, idx) => {
      const value = getStoredReflection(section.id, idx).trim();
      return { prompt, value };
    });

    const score = getScoring(section.id);

    return {
      sectionLabel: section.label,
      categories,
      reflection,
      score
    };
  }

  function gatherResponsesBySectionIds(sectionIds) {
    return sectionIds.map((sectionId) => {
      const section = sections.find((item) => item.id === sectionId);
      return gatherSectionResponses(section);
    });
  }

  function gatherAllResponses() {
    return gatherResponsesBySectionIds(sections.map((section) => section.id));
  }

  function appendSectionAnalysisHtml(section, htmlParts) {
    const sectionId = sections.find((entry) => entry.label === section.sectionLabel)?.id;
    if (!sectionId) return;
    const analysis = buildSectionAnalysis(sectionId);

    htmlParts.push("<h3>Analysis & Focus Areas</h3>");
    htmlParts.push(
      `<p><strong>${escapeHtml(analysis.headline)}</strong> Completion ${analysis.completionPercent}% (${analysis.score.count}/${analysis.totalQuestions}).</p>`
    );
    htmlParts.push("<ul>");
    analysis.focusAreas.slice(0, 6).forEach((item) => {
      htmlParts.push(`<li>Focus Q${item.questionNumber}: ${item.score}/5 - ${escapeHtml(item.question)}</li>`);
    });
    analysis.strongest.slice(0, 4).forEach((item) => {
      htmlParts.push(`<li>Strength Q${item.questionNumber}: ${item.score}/5 - ${escapeHtml(item.question)}</li>`);
    });
    analysis.weakestCategories.forEach((cat) => {
      htmlParts.push(`<li>Category focus: ${escapeHtml(cat.title)} (avg ${cat.average.toFixed(1)}/5)</li>`);
    });
    if (analysis.missingNumbers.length > 0) {
      htmlParts.push(`<li>Unanswered: ${escapeHtml(analysis.missingNumbers.map((n) => `Q${n}`).join(", "))}</li>`);
    }
    htmlParts.push("</ul>");
  }

  function appendSectionAnalysisText(section, lines) {
    const sectionId = sections.find((entry) => entry.label === section.sectionLabel)?.id;
    if (!sectionId) return;
    const analysis = buildSectionAnalysis(sectionId);

    lines.push("Analysis & Focus Areas");
    lines.push(analysis.headline);
    lines.push(`Completion: ${analysis.completionPercent}% (${analysis.score.count}/${analysis.totalQuestions})`);
    analysis.focusAreas.slice(0, 6).forEach((item) => {
      lines.push(`- Focus Q${item.questionNumber}: ${item.score}/5 - ${item.question}`);
    });
    analysis.strongest.slice(0, 4).forEach((item) => {
      lines.push(`- Strength Q${item.questionNumber}: ${item.score}/5 - ${item.question}`);
    });
    analysis.weakestCategories.forEach((cat) => {
      lines.push(`- Category focus: ${cat.title} (avg ${cat.average.toFixed(1)}/5)`);
    });
    if (analysis.missingNumbers.length > 0) {
      lines.push(`- Unanswered: ${analysis.missingNumbers.map((n) => `Q${n}`).join(", ")}`);
    }
    lines.push("");
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function buildReportHtml(sectionIds) {
    const orgName = dom.orgName.value.trim() || "Not provided";
    const completedBy = dom.completedBy.value.trim() || "Not provided";
    const completedDate = dom.completedDate.value || new Date().toISOString().slice(0, 10);

    const responses = gatherResponsesBySectionIds(sectionIds);

    let html = `
      <html>
      <head>
        <meta charset="utf-8" />
        <title>Dropping Anchor Self-Assessment Report</title>
        <style>
          body { font-family: Arial, sans-serif; line-height: 1.4; color: #222; }
          h1, h2, h3 { color: #2f4654; }
          .meta { background: #f4f6f8; padding: 12px; border: 1px solid #d3d9df; margin-bottom: 14px; }
          .q { margin-bottom: 12px; }
          .score { font-weight: bold; }
          .comment { color: #333; }
          .footer-note { margin-top: 20px; padding-top: 10px; border-top: 1px solid #d3d9df; font-size: 12px; color: #435864; }
          hr { border: 0; border-top: 1px solid #d3d9df; margin: 18px 0; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(metadata.title)}</h1>
        <p>${escapeHtml(metadata.subtitle)}</p>
        <div class="meta">
          <p><strong>Organisation / Group:</strong> ${escapeHtml(orgName)}</p>
          <p><strong>Completed by:</strong> ${escapeHtml(completedBy)}</p>
          <p><strong>Date:</strong> ${escapeHtml(completedDate)}</p>
        </div>
    `;

    responses.forEach((section) => {
      const totalQuestions = section.categories.reduce((sum, c) => sum + c.questions.length, 0);
      html += `<h2>${escapeHtml(section.sectionLabel)}</h2>`;
      html += `<p><strong>Scored:</strong> ${section.score.count}/${totalQuestions} | <strong>Average:</strong> ${section.score.avg.toFixed(2)} | <strong>Section Score:</strong> ${Math.round(section.score.percent)}%</p>`;
      const sectionAnalysisParts = [];
      appendSectionAnalysisHtml(section, sectionAnalysisParts);
      html += sectionAnalysisParts.join("");

      section.categories.forEach((category) => {
        html += `<h3>${escapeHtml(category.title)}</h3>`;
        category.questions.forEach((q) => {
          html += `<div class="q"><p>${escapeHtml(q.question)}</p><p class="score">Score: ${q.score === null ? "Not scored" : q.score + "/5"}</p><p class="comment">Comment: ${escapeHtml(q.comment || "-")}</p></div>`;
        });
      });

      html += `<h3>Reflection</h3>`;
      section.reflection.forEach((item) => {
        html += `<p><strong>${escapeHtml(item.prompt)}</strong><br/>${escapeHtml(item.value || "-")}</p>`;
      });

      html += "<hr/>";
    });

    html += `<p class="footer-note">This scoring was carried out at <a href="https://anchors.scdc.org.uk/">anchors.scdc.org.uk</a>, a tool developed by the Scottish Community Development Centre.</p>`;
    html += "</body></html>";
    return html;
  }

  function buildReportPlainText(sectionIds) {
    const orgName = dom.orgName.value.trim() || "Not provided";
    const completedBy = dom.completedBy.value.trim() || "Not provided";
    const completedDate = dom.completedDate.value || new Date().toISOString().slice(0, 10);

    const lines = [];
    lines.push(metadata.title);
    lines.push(metadata.subtitle);
    lines.push("");
    lines.push(`Organisation / Group: ${orgName}`);
    lines.push(`Completed by: ${completedBy}`);
    lines.push(`Date: ${completedDate}`);
    lines.push("");

    gatherResponsesBySectionIds(sectionIds).forEach((section) => {
      const totalQuestions = section.categories.reduce((sum, c) => sum + c.questions.length, 0);
      lines.push(section.sectionLabel);
      lines.push(`Scored: ${section.score.count}/${totalQuestions} | Average: ${section.score.avg.toFixed(2)} | Section Score: ${Math.round(section.score.percent)}%`);
      lines.push("");
      appendSectionAnalysisText(section, lines);

      section.categories.forEach((category) => {
        lines.push(category.title);
        category.questions.forEach((q) => {
          lines.push(`- ${q.question}`);
          lines.push(`  Score: ${q.score === null ? "Not scored" : q.score + "/5"}`);
          lines.push(`  Comment: ${q.comment || "-"}`);
        });
        lines.push("");
      });

      lines.push("Reflection");
      section.reflection.forEach((item) => {
        lines.push(`- ${item.prompt}`);
        lines.push(`  ${item.value || "-"}`);
      });
      lines.push("");
      lines.push("------------------------------------------------------------");
      lines.push("");
    });

    lines.push("This scoring was carried out at anchors.scdc.org.uk, a tool developed by the Scottish Community Development Centre.");
    lines.push("https://anchors.scdc.org.uk/");
    return lines.join("\n");
  }

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function exportWord() {
    const active = getActiveSection();
    const html = buildReportHtml([active.id]);
    const blob = new Blob(["\ufeff", html], {
      type: "application/msword"
    });
    const sectionSlug = hashForSectionId(active.id);
    downloadBlob(blob, `dropping-anchor-${sectionSlug}-${new Date().toISOString().slice(0, 10)}.doc`);
  }

  function exportPdf() {
    const active = getActiveSection();
    const text = buildReportPlainText([active.id]);
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({ unit: "pt", format: "a4" });

    const margin = 40;
    const maxWidth = 515;
    const lines = pdf.splitTextToSize(text, maxWidth);

    let y = margin;
    const pageHeight = pdf.internal.pageSize.getHeight();

    lines.forEach((line) => {
      if (y > pageHeight - margin) {
        pdf.addPage();
        y = margin;
      }
      pdf.text(line, margin, y);
      y += 14;
    });

    const sectionSlug = hashForSectionId(active.id);
    pdf.save(`dropping-anchor-${sectionSlug}-${new Date().toISOString().slice(0, 10)}.pdf`);
  }

  function render() {
    const active = getActiveSection();
    dom.subtitle.textContent = metadata.subtitle;
    dom.sectionTitle.textContent = active.label;
    dom.sectionInstructions.textContent = active.instructions;

    renderTabs();
    renderScale(active);
    renderQuestions(active);
    calculateAndRenderSummary();
    setSummaryCollapsed(state.summaryCollapsed);
  }

  dom.calculateBtn.addEventListener("click", calculateAndRenderSummary);
  dom.importBtn.addEventListener("click", () => {
    dom.importFileInput.click();
  });
  dom.exportWordBtn.addEventListener("click", exportWord);
  dom.exportPdfBtn.addEventListener("click", exportPdf);
  dom.summaryToggle.addEventListener("click", () => {
    setSummaryCollapsed(!state.summaryCollapsed);
  });

  dom.questionsContainer.addEventListener("change", (event) => {
    if (event.target instanceof HTMLInputElement && event.target.type === "radio") {
      const sectionId = event.target.dataset.sectionId;
      const categoryIndex = Number(event.target.dataset.categoryIndex);
      const questionIndex = Number(event.target.dataset.questionIndex);
      setStoredAnswer(sectionId, categoryIndex, questionIndex, { score: Number(event.target.value) });
      saveSession();
      calculateAndRenderSummary();
    }
  });

  dom.questionsContainer.addEventListener("input", (event) => {
    if (event.target instanceof HTMLTextAreaElement) {
      if (event.target.id.includes("-reflection-")) {
        const parts = event.target.id.split("-reflection-");
        const sectionId = parts[0];
        const idx = Number(parts[1]);
        setStoredReflection(sectionId, idx, event.target.value);
      } else {
        const sectionId = event.target.dataset.sectionId;
        const categoryIndex = Number(event.target.dataset.categoryIndex);
        const questionIndex = Number(event.target.dataset.questionIndex);
        setStoredAnswer(sectionId, categoryIndex, questionIndex, { comment: event.target.value });
      }
      saveSession();
    }
  });

  [dom.orgName, dom.completedBy, dom.completedDate].forEach((input) => {
    input.addEventListener("input", () => {
      syncMetadataFromInputs();
      saveSession();
    });
    input.addEventListener("change", () => {
      syncMetadataFromInputs();
      saveSession();
    });
  });

  dom.newSessionBtn.addEventListener("click", () => {
    const ok = window.confirm("This will remove any entered information. Are you sure?");
    if (!ok) return;
    localStorage.removeItem(STORAGE_KEY);
    state.activeSectionId = sections[0].id;
    state.summaryCollapsed = true;
    state.session = {
      metadata: { orgName: "", completedBy: "", completedDate: "" },
      answers: {},
      reflection: {}
    };
    syncUrlHashToActiveSection();
    applyMetadataToInputs();
    render();
    setImportStatus("");
  });

  dom.importFileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    if (!file) return;
    setImportStatus(`Importing ${file.name}...`);
    try {
      const result = await importFromSelectedFile(file);
      if (!result.importedCount) {
        setImportStatus(
          "No answers were detected. Import works with reports exported from this tool (.doc, .docx, .pdf).",
          "error"
        );
      } else {
        if (result.importedSectionIds.length === 1) {
          const [importedSectionId] = result.importedSectionIds;
          if (importedSectionId && importedSectionId !== state.activeSectionId) {
            state.activeSectionId = importedSectionId;
          }
        }
        saveSession();
        render();
        syncUrlHashToActiveSection();
        setImportStatus(`Imported ${result.importedCount} scored answers from ${file.name}.`, "success");
      }
    } catch (error) {
      setImportStatus(`Import failed: ${error.message || "Unable to read file."}`, "error");
    } finally {
      dom.importFileInput.value = "";
    }
  });

  loadSession();
  const hashedSection = sectionIdFromHash(window.location.hash);
  if (hashedSection) {
    state.activeSectionId = hashedSection;
  }
  applyMetadataToInputs();
  render();
  syncUrlHashToActiveSection();

  window.addEventListener("hashchange", () => {
    const sectionId = sectionIdFromHash(window.location.hash);
    if (!sectionId || sectionId === state.activeSectionId) return;
    state.activeSectionId = sectionId;
    saveSession();
    render();
  });
})();
