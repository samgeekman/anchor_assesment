(function () {
  const { sections, metadata } = window.ASSESSMENT_DATA;
  const STORAGE_KEY = "scdc_dropping_anchor_session_v1";
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
    exportWordBtn: document.getElementById("exportWordBtn"),
    exportPdfBtn: document.getElementById("exportPdfBtn"),
    newSessionBtn: document.getElementById("newSessionBtn"),
    summaryPanel: document.getElementById("summaryPanel"),
    summaryToggle: document.getElementById("summaryToggle"),
    summaryContent: document.getElementById("summaryContent"),
    progressFill: document.getElementById("progressFill"),
    progressLabel: document.getElementById("progressLabel"),
    orgName: document.getElementById("orgName"),
    completedBy: document.getElementById("completedBy"),
    completedDate: document.getElementById("completedDate")
  };

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
        const name = `${sectionId}-${cIdx}-${qIdx}-score`;
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

  function getLowestScoredItems(sectionId, limit) {
    const section = sections.find((item) => item.id === sectionId);
    const scored = [];

    section.categories.forEach((category, cIdx) => {
      category.questions.forEach((question, qIdx) => {
        const name = `${sectionId}-${cIdx}-${qIdx}-score`;
        const selected = document.querySelector(`input[name="${name}"]:checked`);
        if (selected) {
          scored.push({
            question,
            score: Number(selected.value),
            category: category.title
          });
        }
      });
    });

    return scored.sort((a, b) => a.score - b.score).slice(0, limit);
  }

  function calculateAndRenderSummary() {
    const active = getActiveSection();
    const score = getScoring(active.id);

    const questionTotal = getQuestionTotal(active);
    const completionPercent = questionTotal ? Math.round((score.count / questionTotal) * 100) : 0;
    dom.scoredCount.textContent = `${score.count} / ${questionTotal}`;
    dom.averageScore.textContent = score.avg.toFixed(1);
    dom.percentScore.textContent = `${Math.round(score.percent)}%`;
    dom.progressFill.style.width = `${completionPercent}%`;
    dom.progressLabel.textContent = `${completionPercent}%`;

    const lowest = getLowestScoredItems(active.id, 3);
    const container = dom.resultNarrative;
    container.innerHTML = "";

    const p1 = document.createElement("p");
    p1.textContent = getScoreMessage(score.avg);
    container.appendChild(p1);

    if (lowest.length > 0) {
      const p2 = document.createElement("p");
      p2.textContent = "Priority areas (lowest scored):";
      container.appendChild(p2);

      const list = document.createElement("ul");
      lowest.forEach((item) => {
        const li = document.createElement("li");
        li.textContent = `${item.score}/5 - ${item.question}`;
        list.appendChild(li);
      });
      container.appendChild(list);
    }

    if (questionTotal > 0 && score.count === questionTotal) {
      setSummaryCollapsed(false);
    }
  }

  function gatherSectionResponses(section) {
    const categories = section.categories.map((category, cIdx) => {
      const questions = category.questions.map((question, qIdx) => {
        const name = `${section.id}-${cIdx}-${qIdx}-score`;
        const commentId = `${section.id}-${cIdx}-${qIdx}-comment`;
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

  function gatherAllResponses() {
    return sections.map(gatherSectionResponses);
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function buildReportHtml() {
    const orgName = dom.orgName.value.trim() || "Not provided";
    const completedBy = dom.completedBy.value.trim() || "Not provided";
    const completedDate = dom.completedDate.value || new Date().toISOString().slice(0, 10);

    const responses = gatherAllResponses();

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

  function buildReportPlainText() {
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

    gatherAllResponses().forEach((section) => {
      const totalQuestions = section.categories.reduce((sum, c) => sum + c.questions.length, 0);
      lines.push(section.sectionLabel);
      lines.push(`Scored: ${section.score.count}/${totalQuestions} | Average: ${section.score.avg.toFixed(2)} | Section Score: ${Math.round(section.score.percent)}%`);
      lines.push("");

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
    const html = buildReportHtml();
    const blob = new Blob(["\ufeff", html], {
      type: "application/msword"
    });
    downloadBlob(blob, `dropping-anchor-report-${new Date().toISOString().slice(0, 10)}.doc`);
  }

  function exportPdf() {
    const text = buildReportPlainText();
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

    pdf.save(`dropping-anchor-report-${new Date().toISOString().slice(0, 10)}.pdf`);
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
    applyMetadataToInputs();
    render();
  });

  loadSession();
  applyMetadataToInputs();
  render();
})();
