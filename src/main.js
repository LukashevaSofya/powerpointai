import "./styles.css";
import html2canvas from "html2canvas";

/* ═══════════════════════════════════════════
   TEMPLATE DEFINITIONS
   ═══════════════════════════════════════════ */
const TEMPLATES = [
  { id: "template-1", file: "cover_2lines.html", name: "Обложка — 2 строки", category: "обложка" },
  { id: "template-2", file: "cover_3lines.html", name: "Обложка — 3 строки", category: "обложка" },
  { id: "template-3", file: "divider_text.html", name: "Разделитель (синий)", category: "разделитель" },
  { id: "template-4", file: "text_block.html", name: "Текстовый блок", category: "контент" },
  { id: "template-5", file: "divider_covers.html", name: "Разделитель (тёмный)", category: "разделитель" },
  { id: "template-6", file: "long_list.html", name: "Длинный список", category: "контент" },
  { id: "template-7", file: "slide_image_1.html", name: "Картинка + текст (1)", category: "контент" },
  { id: "template-8", file: "slide_image_2.html", name: "Картинка + текст (2)", category: "контент" },
  { id: "template-9", file: "bullets_slide.html", name: "Буллеты", category: "контент" },
  { id: "template-10", file: "timeline_slide.html", name: "Таймлайн", category: "контент" },
  { id: "template-11", file: "table_slide.html", name: "Таблица", category: "контент" },
];

const PREVIEW_DATA = {
  "template-1": { title: "Финансовые\nрезультаты", subtitle: "Итоги 2024 года", name: "Иван Иванов", position: "Финансовый аналитик" },
  "template-2": { title: "Стратегия\nразвития", subtitle: "На 2025–2027 годы", description: "Ключевые направления роста компании", name: "Иван Иванов", position: "Аналитик" },
  "template-3": { divider_text: "Раздел\nпервый" },
  "template-4": { title: "Обзор рынка", paragraph1: "Российский фондовый рынок продолжает демонстрировать устойчивый рост на фоне улучшения макроэкономических показателей и растущего интереса частных инвесторов." },
  "template-5": { divider_text: "Итоги\nгода" },
  "template-6": { title: "Ключевые задачи", list_item_1: "Анализ рыночных тенденций", list_item_2: "Оптимизация портфеля", list_item_3: "Разработка стратегии", list_item_4: "Управление рисками", list_item_5: "Контроль качества", list_item_6: "Обучение персонала" },
  "template-7": { title: "О компании", paragraph1: "Финам — ведущая инвестиционная компания России с более чем 30-летним опытом работы на финансовых рынках.", paragraph2: "Мы предоставляем полный спектр брокерских и инвестиционных услуг." },
  "template-8": { title: "Продукты", paragraph1: "Широкая линейка финансовых инструментов для любого уровня инвестора.", paragraph2: "Индивидуальный подход к каждому клиенту и персональное сопровождение." },
  "template-9": { title: "Преимущества", bullet_1: "Надежность и стабильность", bullet_2: "Профессиональная команда", bullet_3: "Современные технологии", bullet_4: "Индивидуальный подход", bullet_5: "Прозрачные условия" },
  "template-10": { title: "Этапы проекта", title_1: "Q1", text_1: "Исследование рынка", title_2: "Q2", text_2: "Разработка стратегии", title_3: "Q3", text_3: "Реализация плана", title_4: "Q4", text_4: "Подведение итогов" },
  "template-11": { title: "Показатели", col_1: "2020", col_2: "2021", col_3: "2022", col_4: "2023", col_5: "2024", col_6: "Δ%", row1_name: "Выручка", row1_c1: "15.2", row1_c2: "18.4", row1_c3: "21.1", row1_c4: "24.5", row1_c5: "28.3", row1_c6: "+15%", row2_name: "EBITDA", row2_c1: "3.1", row2_c2: "4.2", row2_c3: "5.8", row2_c4: "7.1", row2_c5: "8.9", row2_c6: "+25%", row3_name: "Чист. прибыль", row3_c1: "1.8", row3_c2: "2.5", row3_c3: "3.2", row3_c4: "4.0", row3_c5: "5.1", row3_c6: "+28%" },
};

/* ═══════════════════════════════════════════
   AI SYSTEM PROMPTS
   ═══════════════════════════════════════════ */
const SYSTEM_PROMPT_PRESENTATION = `Ты — генератор корпоративных презентаций компании Финам.
У тебя ровно 11 типов слайдов. Используй ТОЛЬКО их templateId.

1. template-1 — обложка (2 строки). Поля: title, subtitle, name, position
2. template-2 — обложка (3 строки). Поля: title, subtitle, description, name, position
3. template-3 — разделитель (синий фон). Поля: divider_text
4. template-4 — текстовый слайд. Поля: title, paragraph1
5. template-5 — разделитель (тёмный фон). Поля: divider_text
6. template-6 — длинный список (до 13 пунктов). Поля: title, list_item_1...list_item_13
7. template-7 — слайд с картинкой (текст слева). Поля: title, paragraph1, paragraph2
8. template-8 — слайд с картинкой (текст справа). Поля: title, paragraph1, paragraph2
9. template-9 — буллеты (до 8 пунктов). Поля: title, bullet_1...bullet_8
10. template-10 — таймлайн (4 колонки). Поля: title, title_1, text_1, title_2, text_2, title_3, text_3, title_4, text_4
11. template-11 — таблица (6 столбцов, до 9 строк). Поля: title, col_1..col_6, row1_name, row1_c1..row1_c6, row2_name...

Правила:
- title на обложках (template-1, template-2): СТРОГО 1-3 слова крупными буквами.
- subtitle: 3-6 слов.
- divider_text: 1-3 слова (крупный шрифт).
- paragraph1/paragraph2: развёрнутый текст 1-3 предложения.
- bullet_X / list_item_X: короткие фразы 3-7 слов.
- Количество слайдов: 5-10 (по теме, не растягивай искусственно).
- Начинай с обложки (template-1 или template-2).
- Используй разделители для логических секций.
- Шаблоны можно повторять если нужно.
- Неиспользуемые поля заполняй пустой строкой "".
- Возвращай ТОЛЬКО валидный JSON, без markdown-обёрток. Формат:
{"slides":[{"templateId":"template-1","fields":{"title":"...","subtitle":"..."}}]}`;

const SYSTEM_PROMPT_CHAT = `Ты — корпоративный AI-ассистент компании Финам, встроенный в PowerPoint. Помогай с текстами, идеями для презентаций, формулировками. Отвечай кратко и по делу на русском языке. Если пользователь попросит создать презентацию — просто сделай это, не задавай лишних вопросов.`;

const SYSTEM_PROMPT_SELECTION = `Ты — AI-редактор текста, встроенный в PowerPoint. Тебе дают исходный текст со слайда и команду пользователя. Верни ТОЛЬКО готовый текст-замену. Никаких комментариев, пояснений, кавычек — только финальный текст.`;

/* ═══════════════════════════════════════════
   STATE
   ═══════════════════════════════════════════ */
const state = {
  officeReady: false,
  previewMode: false,
  busy: false,
  activeTab: "chat",
  templateCache: {},
  previewCache: {},
  catalogRendered: false,
  displayMessages: [],
  chatHistory: [],
  selectedText: null,
  lastAiReplacement: null,
  apiKey: localStorage.getItem("finam_api_key") || "",
  authorName: localStorage.getItem("finam_author_name") || "",
  authorPosition: localStorage.getItem("finam_author_position") || "",
  model: localStorage.getItem("finam_model") || "google/gemma-3-12b-it:free",
  settingsOpen: false,
};

/* ═══════════════════════════════════════════
   UTILITIES
   ═══════════════════════════════════════════ */
function escapeHtml(s) {
  return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}
function truncateWords(text, n = 5) {
  const words = text.trim().split(/\s+/);
  return words.length <= n ? words.join(" ") : words.slice(0, n).join(" ") + "…";
}

/* ═══════════════════════════════════════════
   RENDER APP
   ═══════════════════════════════════════════ */
function renderApp() {
  document.getElementById("app").innerHTML = `
    <div class="app-container">
      <header class="app-header">
        <div class="header-left">
          <div class="logo-text">ФИНАМ</div>
          <span class="header-subtitle">AI Ассистент</span>
        </div>
        <div class="header-right">
          <button class="icon-btn ${state.settingsOpen ? 'active' : ''}" id="settingsBtn" title="Настройки">⚙️</button>
        </div>
      </header>

      <div class="settings-panel ${state.settingsOpen ? 'open' : ''}" id="settingsPanel">
        <div class="settings-inner">
          <div class="settings-group">
            <label>API ключ OpenRouter</label>
            <input type="password" id="apiKeyInput" placeholder="sk-or-v1-..." value="${escapeHtml(state.apiKey)}" />
          </div>
          <div class="settings-group">
            <label>Модель AI</label>
            <select id="modelSelect">
              <option value="google/gemma-3-12b-it:free" ${state.model==="google/gemma-3-12b-it:free"?"selected":""}>Gemma 3 12B (free)</option>
              <option value="openrouter/free" ${state.model==="openrouter/free"?"selected":""}>Авто-роутер (free)</option>
              <option value="openai/gpt-4o" ${state.model==="openai/gpt-4o"?"selected":""}>GPT-4o (платная)</option>
              <option value="anthropic/claude-3.5-sonnet" ${state.model==="anthropic/claude-3.5-sonnet"?"selected":""}>Claude 3.5 Sonnet (платная)</option>
            </select>
          </div>
          <div class="settings-row">
            <div class="settings-group">
              <label>Имя автора</label>
              <input type="text" id="authorNameInput" placeholder="Иван Иванов" value="${escapeHtml(state.authorName)}" />
            </div>
            <div class="settings-group">
              <label>Должность</label>
              <input type="text" id="authorPosInput" placeholder="Аналитик" value="${escapeHtml(state.authorPosition)}" />
            </div>
          </div>
          <button class="btn-save" id="saveSettingsBtn">Сохранить</button>
        </div>
      </div>

      <div class="tabs">
        <button class="tab-btn ${state.activeTab==='chat'?'active':''}" data-tab="chat">💬 Чат</button>
        <button class="tab-btn ${state.activeTab==='catalog'?'active':''}" data-tab="catalog">📑 Шаблоны</button>
      </div>

      <div id="viewChat" class="tab-view ${state.activeTab==='chat'?'active':''}">
        <div class="chat-area" id="chatArea">
          <div class="chat-welcome" id="chatWelcome" style="${state.displayMessages.length > 0 ? 'display:none' : ''}">
            <div class="welcome-icon">🎯</div>
            <h2>AI Ассистент</h2>
            <p>Создавайте презентации, редактируйте текст на слайдах или задавайте вопросы</p>
            <div class="quick-prompts">
              <button class="quick-prompt" data-prompt="Создай презентацию про итоги года">📊 Итоги года</button>
              <button class="quick-prompt" data-prompt="Создай презентацию про новый продукт">🚀 Новый продукт</button>
              <button class="quick-prompt" data-prompt="Создай презентацию про стратегию развития">📈 Стратегия</button>
            </div>
          </div>
          <div class="messages" id="messagesContainer"></div>
        </div>

        <div class="selection-bubble" id="selectionBubble" style="display:none">
          <div class="sb-header">✏️ Выделенный текст:</div>
          <div class="sb-text" id="selectionText"></div>
        </div>

        <div class="input-area">
          <div class="input-wrapper">
            <textarea id="promptInput" placeholder="Напишите сообщение..." rows="1"></textarea>
            <button class="send-btn" id="sendBtn" ${state.busy?'disabled':''}>➤</button>
          </div>
          <div class="input-hint">Enter — отправить · Shift+Enter — новая строка</div>
        </div>
      </div>

      <div id="viewCatalog" class="tab-view ${state.activeTab==='catalog'?'active':''}">
        <div class="catalog-container">
          <div class="catalog-grid" id="catalogGrid"></div>
        </div>
      </div>
    </div>`;
}

/* ═══════════════════════════════════════════
   UI HELPERS
   ═══════════════════════════════════════════ */
function $(id) { return document.getElementById(id); }

function scrollChat() {
  const area = $("chatArea");
  if (area) area.scrollTop = area.scrollHeight;
}

function addMessage(role, content, opts = {}) {
  state.displayMessages.push({ role, content, ...opts });
  const container = $("messagesContainer");
  const welcome = $("chatWelcome");
  if (!container) return;
  if (welcome) welcome.style.display = "none";

  const div = document.createElement("div");
  div.className = `message message-${role}`;
  if (opts.id) div.id = opts.id;

  if (role === "progress") {
    div.innerHTML = `<div class="msg-content"><div class="progress-indicator"><div class="spinner"></div><span class="progress-text">${escapeHtml(content)}</span></div></div>`;
  } else if (role === "assistant" && opts.showReplace) {
    div.innerHTML = `<div class="msg-content">${escapeHtml(content)}<br><button class="btn-replace" id="btnReplace_${Date.now()}">✅ Заменить текст на слайде</button></div>`;
    setTimeout(() => {
      const btn = div.querySelector(".btn-replace");
      if (btn) btn.addEventListener("click", () => applyReplacement());
    }, 0);
  } else {
    div.innerHTML = `<div class="msg-content">${role === 'assistant' ? formatAssistantText(content) : escapeHtml(content)}</div>`;
  }

  container.appendChild(div);
  scrollChat();
}

function formatAssistantText(text) {
  return escapeHtml(text).replace(/\n/g, "<br>");
}

function showProgress(text) {
  const existing = $("currentProgress");
  if (existing) {
    existing.querySelector(".progress-text").textContent = text;
  } else {
    addMessage("progress", text, { id: "currentProgress" });
  }
}

function hideProgress() {
  const el = $("currentProgress");
  if (el) el.remove();
}

function setBusy(v) {
  state.busy = v;
  const btn = $("sendBtn");
  if (btn) btn.disabled = v;
}

function showNotification(text, type = "warning") {
  const existing = document.querySelector(".notification");
  if (existing) existing.remove();
  const el = document.createElement("div");
  el.className = `notification ${type}`;
  el.textContent = text;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 3500);
}

/* ═══════════════════════════════════════════
   TEMPLATE LOADING
   ═══════════════════════════════════════════ */
async function loadTemplate(templateId) {
  if (state.templateCache[templateId]) return state.templateCache[templateId];
  const tmpl = TEMPLATES.find(t => t.id === templateId);
  if (!tmpl) throw new Error("Unknown template: " + templateId);
  const res = await fetch(`/templates/${tmpl.file}`);
  if (!res.ok) throw new Error(`Failed to load ${tmpl.file}`);
  const html = await res.text();
  state.templateCache[templateId] = html;
  return html;
}

function fillTemplate(html, fields) {
  let result = html;
  const f = { ...fields };
  if (state.authorName && !f.name) f.name = state.authorName;
  if (state.authorPosition && !f.position) f.position = state.authorPosition;
  for (const [key, val] of Object.entries(f)) {
    result = result.replaceAll(`{{${key}}}`, String(val ?? ""));
  }
  return result.replace(/\{\{[^}]+\}\}/g, "");
}

/* ═══════════════════════════════════════════
   HTML → PNG RENDERING
   ═══════════════════════════════════════════ */
async function renderSlideToBase64(templateId, fields) {
  const html = await loadTemplate(templateId);
  const filled = fillTemplate(html, fields);

  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;left:-9999px;top:0;width:1920px;height:1080px;border:0;opacity:0;pointer-events:none;";
  document.body.appendChild(iframe);

  try {
    const doc = iframe.contentDocument;
    doc.open();
    doc.write(filled.replace("<head>",
      `<head><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">`
    ));
    doc.close();
    await new Promise(r => setTimeout(r, 900));
    const canvas = await html2canvas(doc.body, {
      width: 1920, height: 1080, scale: 1,
      useCORS: true, allowTaint: true, logging: false,
    });
    return canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
  } finally {
    document.body.removeChild(iframe);
  }
}

/* ═══════════════════════════════════════════
   CATALOG
   ═══════════════════════════════════════════ */
function renderCatalog() {
  const grid = $("catalogGrid");
  if (!grid || state.catalogRendered) return;
  state.catalogRendered = true;
  grid.innerHTML = "";

  TEMPLATES.forEach(tmpl => {
    const card = document.createElement("div");
    card.className = "catalog-card";
    card.innerHTML = `
      <div class="card-preview" id="prev-${tmpl.id}">
        <div class="preview-loading"><div class="spinner"></div></div>
      </div>
      <div class="card-info"><div class="card-name">${escapeHtml(tmpl.name)}</div></div>`;
    card.addEventListener("click", () => insertFromCatalog(tmpl.id));
    grid.appendChild(card);
  });

  // Generate iframe previews asynchronously
  generatePreviews();
}

async function generatePreviews() {
  for (const tmpl of TEMPLATES) {
    try {
      const html = await loadTemplate(tmpl.id);
      const filled = fillTemplate(html, PREVIEW_DATA[tmpl.id] || {});
      const baseUrl = window.location.origin;
      const absoluteHtml = filled.replace(/src="\//g, `src="${baseUrl}/`);

      const container = $(`prev-${tmpl.id}`);
      if (!container) continue;

      // Calculate scale based on container width
      const containerWidth = container.offsetWidth || 160;
      const scale = containerWidth / 1920;

      const iframe = document.createElement("iframe");
      iframe.style.cssText = `transform:scale(${scale});width:1920px;height:1080px;border:none;transform-origin:0 0;pointer-events:none;display:block;`;
      iframe.srcdoc = absoluteHtml;
      iframe.setAttribute("sandbox", "allow-same-origin");
      iframe.setAttribute("loading", "lazy");

      // Remove loading spinner and add iframe
      container.innerHTML = "";
      container.appendChild(iframe);
    } catch (e) {
      console.warn("Preview failed for", tmpl.id, e);
    }
  }
}

async function insertFromCatalog(templateId) {
  if (state.busy) return;
  setBusy(true);

  const card = document.querySelector(`[id="prev-${templateId}"]`)?.parentElement;
  let overlay;
  if (card) {
    overlay = document.createElement("div");
    overlay.className = "card-inserting";
    overlay.innerHTML = '<div class="spinner"></div> Вставка...';
    card.style.position = "relative";
    card.appendChild(overlay);
  }

  try {
    const b64 = await renderSlideToBase64(templateId, PREVIEW_DATA[templateId] || {});
    await createAndInsertPptx([{ base64: b64 }]);
    showNotification("✅ Слайд вставлен в презентацию!", "success");
  } catch (e) {
    showNotification("Ошибка вставки: " + e.message, "error");
  } finally {
    if (overlay) overlay.remove();
    setBusy(false);
  }
}

/* ═══════════════════════════════════════════
   POWERPOINT INTEGRATION
   ═══════════════════════════════════════════ */
async function createAndInsertPptx(slides) {
  if (!window.PptxGenJS) throw new Error("PptxGenJS не загружен");
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  for (const s of slides) {
    const slide = pres.addSlide();
    slide.addImage({ data: "image/png;base64," + s.base64, x: 0, y: 0, w: "100%", h: "100%" });
  }

  const pptxBase64 = await pres.write({ outputType: "base64" });

  if (state.previewMode) {
    await pres.writeFile({ fileName: "Finam_Presentation.pptx" });
    return;
  }

  return new Promise((resolve) => {
    if (Office.context.presentation?.insertSlidesFromBase64Async) {
      Office.context.presentation.insertSlidesFromBase64Async(pptxBase64, { formatting: "KeepSourceFormatting" }, (res) => {
        if (res.status === Office.AsyncResultStatus.Failed) {
          pres.writeFile({ fileName: "Finam_Presentation.pptx" }).then(resolve);
        } else {
          resolve();
        }
      });
    } else {
      pres.writeFile({ fileName: "Finam_Presentation.pptx" }).then(resolve);
    }
  });
}

function applyReplacement() {
  if (!state.lastAiReplacement || !state.selectedText) return;
  if (state.previewMode) {
    showNotification("Замена работает только в PowerPoint", "warning");
    return;
  }
  Office.context.document.setSelectedDataAsync(state.lastAiReplacement, { coercionType: Office.CoercionType.Text }, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      addMessage("assistant", "✅ Текст заменён на слайде!");
      state.lastAiReplacement = null;
      hideSelectionBubble();
    } else {
      showNotification("Не удалось заменить текст", "error");
    }
  });
}

/* ═══════════════════════════════════════════
   AI / LLM
   ═══════════════════════════════════════════ */
async function callLLM(messages) {
  const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: "Bearer " + state.apiKey },
    body: JSON.stringify({ model: state.model, messages, temperature: 0.7 }),
  });
  if (!res.ok) {
    const errBody = await res.text().catch(() => "");
    throw new Error(`API ошибка ${res.status}: ${errBody.slice(0, 200)}`);
  }
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data.choices[0].message.content;
}

function isPresentationRequest(text) {
  const lower = text.toLowerCase();
  const hasTarget = /(презентац|слайд)/i.test(lower);
  const hasAction = /(созда|сделай|сгенерир|подготов|составь|напиши|сделать|генер)/i.test(lower);
  return hasTarget && hasAction;
}

/* ═══════════════════════════════════════════
   CHAT HANDLERS
   ═══════════════════════════════════════════ */
async function handleSend() {
  const input = $("promptInput");
  const text = input?.value.trim();
  if (!text || state.busy) return;

  if (!state.apiKey) {
    showNotification("Вставьте API ключ в настройках ⚙️", "warning");
    return;
  }

  input.value = "";
  input.style.height = "auto";

  if (state.selectedText) {
    await handleSelectionEdit(text);
  } else if (isPresentationRequest(text)) {
    await handlePresentationGeneration(text);
  } else {
    await handleGeneralChat(text);
  }
}

async function handleGeneralChat(userPrompt) {
  setBusy(true);
  addMessage("user", userPrompt);
  state.chatHistory.push({ role: "user", content: userPrompt });
  showProgress("Думаю...");

  try {
    const messages = [
      { role: "system", content: SYSTEM_PROMPT_CHAT },
      ...state.chatHistory.slice(-10),
    ];
    const reply = await callLLM(messages);
    hideProgress();
    addMessage("assistant", reply);
    state.chatHistory.push({ role: "assistant", content: reply });
  } catch (e) {
    hideProgress();
    addMessage("assistant", "❌ Ошибка: " + e.message);
  } finally {
    setBusy(false);
  }
}

async function handleSelectionEdit(userPrompt) {
  setBusy(true);
  addMessage("user", userPrompt);
  showProgress("Редактирую текст...");

  try {
    const messages = [
      { role: "system", content: SYSTEM_PROMPT_SELECTION },
      { role: "user", content: `Исходный текст: "${state.selectedText}"\nКоманда: ${userPrompt}` },
    ];
    const newText = await callLLM(messages);
    hideProgress();
    state.lastAiReplacement = newText.trim();
    addMessage("assistant", state.lastAiReplacement, { showReplace: true });
  } catch (e) {
    hideProgress();
    addMessage("assistant", "❌ Ошибка: " + e.message);
  } finally {
    setBusy(false);
  }
}

async function handlePresentationGeneration(userPrompt) {
  setBusy(true);
  addMessage("user", userPrompt);
  showProgress("AI планирует структуру презентации...");

  try {
    const prompt = state.authorName
      ? userPrompt + `\n\nАвтор: ${state.authorName}, ${state.authorPosition}`
      : userPrompt;

    const resp = await callLLM([
      { role: "system", content: SYSTEM_PROMPT_PRESENTATION },
      { role: "user", content: prompt },
    ]);

    // Parse JSON from response
    const jsonMatch = resp.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error("AI не вернул JSON. Попробуйте ещё раз.");
    const plan = JSON.parse(jsonMatch[0]);
    if (!plan.slides?.length) throw new Error("AI не создал слайды. Попробуйте другой запрос.");

    addMessage("assistant", `📋 План: ${plan.slides.length} слайдов. Начинаю генерацию...`);

    const renderedSlides = [];
    for (let i = 0; i < plan.slides.length; i++) {
      const slideData = plan.slides[i];
      showProgress(`Генерация слайда ${i + 1} из ${plan.slides.length}...`);
      try {
        const b64 = await renderSlideToBase64(slideData.templateId, slideData.fields);
        renderedSlides.push({ base64: b64 });
      } catch (e) {
        console.warn("Пропуск слайда", i + 1, e);
      }
    }

    if (renderedSlides.length === 0) throw new Error("Не удалось сгенерировать ни одного слайда");

    showProgress("Собираю файл презентации...");
    await createAndInsertPptx(renderedSlides);
    hideProgress();
    addMessage("assistant", `✅ Презентация из ${renderedSlides.length} слайдов готова и вставлена!`);
  } catch (e) {
    hideProgress();
    addMessage("assistant", "❌ Ошибка генерации: " + e.message);
  } finally {
    setBusy(false);
  }
}

/* ═══════════════════════════════════════════
   SELECTION HANDLING
   ═══════════════════════════════════════════ */
function onSelectionChanged() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded && result.value?.trim()) {
      const text = result.value.trim();
      state.selectedText = text;
      showSelectionBubble(text);
    } else {
      hideSelectionBubble();
    }
  });
}

function showSelectionBubble(text) {
  const bubble = $("selectionBubble");
  const textEl = $("selectionText");
  if (!bubble || !textEl) return;
  textEl.textContent = truncateWords(text, 6);
  bubble.style.display = "block";
}

function hideSelectionBubble() {
  state.selectedText = null;
  state.lastAiReplacement = null;
  const bubble = $("selectionBubble");
  if (bubble) bubble.style.display = "none";

  // Remove all replace buttons
  document.querySelectorAll(".btn-replace").forEach(b => b.remove());
}

/* ═══════════════════════════════════════════
   EVENT BINDING
   ═══════════════════════════════════════════ */
function attachEvents() {
  // Tabs
  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      const tab = btn.dataset.tab;
      state.activeTab = tab;
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.toggle("active", b.dataset.tab === tab));
      $("viewChat").classList.toggle("active", tab === "chat");
      $("viewCatalog").classList.toggle("active", tab === "catalog");
      if (tab === "catalog") renderCatalog();
    });
  });

  // Settings
  $("settingsBtn")?.addEventListener("click", () => {
    state.settingsOpen = !state.settingsOpen;
    $("settingsPanel")?.classList.toggle("open", state.settingsOpen);
    $("settingsBtn")?.classList.toggle("active", state.settingsOpen);
  });

  $("saveSettingsBtn")?.addEventListener("click", () => {
    state.apiKey = $("apiKeyInput")?.value || "";
    state.model = $("modelSelect")?.value || state.model;
    state.authorName = $("authorNameInput")?.value || "";
    state.authorPosition = $("authorPosInput")?.value || "";
    localStorage.setItem("finam_api_key", state.apiKey);
    localStorage.setItem("finam_model", state.model);
    localStorage.setItem("finam_author_name", state.authorName);
    localStorage.setItem("finam_author_position", state.authorPosition);
    state.settingsOpen = false;
    $("settingsPanel")?.classList.remove("open");
    $("settingsBtn")?.classList.remove("active");
    showNotification("✅ Настройки сохранены", "success");
  });

  // Send
  $("sendBtn")?.addEventListener("click", handleSend);

  // Textarea
  const textarea = $("promptInput");
  if (textarea) {
    textarea.addEventListener("keydown", (e) => {
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    });
    textarea.addEventListener("input", () => {
      textarea.style.height = "auto";
      textarea.style.height = Math.min(textarea.scrollHeight, 120) + "px";
    });
  }

  // Quick prompts
  document.querySelectorAll(".quick-prompt").forEach(btn => {
    btn.addEventListener("click", () => {
      const input = $("promptInput");
      if (input) {
        input.value = btn.dataset.prompt;
        handleSend();
      }
    });
  });
}

/* ═══════════════════════════════════════════
   BOOTSTRAP
   ═══════════════════════════════════════════ */
async function init() {
  renderApp();
  attachEvents();

  if (!window.Office) {
    state.previewMode = true;
    console.log("Preview mode (no Office.js)");
    return;
  }

  Office.onReady((info) => {
    state.officeReady = true;
    state.previewMode = info.host !== Office.HostType.PowerPoint;
    if (info.host === Office.HostType.PowerPoint) {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        onSelectionChanged
      );
    }
  });
}

init();