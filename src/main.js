import "./styles.css";
import html2canvas from "html2canvas";
import { generateNativePptx } from "./pptx-generator.js";

/* ─── Карта шаблонов ─── */
const TEMPLATE_MAP = {
  "template-1": "cover_2lines.html",
  "template-2": "cover_3lines.html",
  "template-3": "divider_text.html",
  "template-4": "text_block.html",
  "template-5": "divider_covers.html",
  "template-6": "long_list.html",
  "template-7": "slide_image_1.html",
  "template-8": "slide_image_2.html",
  "template-9": "bullets_slide.html",
  "template-10": "timeline_slide.html",
  "template-11": "table_slide.html"
};

/* ─── Состояние ─── */
const state = {
  officeReady: false,
  previewMode: false,
  busy: false,
  catalog: [],
  templateCache: {},
  messages: [],
  apiKey: localStorage.getItem("finam_api_key") || "",
  authorName: localStorage.getItem("finam_author_name") || "",
  authorPosition: localStorage.getItem("finam_author_position") || "",
  model: localStorage.getItem("finam_model") || "google/gemma-3-12b-it:free",
  settingsOpen: false,
  lastSlides: null,
  debugInfo: ""
};

/* ─── Визуальный дебаггер для WebView2 ─── */
function logDebug(msg) {
  state.debugInfo += `> ${msg}\n`;
  console.log("[DEBUG]", msg);
  const debugEl = document.getElementById("debug-log");
  if (debugEl) {
    debugEl.textContent = state.debugInfo;
    debugEl.style.display = "block";
  }
}

window.onerror = (msg, url, line) => {
  logDebug(`JS Error: ${msg} at ${line}`);
};

/* ─── Промпт для AI ─── */
const SYSTEM_PROMPT = `Ты — генератор корпоративных презентаций компании Финам.
У тебя ровно 11 типов слайдов. Используй ТОЛЬКО их templateId.

1. template-1 — обложка (2 строки)
   Поля: title, subtitle, name, position

2. template-2 — обложка (3 строки)
   Поля: title, subtitle, description, name, position

3. template-3 — разделитель (синий фон)
   Поля: divider_text

4. template-4 — текстовый слайд
   Поля: title, paragraph1

5. template-5 — разделитель (тёмный фон)
   Поля: divider_text

6. template-6 — длинный список (до 13 пунктов)
   Поля: title, list_item_1, list_item_2, ... list_item_13

7. template-7 — слайд с картинкой (текст слева)
   Поля: title, paragraph1, paragraph2

8. template-8 — слайд с картинкой (текст сверху)
   Поля: title, paragraph1, paragraph2

9. template-9 — буллеты (до 8 пунктов)
   Поля: title, bullet_1, bullet_2, ... bullet_8

10. template-10 — таймлайн (4 колонки текста)
    Поля: title, title_1, text_1, title_2, text_2, title_3, text_3, title_4, text_4

11. template-11 — таблица показателей (6 столбцов, до 9 строк)
    Поля: title, col_1, col_2, col_3, col_4, col_5, col_6,  row1_name, row1_c1, row1_c2... row1_c6,  row2_name, row2_c1...row2_c6 (до row9...)

Правила:
- Крупные шрифты (например title) делай СТРОГО 2-3 слова.
- Мелкие шрифты (например подзаголовки) делай около 5-7 слов, пиши по логике.
- Создавай от 5 до 12 слайдов (сколько нужно для раскрытия темы, иногда хватит 5, иногда 7).
- Используй шаблоны многократно если нужно.
- Возвращай ТОЛЬКО валидный JSON — никакого текста до или после.
- Неиспользуемые поля заполняй пустой строкой "".

Формат ответа (СТРОГО JSON):
{
  "slides": [
    {
      "templateId": "template-1",
      "fields": {
        "title": "Краткий заголовок",
        "subtitle": "Подзаголовок из 5 слов",
        "name": "",
        "position": ""
      }
    }
  ]
}`;

/* ─── Рендер приложения ─── */
function renderApp() {
  document.querySelector("#app").innerHTML = `
    <div class="app-container">
      <header class="app-header">
        <div class="header-left">
          <div class="logo-text">ФИНАМ</div>
          <span class="header-title">AI Презентации</span>
        </div>
        <div class="header-right">
          <span class="status-badge" id="status">${state.apiKey ? "✅" : "⚠️"}</span>
          <button class="icon-btn" id="settingsBtn" title="Настройки">⚙️</button>
        </div>
      </header>

      <div class="settings-panel ${state.settingsOpen ? 'open' : ''}" id="settingsPanel">
        <div class="settings-group">
          <label>API ключ OpenRouter</label>
          <input type="password" id="apiKeyInput" placeholder="sk-or-v1-..." value="${escapeHtml(state.apiKey)}" />
        </div>
        <div class="settings-row">
          <div class="settings-group">
            <label>Имя автора</label>
            <input type="text" id="authorNameInput" placeholder="Иван Иванов" value="${escapeHtml(state.authorName)}" />
          </div>
          <div class="settings-group">
            <label>Должность</label>
            <input type="text" id="authorPositionInput" placeholder="Аналитик" value="${escapeHtml(state.authorPosition)}" />
          </div>
        </div>
        <div class="settings-group">
          <label>Модель AI</label>
          <select id="modelSelect">
            <option value="google/gemma-3-12b-it:free" ${state.model === "google/gemma-3-12b-it:free" ? "selected" : ""}>Gemma 3 12B (free)</option>
            <option value="openrouter/free" ${state.model === "openrouter/free" ? "selected" : ""}>Авто-роутер (free)</option>
            <option value="openai/gpt-4o" ${state.model === "openai/gpt-4o" ? "selected" : ""}>GPT-4o (платная)</option>
            <option value="anthropic/claude-3.5-sonnet" ${state.model === "anthropic/claude-3.5-sonnet" ? "selected" : ""}>Claude 3.5 Sonnet (платная)</option>
          </select>
        </div>
        <button class="btn-save-settings" id="saveSettingsBtn">Сохранить</button>
      </div>

      <div class="chat-area" id="chatArea">
        <div class="chat-welcome" id="chatWelcome" style="${state.messages.length > 0 ? 'display:none' : ''}">
          <div class="welcome-icon">🎨</div>
          <h2>Создайте презентацию</h2>
          <p>Опишите тему презентации — AI подберёт подходящие слайды Финам</p>
          <div class="quick-prompts">
            <button class="quick-prompt" data-prompt="Создай презентацию про итоги компании за 2024 год">📊 Итоги года</button>
            <button class="quick-prompt" data-prompt="Презентация нового продукта для клиентов">🚀 Новый продукт</button>
            <button class="quick-prompt" data-prompt="Обзор рынка акций и инвестиционные идеи на 2025 год">📈 Обзор рынка</button>
          </div>
        </div>
        <div class="messages" id="messages">
        </div>
      </div>

      <div class="slides-preview" id="slidesPreview" style="display:none">
        <div class="preview-header">
          <span class="preview-title" id="previewTitle">Превью слайдов</span>
          <div class="preview-actions">
            <button class="btn-action" id="insertAllBtn" title="Скачать как PPTX файл">📥 Скачать как PPTX</button>
            <button class="btn-action btn-secondary" id="closePreviewBtn">✕</button>
          </div>
        </div>
        <div class="preview-slides" id="previewSlides"></div>
      </div>

      <div class="input-area">
        <div class="input-wrapper">
          <textarea id="prompt" placeholder="Опишите тему презентации..." rows="1"></textarea>
          <button class="send-btn" id="sendBtn" ${state.busy || !state.apiKey ? 'disabled' : ''}>
            <span class="send-icon">➤</span>
          </button>
        </div>
        <div class="input-hint">Enter — отправить, Shift+Enter — новая строка</div>
      </div>

      <div class="assistant-panel">
        <div class="assistant-header">✨ AI Ассистент Слайда</div>
        <div class="assistant-content">
          <button id="pullTextBtn" class="btn-action" style="width: 100%; margin-bottom: 8px;">📥 Взять текст со слайда</button>
          <textarea id="assistantText" placeholder="Здесь появится выделенный в PowerPoint текст..." rows="3" readonly></textarea>
          <textarea id="assistantPrompt" placeholder="Что сделать с текстом? (например: сократи, перепиши официально)" rows="2"></textarea>
          <button id="applyTextBtn" class="btn-action" style="width: 100%; margin-top: 8px;" disabled>Заменить текст на слайде ✨</button>
        </div>
      </div>
    </div>
  `;
}

function escapeHtml(str) {
  return String(str).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/* ─── UI-хелперы ─── */
function setStatus(text) {
  const el = document.getElementById("status");
  if (el) el.textContent = text;
}

function addMessage(role, content) {
  state.messages.push({ role, content });
  const messagesEl = document.getElementById("messages");
  const welcomeEl = document.getElementById("chatWelcome");
  if (!messagesEl) return;

  if (welcomeEl) welcomeEl.style.display = "none";

  const msgDiv = document.createElement("div");
  msgDiv.className = `message message-${role}`;

  if (role === "progress") {
    msgDiv.innerHTML = `<div class="msg-content"><div class="progress-indicator"><div class="spinner"></div><span class="progress-text">${escapeHtml(content)}</span></div></div>`;
    msgDiv.id = "currentProgress";
  } else {
    msgDiv.innerHTML = `<div class="msg-content">${role === 'assistant' ? content : escapeHtml(content)}</div>`;
  }

  messagesEl.appendChild(msgDiv);
  messagesEl.scrollTop = messagesEl.scrollHeight;
}

function updateProgress(text) {
  const el = document.getElementById("currentProgress");
  if (el) {
    el.querySelector(".progress-text").textContent = text;
  } else {
    addMessage("progress", text);
  }
}

function removeProgress() {
  const el = document.getElementById("currentProgress");
  if (el) el.remove();
}

function setBusy(busy) {
  state.busy = busy;
  const btn = document.getElementById("sendBtn");
  if (btn) btn.disabled = busy || !state.apiKey;
}

/* ─── Загрузка каталога и шаблонов ─── */
async function loadCatalog() {
  const res = await fetch("/slide-catalog.json");
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  state.catalog = await res.json();
}

async function loadTemplate(templateId) {
  if (state.templateCache[templateId]) return state.templateCache[templateId];
  const file = TEMPLATE_MAP[templateId];
  if (!file) throw new Error(`Неизвестный templateId: ${templateId}`);
  const res = await fetch(`/templates/${file}`);
  if (!res.ok) throw new Error(`Не удалось загрузить шаблон ${file}`);
  const html = await res.text();
  state.templateCache[templateId] = html;
  return html;
}

/* ─── Заполнение шаблона ─── */
function fillTemplate(html, fields) {
  let result = html;
  const f = { ...fields };
  if (state.authorName && !f.name) f.name = state.authorName;
  if (state.authorPosition && !f.position) f.position = state.authorPosition;

  for (const [key, value] of Object.entries(f)) {
    result = result.replaceAll(`{{${key}}}`, String(value ?? ""));
  }
  return result.replace(/\{\{[^}]+\}\}/g, "");
}

/* ─── HTML → PNG ─── */
async function htmlToPngBase64(html) {
  const iframe = document.createElement("iframe");
  iframe.style.cssText = "position:fixed;left:-9999px;width:1920px;height:1080px;border:0;";
  document.body.appendChild(iframe);
  try {
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open(); doc.write(html); doc.close();
    await new Promise(r => setTimeout(r, 1500));
    const canvas = await html2canvas(doc.body, { width: 1920, height: 1080, scale: 1 });
    return canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
  } finally {
    document.body.removeChild(iframe);
  }
}

/* ─── Вставка в PPT ─── */
async function insertImageSlide(base64Png) {
  if (state.previewMode) return;
  await PowerPoint.run(async (context) => {
    const newSlide = context.presentation.slides.add();
    const image = newSlide.shapes.addImage(base64Png);
    image.left = 0; image.top = 0; image.width = 960; image.height = 540;
    await context.sync();
  }).catch(async (error) => {
    logDebug("PowerPoint.run add() failed, falling back to setSelectedDataAsync");
    // Fallback if the version of PowerPoint doesn't support slides.add() properly
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(base64Png, {
        coercionType: Office.CoercionType.Image
      }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          logDebug("Fallback failed: " + asyncResult.error.message);
          resolve(); 
        } else {
          resolve();
        }
      });
    });
  });
}

/* ─── AI ─── */
async function callLLM(messages) {
  const res = await fetch("https://openrouter.ai/api/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + state.apiKey,
    },
    body: JSON.stringify({ model: state.model, messages, temperature: 0.7 })
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  return data.choices[0].message.content;
}

async function generatePresentation(prompt) {
  setBusy(true);
  addMessage("user", prompt);
  addMessage("progress", "AI думает...");
  try {
    const resp = await callLLM([{ role: "system", content: SYSTEM_PROMPT }, { role: "user", content: prompt }]);
    const plan = JSON.parse(resp.match(/\{[\s\S]*\}/)[0]);
    state.lastSlides = plan.slides;
    removeProgress();
    addMessage("assistant", `✅ Готово! ${plan.slides.length} слайдов. Открываю превью...`);
    showSlidesPreview(plan.slides);
  } catch (e) {
    removeProgress();
    addMessage("assistant", `❌ Ошибка: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

async function showSlidesPreview(slides) {
  const previewEl = document.getElementById("slidesPreview");
  const container = document.getElementById("previewSlides");
  if (!previewEl || !container) return;
  container.innerHTML = "";
  previewEl.style.display = "flex";
  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];
    if (!TEMPLATE_MAP[slide.templateId]) {
      console.warn("AI сгенерировал неизвестный шаблон:", slide.templateId);
      continue;
    }
    const html = await loadTemplate(slide.templateId);
    const card = document.createElement("div");
    card.className = "preview-card";
    card.innerHTML = `
      <div class="preview-card-header">
        <span class="slide-number">${i + 1}</span>
        <span class="slide-type">${slide.templateId}</span>
      </div>
      <div class="preview-frame">
        <iframe srcdoc="${escapeHtml(fillTemplate(html, slide.fields))}"></iframe>
      </div>
    `;
    container.appendChild(card);
  }
}

async function insertAllSlides() {
  if (!state.lastSlides) return;
  setBusy(true);
  try {
    addMessage("progress", "Создаю презентацию...");
    await generateNativePptx(state.lastSlides);
    removeProgress();
    addMessage("assistant", "✅ Презентация успешно скачана! Откройте файл, чтобы редактировать текст.");
  } catch (e) {
    removeProgress();
    addMessage("assistant", `❌ Ошибка генерации: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

/* ─── События ─── */
function attachEvents() {
  /* ─── Ассистент Слайда ─── */
  document.getElementById("pullTextBtn")?.addEventListener("click", () => {
    if (!state.officeReady || state.previewMode) {
      alert("Доступно только внутри PowerPoint!");
      return;
    }
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        document.getElementById("assistantText").value = result.value;
        document.getElementById("applyTextBtn").disabled = false;
        logDebug("Extracted text: " + result.value);
      } else {
        logDebug("Get text error: " + result.error.message);
        alert("Ошибка! Выделите текст на слайде (зайдите внутрь текстового блока).");
      }
    });
  });

  document.getElementById("applyTextBtn")?.addEventListener("click", async () => {
    const originalText = document.getElementById("assistantText").value;
    const prompt = document.getElementById("assistantPrompt").value.trim();
    if (!originalText || !prompt) return;
    
    setBusy(true);
    document.getElementById("applyTextBtn").textContent = "Думаю...";
    try {
      const messages = [
        { role: "system", content: "Ты встроенный ИИ-помощник в PowerPoint. Тебе дадут исходный текст со слайда и команду. Измени текст согласно команде и верни ТОЛЬКО новый текст. Никаких комментариев." },
        { role: "user", content: `Исходный текст: "${originalText}"\nКоманда: ${prompt}` }
      ];
      const newText = await callLLM(messages);
      
      Office.context.document.setSelectedDataAsync(newText.trim(), { coercionType: Office.CoercionType.Text }, (result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
            document.getElementById("assistantText").value = newText.trim();
            logDebug("Text replaced successfully");
         } else {
            logDebug("Set text error: " + result.error.message);
            alert("Не удалось заменить текст.");
         }
      });
    } catch (e) {
       alert("Ошибка ИИ: " + e.message);
    } finally {
       setBusy(false);
       document.getElementById("applyTextBtn").textContent = "Заменить текст на слайде ✨";
    }
  });

  /* ─── Оригинальные обработчики ─── */
  document.getElementById("sendBtn")?.addEventListener("click", () => {
    const val = document.getElementById("prompt").value.trim();
    if (val) generatePresentation(val);
  });
  document.getElementById("settingsBtn")?.addEventListener("click", () => {
    state.settingsOpen = !state.settingsOpen;
    const panel = document.getElementById("settingsPanel");
    if (panel) panel.classList.toggle("open", state.settingsOpen);
  });
  document.getElementById("saveSettingsBtn")?.addEventListener("click", () => {
    state.apiKey = document.getElementById("apiKeyInput").value;
    state.authorName = document.getElementById("authorNameInput").value;
    state.authorPosition = document.getElementById("authorPositionInput").value;
    state.model = document.getElementById("modelSelect").value;
    localStorage.setItem("finam_api_key", state.apiKey);
    localStorage.setItem("finam_author_name", state.authorName);
    localStorage.setItem("finam_author_position", state.authorPosition);
    localStorage.setItem("finam_model", state.model);
    state.settingsOpen = false;
    document.getElementById("settingsPanel").classList.remove("open");
    setBusy(state.busy); // Update UI
    setStatus(state.apiKey ? "✅" : "⚠️");
  });
  document.getElementById("insertAllBtn")?.addEventListener("click", insertAllSlides);
  document.getElementById("closePreviewBtn")?.addEventListener("click", () => {
    document.getElementById("slidesPreview").style.display = "none";
  });
  document.querySelectorAll(".quick-prompt").forEach(b => {
    b.addEventListener("click", () => generatePresentation(b.dataset.prompt));
  });
}

/* ─── Инициализация ─── */
async function bootstrap() {
  logDebug("Bootstrap...");
  renderApp(); attachEvents();
  try { await loadCatalog(); logDebug("Catalog OK"); } catch (e) { logDebug("Catalog Error"); }
  
  if (!window.Office) { 
    state.previewMode = true; logDebug("No Office.js"); 
    return; 
  }

  Office.onReady(info => {
    logDebug("Office ready: " + info.host);
    state.officeReady = true;
    state.previewMode = info.host !== Office.HostType.PowerPoint;
  });
}

bootstrap();