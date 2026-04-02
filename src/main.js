import "./styles.css";
import html2canvas from "html2canvas";

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

      <div class="tabs">
          <button class="tab-btn active" id="tabChatBtn">💬 ИИ Чат</button>
          <button class="tab-btn" id="tabCatalogBtn">📑 Каталог</button>
      </div>

      <!-- Вкладка ЧАТ -->
      <div id="viewChat" class="tab-view active">
        <div class="chat-area" id="chatArea">
            <div class="chat-welcome" id="chatWelcome" style="${state.messages.length > 0 ? 'display:none' : ''}">
            <div class="welcome-icon">🎨</div>
            <h2>Сгенерировать слайд</h2>
            <p>Напишите, что нужно на слайде (ИИ сам выберет дизайн из каталога).</p>
            <div class="quick-prompts">
                <button class="quick-prompt" data-prompt="Сделай 1 слайд про итоги года">📊 Итог года</button>
                <button class="quick-prompt" data-prompt="Слайд про новый продукт">🚀 Новый продукт</button>
            </div>
            </div>
            <div class="messages" id="messages"></div>
        </div>

        <div class="selection-bubble" id="selectionBubble" style="display:none">
            <div class="sb-header">✏️ Выделен текст на слайде (пишите ниже как исправить):</div>
            <div class="sb-text" id="selectionText"></div>
            <button class="btn-action" id="applySelectionBtn" style="display:none; margin-top: 5px;">✅ Применить к слайду</button>
        </div>

        <div class="input-area">
            <div class="input-wrapper">
            <textarea id="prompt" placeholder="Опишите что сделать..." rows="1"></textarea>
            <button class="send-btn" id="sendBtn" ${state.busy || !state.apiKey ? 'disabled' : ''}>
                <span class="send-icon">➤</span>
            </button>
            </div>
            <div class="input-hint">Enter — отправить, Shift+Enter — новая строка</div>
        </div>
      </div>

      <!-- Вкладка КАТАЛОГ -->
      <div id="viewCatalog" class="tab-view" style="display: none;">
         <div class="catalog-grid" id="catalogGrid"></div>
      </div>

      <!-- РЕДАКТОР СЛАЙДА (Открывается и из чата, и из каталога) -->
      <div class="slide-editor" id="slideEditor" style="display:none">
        <div class="editor-header">
          <span class="preview-title" id="editorTitle">Предпросмотр</span>
          <div class="preview-actions">
            <!-- Кнопка только одна: вставить в PPTX как картинку (складной HTML -> Image) -->
            <button class="btn-action" id="insertEditedSlideBtn">📥 Вставить на слайд</button>
            <button class="btn-action btn-secondary" id="closeEditorBtn">✕</button>
          </div>
        </div>
        <!-- Инструкция: Кликайте на текст для изменения -->
        <div class="editor-hint">💡 Вы можете кликнуть на любой текст ниже и отредактировать его!</div>
        <div class="editor-frame-container" id="editorFrameContainer"></div>
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
    // Fallback
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(base64Png, {
        coercionType: Office.CoercionType.Image
      }, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.error("Fallback failed: " + asyncResult.error.message);
        }
        resolve();
      });
    });
  });
}

/* ─── Каталог и Редактор ─── */
function renderCatalogGrid() {
  const grid = document.getElementById("catalogGrid");
  if (!grid) return;
  grid.innerHTML = "";
  Object.keys(TEMPLATE_MAP).forEach(tid => {
    const item = document.createElement("div");
    item.className = "catalog-item";
    item.dataset.tid = tid;
    item.innerHTML = `<span>${tid}</span>`;
    item.addEventListener("click", () => openSlideEditor(tid, { title: "Ваш заголовок" }));
    grid.appendChild(item);
  });
}

async function openSlideEditor(templateId, fields) {
  setBusy(true);
  try {
    const html = await loadTemplate(templateId);
    let filled = fillTemplate(html, fields);
    
    // Inject contenteditable body
    filled = filled.replace('<body>', '<body contenteditable="true" spellcheck="false" style="outline:none;">');
    
    document.getElementById("editorTitle").textContent = templateId;
    const container = document.getElementById("editorFrameContainer");
    container.innerHTML = `<iframe id="livePreviewFrame" srcdoc="${escapeHtml(filled)}"></iframe>`;
    document.getElementById("slideEditor").style.display = "flex";
  } catch(e) {
    console.error(e);
  } finally {
    setBusy(false);
  }
}

document.getElementById("insertEditedSlideBtn")?.addEventListener("click", async () => {
    const iframe = document.getElementById("livePreviewFrame");
    if (!iframe) return;
    setBusy(true);
    try {
        const doc = iframe.contentDocument || iframe.contentWindow.document;
        document.getElementById("insertEditedSlideBtn").textContent = "Сохранение...";
        // Deselect editable so caret isn't captured
        doc.body.blur(); 
        const canvas = await html2canvas(doc.body, { width: 1920, height: 1080, scale: 1 });
        const b64 = canvas.toDataURL("image/png").replace(/^data:image\/png;base64,/, "");
        await insertImageSlide(b64);
        document.getElementById("slideEditor").style.display = "none";
        addMessage("assistant", "✅ Слайд успешно вставлен в презентацию!");
    } catch(e) {
        alert("Ошибка вставки: " + e.message);
    } finally {
        document.getElementById("insertEditedSlideBtn").textContent = "📥 Вставить на слайд";
        setBusy(false);
    }
});
document.getElementById("closeEditorBtn")?.addEventListener("click", () => {
    document.getElementById("slideEditor").style.display = "none";
});


/* ─── AI и Сгенерировать слайд ─── */
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
  if (data.error) throw new Error(data.error.message);
  return data.choices[0].message.content;
}

async function generateSingleSlide(prompt) {
  setBusy(true);
  addMessage("user", prompt);
  addMessage("progress", "AI выбирает дизайн и пишет тексты...");
  try {
    const SINGLE_SLIDE_PROMPT = SYSTEM_PROMPT + "\n\nВНИМАНИЕ: Сгенерируй РОВНО 1 (ОДИН) слайд под запрос пользователя. Выбери самый подходящий шаблон.";
    const resp = await callLLM([{ role: "system", content: SINGLE_SLIDE_PROMPT }, { role: "user", content: prompt }]);
    const plan = JSON.parse(resp.match(/\{[\s\S]*\}/)[0]);
    if (plan.slides && plan.slides.length > 0) {
      removeProgress();
      addMessage("assistant", `✅ Слайд готов! Можете отредактировать текст перед вставкой.`);
      openSlideEditor(plan.slides[0].templateId, plan.slides[0].fields);
    }
  } catch (e) {
    removeProgress();
    addMessage("assistant", `❌ Ошибка: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

/* ─── Обработчик изменения выделения (Авто-ассистент) ─── */
function onSelectionChanged() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value.trim().length > 0) {
        state.selectionText = result.value.trim();
        document.getElementById("selectionBubble").style.display = "block";
        document.getElementById("selectionText").textContent = state.selectionText;
        document.getElementById("applySelectionBtn").style.display = "none"; // Hide until AI replies
      }
    });
}


/* ─── События ─── */
function attachEvents() {
  /* Вкладки */
  document.getElementById("tabChatBtn")?.addEventListener("click", () => {
    document.getElementById("tabChatBtn").classList.add("active");
    document.getElementById("tabCatalogBtn").classList.remove("active");
    document.getElementById("viewChat").style.display = "flex";
    document.getElementById("viewCatalog").style.display = "none";
  });
  document.getElementById("tabCatalogBtn")?.addEventListener("click", () => {
    document.getElementById("tabCatalogBtn").classList.add("active");
    document.getElementById("tabChatBtn").classList.remove("active");
    document.getElementById("viewCatalog").style.display = "flex";
    document.getElementById("viewChat").style.display = "none";
    if (document.getElementById("catalogGrid").innerHTML === "") {
        renderCatalogGrid();
    }
  });

  /* Чат и Отправка */
  document.getElementById("sendBtn")?.addEventListener("click", async () => {
    const val = document.getElementById("prompt").value.trim();
    if (!val) return;
    document.getElementById("prompt").value = "";
    
    // Если есть активное выделение текста, то это запрос ассистенту
    if (document.getElementById("selectionBubble").style.display === "block" && state.selectionText) {
        setBusy(true);
        addMessage("user", val);
        addMessage("progress", "Думаю над текстом...");
        try {
            const messages = [
               { role: "system", content: "Ты ИИ встроенный в интерфейс. Тебе дадут исходный текст и команду. Возвращай ТОЛЬКО финальный переделанный текст. Никаких комментариев." },
               { role: "user", content: `Исходный текст: "${state.selectionText}"\nКоманда: ${val}` }
            ];
            const newText = await callLLM(messages);
            removeProgress();
            
            // Запоминаем готовый текст чтобы применить по кнопке
            state.aiGeneratedText = newText.trim();
            addMessage("assistant", `Готовый текст:\n${state.aiGeneratedText}`);
            document.getElementById("applySelectionBtn").style.display = "block";
            
        } catch (e) {
            removeProgress();
            addMessage("assistant", "Ошибка: " + e.message);
        } finally { setBusy(false); }
    } else {
        // Обычная генерация 1 слайда
        generateSingleSlide(val);
    }
  });

  /* Кнопка "Применить выделение" */
  document.getElementById("applySelectionBtn")?.addEventListener("click", () => {
      if (!state.aiGeneratedText) return;
      Office.context.document.setSelectedDataAsync(state.aiGeneratedText, { coercionType: Office.CoercionType.Text }, (result) => {
         if (result.status === Office.AsyncResultStatus.Succeeded) {
             document.getElementById("selectionBubble").style.display = "none";
             state.selectionText = "";
             addMessage("assistant", "✅ Успешно заменено на слайде!");
         } else {
             alert("Не удалось заменить текст: " + result.error.message);
         }
      });
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
    setBusy(state.busy);
    setStatus(state.apiKey ? "✅" : "⚠️");
  });
  document.querySelectorAll(".quick-prompt").forEach(b => {
    b.addEventListener("click", () => {
        document.getElementById("prompt").value = b.dataset.prompt;
        document.getElementById("sendBtn").click();
    });
  });
}

function removeDebugLog() {
    const origConsoleError = window.console.error;
    window.console.error = function() {
        return origConsoleError.apply(this, arguments);
    };
    // We no longer display a debug div. Errors go to Assistant Chat directly if caught.
}

/* ─── Инициализация ─── */
async function bootstrap() {
  removeDebugLog();
  renderApp(); 
  attachEvents();
  try { await loadCatalog(); } catch (e) { console.error("Catalog Error", e); }
  
  if (!window.Office) { 
    state.previewMode = true; 
    return; 
  }

  Office.onReady(info => {
    state.officeReady = true;
    state.previewMode = info.host !== Office.HostType.PowerPoint;
    if (info.host === Office.HostType.PowerPoint) {
        Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    }
  });
}

bootstrap();