"use strict";

const $ = (id) => document.getElementById(id);
const run = $("run");
const kind = $("kind");
const supplier = $("supplier");
const fileInput = $("file");
const send = $("send");
const timeFormat = new Intl.DateTimeFormat("ru-RU", {
  dateStyle: "short", timeStyle: "short", timeZone: "Asia/Irkutsk",
});
const plural = new Intl.PluralRules("ru-RU");
const statusColors = {
  "Отчёт сформирован": "success",
  "Решения учтены": "success",
  "Загружены": "success",
  "Частично обработан": "attention",
  "Требует настройки": "attention",
  "Требует проверки": "attention",
  "Ошибка": "error",
};
const previous = new Map();
let state;
let refreshing = false;
let starting = false;
let sending = false;

function node(tag, content, className = "") {
  const element = document.createElement(tag);
  element.textContent = content;
  element.className = className;
  return element;
}

function keyedRender(key, values, render) {
  const signature = JSON.stringify(values);
  if (previous.get(key) === signature) return;
  previous.set(key, signature);
  render();
}

function status(value) {
  return node("span", value, statusColors[value] || "");
}

function row(...cells) {
  const entry = document.createElement("tr");
  entry.setAttribute("role", "row");
  for (const cell of cells) {
    const td = document.createElement("td");
    td.setAttribute("role", "cell");
    td.append(...(Array.isArray(cell) ? cell : [cell]));
    entry.append(td);
  }
  return entry;
}

function fileRow(entry) {
  const subject = [node("strong", entry.name), node("span", entry.supplier, "subline")];
  const counts = Object.entries(entry.counts).map(([label, count]) =>
    node("span", label + " · " + count, "subline"));
  const links = entry.downloads.map((kind) => {
    const link = node("a", kind === "report" ? "Отчёт XLSX" : "Выгрузка XLSX");
    link.href = "/download?" + new URLSearchParams({ kind, id: entry.id });
    return link;
  });
  return row(subject, status(entry.status), counts.length ? counts : "–", links.length ? links : "–");
}

function renderControls() {
  const review = kind.value === "review";
  $("supplier-field").hidden = review;
  supplier.disabled = review || !supplier.options.length;
  const extension = review ? "xlsx" : state?.profiles[supplier.value];
  fileInput.accept = extension ? "." + extension : ".xlsx,.csv,.xml";
  send.disabled = sending || !state || !fileInput.files.length || (!review && !extension);
  run.disabled = starting || !state || Boolean(state.issue) || Boolean(state.run?.running);
  run.textContent = state?.run?.running ? "Цикл выполняется…" : "Запустить цикл";
  $("upload-hint").textContent = review
    ? "Сохраните решения в Excel, затем загрузите отчёт XLSX."
    : "Формат " + (extension?.toUpperCase() || "по профилю поставщика") + " · до 10 МБ.";
}

function displayState(data) {
  state = data;
  const references = $("references");
  references.textContent = data.issue || "Справочники готовы · поставщиков " + Object.keys(data.profiles).length;
  references.className = data.issue ? "error" : "muted";
  keyedRender("profiles", data.profiles, () => {
    const selected = supplier.value;
    supplier.replaceChildren(...Object.keys(data.profiles).map((item) => new Option(item, item)));
    if (Object.hasOwn(data.profiles, selected)) supplier.value = selected;
  });
  const current = data.run;
  const latestCycle = data.journal.find((entry) => entry.stage === "Загрузка справочников");
  const label = $("run-status");
  label.textContent = current?.running ? "Робот обрабатывает очередь."
    : current ? (current.code === 0 ? "Цикл завершён" : "Ошибка запуска") + " · " + timeFormat.format(new Date(current.at))
    : latestCycle ? "Последний цикл · " + timeFormat.format(new Date(latestCycle.at))
    : "Циклов пока не было.";
  label.className = current?.code ? "error" : "";
  $("run-output").hidden = !current?.output;
  $("output").textContent = current?.output || "";

  const queued = data.incoming.map((entry) => entry.supplier + " · " + entry.name)
    .concat(data.waiting.map((name) => "Проверенный отчёт · " + name));
  const count = queued.length;
  const word = { one: "файл", few: "файла", many: "файлов", other: "файла" }[plural.select(count)];
  $("queue-label").textContent = "В очереди " + count + " " + word;
  keyedRender("queue", queued, () => {
    $("queue").replaceChildren(...queued.map((name) => node("li", name)));
  });
  $("file-count").textContent = data.files.length || "";
  $("files-empty").hidden = data.files.length > 0;
  $("journal-empty").hidden = data.journal.length > 0;
  keyedRender("files", data.files, () => $("files").replaceChildren(...data.files.map(fileRow)));
  keyedRender("journal", data.journal, () => {
    $("journal").replaceChildren(...data.journal.map((entry) => row(
      timeFormat.format(new Date(entry.at)), entry.stage, status(entry.status),
      entry.code && entry.code.length < 40 ? entry.code : "–",
    )));
  });
  renderControls();
}

async function request(url, options) {
  const response = await fetch(url, options);
  const value = await response.json();
  if (!response.ok) throw new Error(value.error || "Действие не выполнено.");
  return value;
}

async function refresh() {
  if (refreshing) return;
  refreshing = true;
  try {
    displayState(await request("/api/status"));
  } catch (error) {
    state = null;
    $("run-status").textContent = "Связь с панелью потеряна. " + error.message;
    $("run-status").className = "error";
    renderControls();
  } finally {
    refreshing = false;
  }
}

function feedback(message, success) {
  $("feedback").textContent = message;
  $("feedback").className = success ? "success" : "error";
  $("feedback").hidden = false;
}

run.addEventListener("click", async () => {
  starting = true;
  $("feedback").hidden = true;
  renderControls();
  try {
    await request("/api/run", { method: "POST", headers: { "X-DomTorg": "1" } });
    await refresh();
  } catch (error) {
    $("run-status").textContent = error.message;
    $("run-status").className = "error";
  } finally {
    starting = false;
    renderControls();
  }
});

$("upload-form").addEventListener("submit", async (event) => {
  event.preventDefault();
  const file = fileInput.files[0];
  if (!file || sending) return;
  sending = true;
  renderControls();
  try {
    const params = new URLSearchParams({ kind: kind.value, supplier: supplier.value, name: file.name });
    await request("/api/upload?" + params, {
      method: "POST", body: file,
      headers: { "Content-Type": "application/octet-stream", "X-DomTorg": "1" },
    });
    fileInput.value = "";
    feedback("Файл в очереди. Запустите цикл.", true);
    await refresh();
  } catch (error) {
    feedback(error.message, false);
  } finally {
    sending = false;
    renderControls();
  }
});

for (const input of [kind, supplier, fileInput]) input.addEventListener("change", renderControls);
$("refresh").addEventListener("click", refresh);

for (const id of ["files", "journal"]) {
  $(id + "-tab").addEventListener("click", () => {
    for (const other of ["files", "journal"]) {
      $(other + "-tab").setAttribute("aria-pressed", String(id === other));
      $(other + "-panel").hidden = id !== other;
    }
  });
}

refresh();
setInterval(refresh, 3000);
