const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const $ = (id) => document.getElementById(id);
const loading = $("loading");
const dropzone = $("dropzone");
const input = $("files");
const chosen = $("chosen");
const convertButton = $("convert");
const resetButton = $("reset");
const outcome = $("outcome");

const worker = new Worker("worker.js", { type: "module" });

let files = [];
let ready = false;
let busy = false;
let downloadUrl = null;

function show(kind, text) {
  const box = document.createElement("p");
  box.className = `message ${kind}`;
  box.setAttribute("role", kind === "error" ? "alert" : "status");
  box.textContent = text;
  outcome.replaceChildren(box);
}

function clearOutcome() {
  outcome.replaceChildren();
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
}

function render() {
  chosen.replaceChildren(
    ...files.map((file) => {
      const item = document.createElement("li");
      item.textContent = file.name;
      return item;
    }),
  );
  convertButton.disabled = busy || files.length === 0;
  resetButton.disabled = busy || (files.length === 0 && !outcome.hasChildNodes());
}

async function convert() {
  if (busy || files.length === 0) return;
  busy = true;
  clearOutcome();
  show(
    "",
    ready
      ? "Conversione in corso."
      : "Conversione in attesa: partirà appena il programma di conversione è caricato.",
  );
  render();
  const buffers = await Promise.all(files.map((file) => file.arrayBuffer()));
  worker.postMessage(
    { type: "convert", files: files.map((file, i) => ({ name: file.name, buffer: buffers[i] })) },
    buffers,
  );
}

function showResult(result) {
  busy = false;
  clearOutcome();
  if (!result.ok) {
    show("error", result.message);
    render();
    return;
  }
  downloadUrl = URL.createObjectURL(new Blob([result.buffer], { type: XLSX_TYPE }));
  const done = document.createElement("p");
  done.className = "message ok";
  done.textContent = "Conversione completata.";
  const link = document.createElement("a");
  link.className = "download";
  link.id = "download";
  link.href = downloadUrl;
  link.download = result.name;
  link.textContent = `Scaricare ${result.name}`;
  const parts = [done, link];
  if (result.warnings.length > 0) {
    const box = document.createElement("section");
    box.className = "warnings";
    box.id = "warnings";
    const title = document.createElement("h2");
    title.textContent = "Attenzione";
    const list = document.createElement("ul");
    list.replaceChildren(
      ...result.warnings.map((warning) => {
        const item = document.createElement("li");
        item.textContent = warning;
        return item;
      }),
    );
    box.append(title, list);
    parts.push(box);
  }
  outcome.replaceChildren(...parts);
  render();
}

function addFiles(added) {
  for (const file of added) {
    if (!files.some((known) => known.name === file.name && known.size === file.size)) {
      files.push(file);
    }
  }
  input.value = "";
  render();
  if (files.length >= 2) convert();
}

worker.onmessage = (event) => {
  const message = event.data;
  if (message.type === "ready") {
    ready = true;
    loading.hidden = true;
    if (busy) show("", "Conversione in corso.");
  } else if (message.type === "failed") {
    busy = false;
    clearOutcome();
    render();
    loading.className = "message error";
    loading.textContent =
      "Impossibile caricare il programma di conversione. Controllare il collegamento a " +
      "Internet, ricaricare la pagina e riprovare.";
  } else if (message.type === "result") {
    showResult(message);
  }
};

input.addEventListener("change", () => addFiles(input.files));
convertButton.addEventListener("click", convert);
resetButton.addEventListener("click", () => {
  files = [];
  busy = false;
  clearOutcome();
  render();
});

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("over");
});
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("over"));
dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("over");
  addFiles(event.dataTransfer.files);
});
// A file dropped beside the drop zone must not make the browser open it.
for (const type of ["dragover", "drop"]) {
  window.addEventListener(type, (event) => event.preventDefault());
}

render();
