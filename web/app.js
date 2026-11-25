const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const $ = (id) => document.getElementById(id);
const loading = $("loading");
const convertButton = $("convert");
const resetButton = $("reset");
const outcome = $("outcome");

// One field per role: the field, not the extension, says which file is which.
const fields = Object.fromEntries(
  ["report", "static"].map((role) => [
    role,
    { zone: $(`${role}-zone`), input: $(role), name: $(`${role}-name`), file: null },
  ]),
);

const worker = new Worker("worker.js", { type: "module" });

let ready = false;
let busy = false;
let downloadUrl = null;

const complete = () => Object.values(fields).every((field) => field.file);
const empty = () => Object.values(fields).every((field) => !field.file);

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
  for (const field of Object.values(fields)) {
    field.name.textContent = field.file ? field.file.name : "Nessun file scelto";
    field.name.classList.toggle("none", !field.file);
  }
  convertButton.disabled = busy || !complete();
  resetButton.disabled = busy || (empty() && !outcome.hasChildNodes());
}

async function convert() {
  if (busy || !complete()) return;
  busy = true;
  clearOutcome();
  show(
    "",
    ready
      ? "Conversione in corso."
      : "Conversione in attesa: partirà appena il programma di conversione è caricato.",
  );
  render();
  const { report, static: data } = fields;
  let buffers;
  try {
    buffers = await Promise.all([report.file.arrayBuffer(), data.file.arrayBuffer()]);
  } catch {
    showResult({
      ok: false,
      message:
        "Impossibile leggere i file scelti, forse modificati dopo averli scelti. " +
        "Premere «Ricominciare», sceglierli di nuovo e ripetere la conversione.",
    });
    return;
  }
  worker.postMessage(
    {
      type: "convert",
      report: { name: report.file.name, buffer: buffers[0] },
      static: { name: data.file.name, buffer: buffers[1] },
    },
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

function choose(field, added) {
  field.input.value = "";
  if (added.length === 0) return;
  if (added.length > 1) {
    clearOutcome();
    show(
      "error",
      "Ogni riquadro accetta un solo file: il file FC_report nel primo, il file dei dati " +
        "statici nel secondo. Trascinare un file per volta.",
    );
    render();
    return;
  }
  // A new file replaces the field's one, so a corrected file is the one converted.
  field.file = added[0];
  render();
  if (complete()) convert();
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

convertButton.addEventListener("click", convert);
resetButton.addEventListener("click", () => {
  for (const field of Object.values(fields)) field.file = null;
  busy = false;
  clearOutcome();
  render();
});

for (const field of Object.values(fields)) {
  const { zone, input } = field;
  input.addEventListener("change", () => choose(field, [...input.files]));
  zone.addEventListener("dragover", (event) => {
    event.preventDefault();
    zone.classList.add("over");
  });
  zone.addEventListener("dragleave", () => zone.classList.remove("over"));
  zone.addEventListener("drop", (event) => {
    event.preventDefault();
    zone.classList.remove("over");
    choose(field, [...event.dataTransfer.files]);
  });
}
// A file dropped beside the fields must not make the browser open it.
for (const type of ["dragover", "drop"]) {
  window.addEventListener(type, (event) => event.preventDefault());
}

render();
