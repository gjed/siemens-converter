// Runs the converter in Pyodide, off the main thread. Everything it loads comes from
// this site: Pyodide, the openpyxl wheels and the converter source built from the
// same commit as the release.
import { loadPyodide } from "./pyodide/pyodide.mjs";

const UNEXPECTED =
  "La conversione non è riuscita per un errore imprevisto del programma. " +
  "Inviare i dettagli tecnici qui sotto a chi ha fornito il programma.";

// Same two calls as the command line: route the files by extension, then convert.
const GLUE = `
import json
from pathlib import Path

from siemens_converter.pipeline import ConversionError, convert, route_files


def run(paths, out_dir):
    try:
        fc_path, static_path = route_files([Path(p) for p in paths])
        result = convert(fc_path, static_path, Path(out_dir))
    except ConversionError as exc:
        return json.dumps({"ok": False, "message": str(exc)})
    return json.dumps(
        {"ok": True, "name": result.output_path.name, "warnings": result.warnings}
    )
`;

const here = (path) => new URL(path, self.location.href).href;

async function start() {
  const [config, source] = await Promise.all([
    fetch(here("config.json")).then((r) => r.json()),
    fetch(here("siemens_converter.zip")).then((r) => r.arrayBuffer()),
  ]);
  const pyodide = await loadPyodide({ indexURL: here("pyodide/") });
  await pyodide.loadPackage(["micropip", "packaging"]);
  pyodide.globals.set("wheel_urls", config.wheels.map(here));
  await pyodide.runPythonAsync(
    "import micropip\nawait micropip.install(wheel_urls.to_py(), deps=False)",
  );
  pyodide.FS.mkdirTree("/pkg");
  pyodide.unpackArchive(source, "zip", { extractDir: "/pkg" });
  pyodide.runPython("import sys\nsys.path.insert(0, '/pkg')\n" + GLUE);
  pyodide.FS.mkdirTree("/work/out");
  return pyodide;
}

const pyodide = start();
pyodide.then(
  () => postMessage({ type: "ready" }),
  (error) => postMessage({ type: "failed", detail: String(error) }),
);

async function convert(files) {
  const py = await pyodide;
  const FS = py.FS;
  const paths = [];
  files.forEach((file, i) => {
    FS.mkdirTree(`/work/in${i}`);
    const path = `/work/in${i}/${file.name}`;
    FS.writeFile(path, new Uint8Array(file.buffer));
    paths.push(path);
  });
  let outPath = null;
  try {
    const outcome = JSON.parse(py.globals.get("run")(py.toPy(paths), "/work/out"));
    if (!outcome.ok) return { type: "result", ok: false, message: outcome.message };
    outPath = `/work/out/${outcome.name}`;
    const bytes = FS.readFile(outPath).slice();
    return {
      type: "result",
      ok: true,
      name: outcome.name,
      warnings: outcome.warnings,
      buffer: bytes.buffer,
    };
  } catch (error) {
    return {
      type: "result",
      ok: false,
      message: `${UNEXPECTED}\n\nDettagli tecnici:\n${error}`,
    };
  } finally {
    for (const path of [...paths, outPath]) {
      if (path && FS.analyzePath(path).exists) FS.unlink(path);
    }
  }
}

onmessage = async (event) => {
  if (event.data.type !== "convert") return;
  let result;
  try {
    result = await convert(event.data.files);
  } catch (error) {
    result = { type: "result", ok: false, message: `${UNEXPECTED}\n\nDettagli tecnici:\n${error}` };
  }
  postMessage(result, result.buffer ? [result.buffer] : []);
};
