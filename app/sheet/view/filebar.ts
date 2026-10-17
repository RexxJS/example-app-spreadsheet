// Files: the workbook saved to and opened from the page's storage by name
// (every sheet, its notes and names, in the Tauri app's format), the saved
// files listed; a new workbook; CSV out (what the selection, or everything
// used, shows; or TSV, or HTML) and in (at the selected cell, its delimiter
// detected; or JSON rows, or an HTML page's table); the Excel example, the Tauri
// app's three sample sheets (RexxJS formulas), an empty sheet, and the way
// home. As an installed app with `documents`, files the person picks too:
// a workbook in the Tauri app's format (.json, .rexxsheet), an XLSX (an
// encrypted one by the password field's password) or CSV opened,
// the workbook saved, the selection exported as CSV; the page never sees
// where they are.
import type { SheetView } from "./view.ts";

export function fileBar(v: SheetView): void {
  const name = state(""), csv = state(""), saved = state(""), shown = state("");
  const listed = async () => {
    const o = await v.cmds.execute("file.list");
    saved.set(o.message);
  };
  ui.hstack(() => {
    bind(ui.textfield("file name", () => {}), name);
    v.btn("Save", async () => { await v.run("file.save", name.value.trim() ? { name: name.value } : {}); await listed(); });
    v.btn("Open", () => v.run("file.open", { name: name.value }));
    v.btn("Delete file", async () => { await v.run("file.delete", { name: name.value }); await listed(); });
    v.btn("New", () => v.run("file.new"));
    bind(ui.text(""), saved);
  });
  ui.hstack(() => {
    v.btn("Export CSV", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportCsv", many ? { range: v.range() } : {});
      shown.set(String(o.output ?? ""));
    });
    v.btn("Export TSV", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportCsv", { ...(many ? { range: v.range() } : {}), delimiter: "tab" });
      shown.set(String(o.output ?? ""));
    });
    v.btn("Export PDF", () => v.run("file.exportPdf", { gridlines: "1" }));
    bind(ui.textarea("CSV"), csv);
    v.btn("Import CSV", () => v.run("file.importCsv", { at: v.here(), text: csv.value }));
    v.btn("Import JSON", () => v.run("file.importJson", { at: v.here(), text: csv.value }));
    v.btn("Import HTML", () => v.run("file.importHtml", { at: v.here(), text: csv.value }));
    v.btn("Export HTML", async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const o = await v.run("file.exportHtml", { ...(many ? { range: v.range() } : {}), header: "1" });
      shown.set(String(o.output ?? ""));
    });
  });
  bind(ui.text_wrapped("", 720), shown);
  if (typeof documents !== "undefined") documentsRow(v);
  ui.hstack(() => {
    v.btn("Paste example", () => v.run("example.paste"));
    v.btn("Budget (RexxJS)", () => v.run("file.open", { sample: "budget" }));
    v.btn("Test sheet (RexxJS)", () => v.run("file.open", { sample: "test" }));
    v.btn("Formatted (RexxJS)", () => v.run("file.open", { sample: "formatted" }));
    v.btn("Clear all", () => v.run("sheet.clear"));
    v.btn("Home", () => browserContext.changePage("/"));
  });
  listed();
}

// Open file, Save file and Export CSV file, through the person's picker.
function documentsRow(v: SheetView): void {
  // What failed, on the status line, as a refused command's reason is.
  const tried = (work: () => Promise<void>) => async () => {
    try {
      await work();
    } catch (e) {
      v.status.set(String((e as Error)?.message ?? e));
    }
  };
  // An encrypted XLSX's password, typed here before Open file (its text
  // shown nowhere).
  let password = "";
  ui.hstack(() => {
    ui.securefield("file password", (t: string) => { password = t; });
    v.btn("Open file", tried(async () => {
      const d = await documents.pick_open("Open a workbook, an XLSX, or a CSV file");
      if (!d) return;
      // An XLSX is bytes: read as base64 (sae's documents.read_base64).
      if (/\.xlsx$/i.test(d.name)) {
        await v.run("file.importXlsx", { base64: await d.read_base64(), name: d.name, ...(password ? { password } : {}) });
        return;
      }
      if (/\.xls$/i.test(d.name)) {
        await v.run("file.importXls", { base64: await d.read_base64(), name: d.name });
        return;
      }
      if (/\.ods$/i.test(d.name)) {
        await v.run("file.importOds", { base64: await d.read_base64(), name: d.name });
        return;
      }
      const text = await d.read_text();
      if (/\.(csv|tsv|tab|txt)$/i.test(d.name)) await v.run("file.importCsv", { at: v.here(), text, ...(/\.(tsv|tab)$/i.test(d.name) ? { delimiter: "tab" } : {}) });
      else if (/\.html?$/i.test(d.name)) await v.run("file.importHtml", { at: v.here(), text });
      // JSON rows (an array), or a workbook in the Tauri app's format.
      else if (/^\s*\[/.test(text)) await v.run("file.importJson", { at: v.here(), text });
      else await v.run("file.import", { text });
    }));
    v.btn("Save file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.json`);
      if (!d) return;
      const o = await v.cmds.execute("file.export");
      await d.write_text(String(o.output));
      print(`sheet: saved ${d.name}`);
      v.status.set(`Saved ${d.name}`);
    }));
    v.btn("Save XLSX file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.xlsx`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportXlsx");
      await d.write_base64(String(o.output));
      print(`sheet: saved ${d.name}`);
      v.status.set(`Saved ${d.name}${v.wb.fileReport.length ? `: ${v.wb.fileReport.join("; ")}` : ""}`);
    }));
    v.btn("Export PDF file", tried(async () => {
      const d = await documents.pick_save(`${v.wb.name}.pdf`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportPdf", { gridlines: "1" });
      await d.write_base64(String(o.output));
      print(`sheet: exported ${d.name}: ${o.message}`);
      v.status.set(`Exported ${d.name}`);
    }));
    v.btn("Export CSV file", tried(async () => {
      const s = v.selection(), many = s && (s[0].col !== s[1].col || s[0].row !== s[1].row);
      const d = await documents.pick_save(`${v.wb.name}.csv`);
      if (!d) return;
      const o = await v.cmds.execute("file.exportCsv", many ? { range: v.range() } : {});
      await d.write_text(String(o.output ?? ""));
      print(`sheet: exported ${d.name}`);
      v.status.set(`Exported ${d.name}`);
    }));
  });
}
