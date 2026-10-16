// The menu bar (menus.ts), on the window of an installed app: each item
// runs the action of the panel button it is named for, or a key's (Copy,
// Paste), so it does just what the panel does; one with neither runs its
// command on the selection. In a web page there is no menu bar (sae
// refuses ui.menu there): the panels are the way in.
import { MENUS, shownLabel } from "../menus.ts";
import type { MenuItem } from "../menus.ts";
import type { SheetView } from "./view.ts";

export function menuBar(v: SheetView): void {
  const catalogue = new Map(v.cmds.catalogue().map((c) => [c.id, c.params]));
  // A command run on the selection: its range, cell or place from it.
  const fill: Record<string, () => string> = { range: () => v.range(), cell: () => v.here(), at: () => v.here() };
  const runOf = (e: MenuItem): (() => unknown) | null => {
    const action = v.actions.get(e.label);
    if (action) return action;
    const needs = (catalogue.get(e.id) ?? "").split(", ").filter((p) => p && !p.endsWith("?"));
    if (!catalogue.has(e.id) || needs.some((p) => !fill[p])) return null;
    return () => v.run(e.id, Object.fromEntries(needs.map((p) => [p, fill[p]()])));
  };
  let items = 0, left = 0;
  try {
    for (const m of MENUS) {
      ui.menu(m.label, () => {
        for (const e of m.items) {
          if (e === "-") { ui.menu_separator(); continue; }
          const run = runOf(e);
          if (!run) { left++; continue; }
          items++;
          ui.menu_item(shownLabel(e), () => run());
        }
      });
    }
  } catch (e) {
    print(`sheet: no menu bar here (${String((e as Error)?.message ?? e)})`);
    return;
  }
  print(`sheet: menus: ${MENUS.length} menus, ${items} items${left ? `, ${left} left out (no such panel here)` : ""}`);
}
