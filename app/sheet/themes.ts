// A workbook's theme: its colour roles (two darks, two lights, six
// accents, a hyperlink's) and its major and minor fonts, which looks,
// cell styles and table styles may name (fill=accent1), so a theme
// changed re-colours them. The theme is the page's, one workbook open at a
// time (as the 1904 date system's epoch is: dates.ts).
export const ROLES = ["dark1", "light1", "dark2", "light2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink"] as const;
export type Role = (typeof ROLES)[number];

export interface Theme {
  readonly name: string;
  readonly colors: Readonly<Record<Role, string>>;
  readonly fonts: { readonly major: string; readonly minor: string };
}

// Excel's default (Office) colours, and this app's own others.
export const THEMES: Readonly<Record<string, Theme>> = {
  Office: { name: "Office", colors: { dark1: "#000000", light1: "#ffffff", dark2: "#44546a", light2: "#e7e6e6", accent1: "#4472c4", accent2: "#ed7d31", accent3: "#a5a5a5", accent4: "#ffc000", accent5: "#5b9bd5", accent6: "#70ad47", hyperlink: "#0563c1" },
    fonts: { major: "Calibri Light", minor: "Calibri" } },
  Greyscale: { name: "Greyscale", colors: { dark1: "#000000", light1: "#ffffff", dark2: "#000000", light2: "#f8f8f8", accent1: "#dddddd", accent2: "#b2b2b2", accent3: "#969696", accent4: "#808080", accent5: "#5f5f5f", accent6: "#4d4d4d", hyperlink: "#5f5f5f" },
    fonts: { major: "Arial", minor: "Arial" } },
  Harbour: { name: "Harbour", colors: { dark1: "#1b1b1b", light1: "#ffffff", dark2: "#17406d", light2: "#dbefff", accent1: "#0f6fc6", accent2: "#009dd9", accent3: "#0bd0d9", accent4: "#10cf9b", accent5: "#7cca62", accent6: "#a5c249", hyperlink: "#f49100" },
    fonts: { major: "Georgia", minor: "Verdana" } },
  Ember: { name: "Ember", colors: { dark1: "#000000", light1: "#ffffff", dark2: "#4e3b30", light2: "#fbeec9", accent1: "#c0504d", accent2: "#d26b2a", accent3: "#e8a33a", accent4: "#9b8357", accent5: "#7a5c46", accent6: "#a17d6c", hyperlink: "#8e5843" },
    fonts: { major: "Georgia", minor: "Georgia" } },
};

let current: Theme = THEMES.Office;
export const theme = (): Theme => current;
export function setTheme(t: Theme): void { current = t; }

// A colour as the grid draws it: a role resolved by the theme; anything
// else as it is.
export function resolved(c: string): string {
  const k = c.trim().toLowerCase();
  return (ROLES as readonly string[]).includes(k) ? current.colors[k as Role] : c;
}
export const isRole = (c: string): boolean => (ROLES as readonly string[]).includes(c.trim().toLowerCase());
