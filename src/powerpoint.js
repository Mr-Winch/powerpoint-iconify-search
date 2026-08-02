import { svgToBase64, svgToPngBase64 } from "./iconify.js";

const FALLBACK_THEME = {
  bodyBackgroundColor: "#f5f5f5",
  bodyForegroundColor: "#242424",
  controlBackgroundColor: "#ffffff",
  controlForegroundColor: "#242424"
};

function rgb(hex) {
  const value = String(hex || "").replace("#", "");
  if (!/^[0-9a-f]{6}$/i.test(value)) return null;
  return [0, 2, 4].map((index) => Number.parseInt(value.slice(index, index + 2), 16));
}

function hex(values) {
  return "#" + values.map((value) => Math.round(value).toString(16).padStart(2, "0")).join("");
}

function blend(first, second, ratio) {
  const a = rgb(first);
  const b = rgb(second);
  if (!a || !b) return first;
  return hex(a.map((value, index) => value * (1 - ratio) + b[index] * ratio));
}

function luminance(color) {
  const values = rgb(color) || [255, 255, 255];
  return values.reduce((sum, value, index) => sum + value * [0.2126, 0.7152, 0.0722][index], 0) / 255;
}

export function readOfficeTheme() {
  const previewTheme = new URLSearchParams(location.search).get("theme");
  const localPreview = ["localhost", "127.0.0.1"].includes(location.hostname);
  if (localPreview && previewTheme === "dark") {
    return {
      bodyBackgroundColor: "#292929",
      bodyForegroundColor: "#f5f5f5",
      controlBackgroundColor: "#333333",
      controlForegroundColor: "#f5f5f5",
      isDarkTheme: true
    };
  }
  const supplied = globalThis.Office?.context?.officeTheme;
  return { ...FALLBACK_THEME, ...(supplied || {}) };
}

export function applyOfficeTheme() {
  const theme = readOfficeTheme();
  const dark = typeof theme.isDarkTheme === "boolean"
    ? theme.isDarkTheme
    : luminance(theme.bodyBackgroundColor) < 0.45;
  const accent = theme.fluentThemeData?.colorBrandBackground || "#D24726";
  const background = theme.bodyBackgroundColor || FALLBACK_THEME.bodyBackgroundColor;
  const foreground = theme.bodyForegroundColor || FALLBACK_THEME.bodyForegroundColor;
  const controlBackground = theme.controlBackgroundColor || FALLBACK_THEME.controlBackgroundColor;
  const controlForeground = theme.controlForegroundColor || FALLBACK_THEME.controlForegroundColor;
  const contrastBase = dark ? "#ffffff" : "#000000";
  const root = document.documentElement;
  const variables = {
    "--office-bg": background,
    "--office-fg": foreground,
    "--office-control-bg": controlBackground,
    "--office-control-fg": controlForeground,
    "--office-muted": blend(foreground, background, 0.38),
    "--office-border": blend(background, contrastBase, dark ? 0.18 : 0.15),
    "--office-border-strong": blend(background, contrastBase, dark ? 0.32 : 0.30),
    "--office-hover": blend(background, contrastBase, dark ? 0.10 : 0.06),
    "--office-selected": blend(background, accent, dark ? 0.27 : 0.16),
    "--office-accent": accent,
    "--office-accent-hover": blend(accent, dark ? "#ffffff" : "#000000", 0.16),
    "--office-accent-text": luminance(accent) > 0.65 ? "#1b1b1b" : "#ffffff",
    "--office-shadow": dark ? "rgba(0,0,0,.34)" : "rgba(0,0,0,.12)"
  };
  for (const [name, value] of Object.entries(variables)) root.style.setProperty(name, value);
  root.dataset.theme = dark ? "dark" : "light";
  return JSON.stringify(theme);
}

export function watchOfficeTheme(onChange = () => {}) {
  let signature = applyOfficeTheme();
  const refresh = () => {
    const next = applyOfficeTheme();
    if (next !== signature) {
      signature = next;
      onChange();
    }
  };
  window.addEventListener("focus", refresh);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) refresh();
  });
  const timer = window.setInterval(refresh, 1500);
  return () => window.clearInterval(timer);
}

function hasPowerPointHost() {
  return Boolean(globalThis.Office?.context?.document?.setSelectedDataAsync);
}

export function supportsSvgInsertion() {
  return Boolean(globalThis.Office?.context?.requirements?.isSetSupported?.("ImageCoercion", "1.2"));
}

function setSelectedData(base64, coercionType, options) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64, {
      coercionType,
      imageLeft: options.left,
      imageTop: options.top,
      imageWidth: options.size,
      imageHeight: options.size
    }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "PowerPoint could not insert the image."));
    });
  });
}

export async function insertIconsIntoPowerPoint(items, options, progress = () => {}) {
  if (!items.length) throw new Error("Select at least one icon.");
  if (!hasPowerPointHost()) {
    await new Promise((resolve) => setTimeout(resolve, 350));
    return { inserted: items.length, simulated: true, format: options.format };
  }

  let format = options.format;
  let fallbackToPng = false;
  if (format === "svg" && !supportsSvgInsertion()) {
    format = "png";
    fallbackToPng = true;
  }

  const size = Math.min(720, Math.max(12, Number(options.size) || 72));
  const columns = Math.min(4, Math.ceil(Math.sqrt(items.length)));
  const gap = Math.max(9, Math.round(size * 0.16));
  const startLeft = 30;
  const startTop = 30;

  for (let index = 0; index < items.length; index += 1) {
    const item = items[index];
    progress(index + 1, items.length, item);
    const base64 = format === "svg"
      ? svgToBase64(item.svg)
      : await svgToPngBase64(item.svg, options.pngResolution);
    await setSelectedData(base64, format === "svg" ? Office.CoercionType.XmlSvg : Office.CoercionType.Image, {
      left: startLeft + (index % columns) * (size + gap),
      top: startTop + Math.floor(index / columns) * (size + gap),
      size
    });
  }
  return { inserted: items.length, simulated: false, format, fallbackToPng };
}
