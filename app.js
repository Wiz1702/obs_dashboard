import define from "./index.js";
import {Runtime, Inspector} from "./runtime.js";

const runtime = new Runtime();
const notebook = document.querySelector("#notebook");
const themeToggle = document.querySelector("#theme-toggle");
const themeLabel = document.querySelector("#theme-label");
const seeAnalytics = document.querySelector("#see-analytics");

function setTheme(theme) {
  document.body.dataset.theme = theme;
  localStorage.setItem("dashboard-theme", theme);

  const isLight = theme === "light";
  themeToggle.setAttribute("aria-pressed", String(isLight));
  themeLabel.textContent = isLight ? "Dark" : "Light";
}

const savedTheme = localStorage.getItem("dashboard-theme");
const prefersLight = window.matchMedia("(prefers-color-scheme: light)").matches;
setTheme(savedTheme || (prefersLight ? "light" : "dark"));

themeToggle.addEventListener("click", () => {
  setTheme(document.body.dataset.theme === "light" ? "dark" : "light");
});

seeAnalytics?.addEventListener("click", () => {
  window.location.href = "./analytics.html";
});

function observer(name) {
  if (name && !name.startsWith("viewof ")) return null;

  const cell = document.createElement("div");
  cell.dataset.name = name || "";
  notebook?.appendChild(cell);
  return new Inspector(cell);
}

if (notebook) {
  runtime.module(define, observer);
}
