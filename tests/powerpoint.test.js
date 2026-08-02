import test from "node:test";
import assert from "node:assert/strict";
import { buildThemeColorRows, prepareOfficeImageData } from "../src/powerpoint.js";

test("builds a PowerPoint-style theme palette", () => {
  const rows = buildThemeColorRows([
    "#FFFFFF", "#000000", "#EEEEEE", "#222222", "#FF0000",
    "#00FF00", "#0000FF", "#FFFF00", "#00FFFF", "#FF00FF"
  ]);
  assert.equal(rows.length, 6);
  assert.ok(rows.every((row) => row.length === 10));
  assert.equal(rows[0][4], "#FF0000");
  assert.match(rows[5][9], /^#[0-9A-F]{6}$/);
});

test("passes raw SVG XML to Office XmlSvg insertion", async () => {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z"/></svg>';
  assert.equal(await prepareOfficeImageData("svg", svg, 256), svg);
});