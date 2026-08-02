import test from "node:test";
import assert from "node:assert/strict";
import { prepareOfficeImageData } from "../src/powerpoint.js";

test("passes raw SVG XML to Office XmlSvg insertion", async () => {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M0 0h24v24H0z"/></svg>';
  assert.equal(await prepareOfficeImageData("svg", svg, 256), svg);
});