// for package maintenance: document auto updater.
// usage: node docup.js

import { promisify } from 'util';
import child_process from 'child_process';
const exec = promisify(child_process.exec);

async function run() {
  await exec("yarn typedoc", {});
  await exec("git add -A", { cwd: "docs" });
  await exec("git commit -a -m \"Update doc\"", { cwd: "docs" });
  await exec("git push", { cwd: "docs" });
  await exec("git commit docs -m \"- Update doc\"", {});
  await exec("git push", {});
  console.info("OK!");
}

run();
