import assert from "node:assert/strict";
import test from "node:test";

import { makeBackupName, snapshotFile, writeFileHandle } from "../file-output.js";

test("names the backup as the pre-overwrite copy", () => {
  assert.equal(
    makeBackupName("未績訂_知文.xlsx", "202608"),
    "未績訂_知文_202608_覆寫前備份.xlsx",
  );
});

test("writes and closes the selected file handle", async () => {
  const events = [];
  const handle = {
    createWritable: async () => ({
      write: async (blob) => events.push(["write", await blob.text()]),
      close: async () => events.push(["close"]),
    }),
  };

  await writeFileHandle(handle, new Blob(["updated"]));
  assert.deepEqual(events, [["write", "updated"], ["close"]]);
});

test("snapshots backup bytes before the source file changes", async () => {
  let contents = "original";
  const source = {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    arrayBuffer: async () => new TextEncoder().encode(contents).buffer,
  };

  const snapshot = await snapshotFile(source);
  contents = "overwritten";

  assert.equal(await snapshot.text(), "original");
});
