import assert from "node:assert/strict";
import test from "node:test";

import { makeBackupName, writeFileHandle } from "../file-output.js";

test("names the backup as the pre-overwrite copy", () => {
  assert.equal(
    makeBackupName("未績訂_知文.xlsx", new Date(2026, 7, 11)),
    "未績訂_知文_20260811_覆寫前備份.xlsx",
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
