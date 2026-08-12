export function makeBackupName(originalName, period = "") {
  const stem = originalName.replace(/\.xlsx$/i, "");
  const now = new Date();
  const month = /^\d{6}$/.test(period)
    ? period
    : `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, "0")}`;
  return `${stem}_${month}_覆寫前備份.xlsx`;
}

export async function snapshotFile(file) {
  return new Blob([await file.arrayBuffer()], { type: file.type });
}

async function blobsMatch(left, right) {
  const [leftBytes, rightBytes] = await Promise.all([left.arrayBuffer(), right.arrayBuffer()]);
  if (leftBytes.byteLength !== rightBytes.byteLength) return false;
  const leftView = new Uint8Array(leftBytes);
  const rightView = new Uint8Array(rightBytes);
  return leftView.every((byte, index) => byte === rightView[index]);
}

export async function writeFileHandle(handle, blob) {
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (error) {
    await writable.abort?.().catch(() => {});
    throw error;
  }

  const writtenFile = await handle.getFile();
  if (!(await blobsMatch(writtenFile, blob))) {
    const error = new Error("檔案寫入後的內容與預期不一致。");
    error.code = "WRITE_VERIFICATION_FAILED";
    throw error;
  }
  return writtenFile;
}
