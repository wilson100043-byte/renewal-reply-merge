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

export async function writeFileHandle(handle, blob) {
  const writable = await handle.createWritable();
  try {
    await writable.write(blob);
    await writable.close();
  } catch (error) {
    await writable.abort?.().catch(() => {});
    throw error;
  }
}
