export function makeBackupName(originalName, date = new Date()) {
  const stem = originalName.replace(/\.xlsx$/i, "");
  const stamp = [date.getFullYear(), date.getMonth() + 1, date.getDate()]
    .map((value) => String(value).padStart(2, "0"))
    .join("");
  return `${stem}_${stamp}_覆寫前備份.xlsx`;
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
