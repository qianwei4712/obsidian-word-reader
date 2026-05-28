import { App, Notice, TFile, normalizePath } from "obsidian";

export async function createNoteFromDocx(app: App, file: TFile): Promise<void> {
  const notePath = normalizePath(
    file.path.replace(/\.docx$/i, ".md"),
  );

  const existing = app.vault.getAbstractFileByPath(notePath);
  if (existing instanceof TFile) {
    await app.workspace.getLeaf(true).openFile(existing);
    new Notice("Opened existing summary note");
    return;
  }

  const content = buildSummaryNote(file);
  const created = await app.vault.create(notePath, content);
  await app.workspace.getLeaf(true).openFile(created);
  new Notice("Created summary note");
}

function buildSummaryNote(file: TFile): string {
  const title = file.basename;
  const sourcePath = file.path;

  return [
    "---",
    `source: "${escapeYamlString(sourcePath)}"`,
    "type: word-note",
    `created: ${formatLocalDate(new Date())}`,
    "---",
    "",
    `# ${title}`,
    "",
    `原文：[[${sourcePath}]]`,
    "",
    "## 摘要",
    "",
    "## 关键结论",
    "",
    "## 待处理",
    "",
    "## 引用摘录",
    "",
  ].join("\n");
}

function formatLocalDate(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function escapeYamlString(value: string): string {
  return value.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}
