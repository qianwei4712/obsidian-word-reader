import {
  App,
  Notice,
  TFile,
  normalizePath,
} from "obsidian";

export interface ReaderSummaryNotices {
  openedExisting: string;
  created: string;
}

export async function createOrOpenReaderSummaryNote(
  app: App,
  file: TFile,
  sourceExtension: string,
  content: string,
  notices: ReaderSummaryNotices,
): Promise<void> {
  const escapedExtension = sourceExtension.replace(
    /[.*+?^${}()|[\]\\]/g,
    "\\$&",
  );
  const notePath = normalizePath(
    file.path.replace(new RegExp(`${escapedExtension}$`, "i"), ".md"),
  );
  const existing = app.vault.getAbstractFileByPath(notePath);
  if (existing instanceof TFile) {
    await app.workspace.getLeaf(true).openFile(existing);
    new Notice(notices.openedExisting);
    return;
  }

  const created = await app.vault.create(notePath, content);
  await app.workspace.getLeaf(true).openFile(created);
  new Notice(notices.created);
}
