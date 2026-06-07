import { App, FileSystemAdapter, Notice, Platform, TFile } from "obsidian";
import type { WordReaderText } from "../i18n";

interface ElectronShell {
  openPath(path: string): Promise<string>;
}

export async function openExternalDocx(
  app: App,
  file: TFile,
  text?: WordReaderText,
): Promise<void> {
  if (!Platform.isDesktopApp) {
    new Notice(
      text?.notices.externalDesktopOnly ??
        "External opening is only available in Obsidian Desktop",
    );
    return;
  }

  const adapter = app.vault.adapter;
  if (!(adapter instanceof FileSystemAdapter)) {
    new Notice(
      text?.notices.externalLocalVaultOnly ??
        "External opening requires a local desktop vault",
    );
    return;
  }

  const shell = await getElectronShell();
  const failureReason = await shell.openPath(adapter.getFullPath(file.path));
  if (failureReason) {
    throw new Error(failureReason);
  }
}

async function getElectronShell(): Promise<ElectronShell> {
  if (!Platform.isDesktopApp) {
    throw new Error("External opening is only available on desktop");
  }

  const electron = await import("electron");
  return electron.shell;
}
