import { App, FileSystemAdapter, Notice, Platform, TFile } from "obsidian";

interface ElectronShell {
  openPath(path: string): Promise<string>;
}

export async function openExternalDocx(app: App, file: TFile): Promise<void> {
  if (!Platform.isDesktopApp) {
    new Notice("External opening is only available in Obsidian Desktop");
    return;
  }

  const adapter = app.vault.adapter;
  if (!(adapter instanceof FileSystemAdapter)) {
    new Notice("External opening requires a local desktop vault");
    return;
  }

  const shell = getElectronShell();
  const failureReason = await shell.openPath(adapter.getFullPath(file.path));
  if (failureReason) {
    throw new Error(failureReason);
  }
}

function getElectronShell(): ElectronShell {
  const electron = require("electron") as { shell: ElectronShell };
  return electron.shell;
}
