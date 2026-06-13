import { Platform } from "obsidian";

export interface ElectronNativeImage {
  isEmpty?(): boolean;
}

export interface ElectronNativeImageModule {
  createFromBuffer(buffer: Uint8Array): ElectronNativeImage;
  createFromDataURL(dataUrl: string): ElectronNativeImage;
}

export interface ElectronClipboard {
  writeImage(image: ElectronNativeImage): void;
}

export interface ElectronDialog {
  showSaveDialog(options: {
    title: string;
    defaultPath: string;
    filters: Array<{
      name: string;
      extensions: string[];
    }>;
  }): Promise<{
    canceled: boolean;
    filePath?: string;
  }>;
}

export interface ElectronShell {
  openPath(path: string): Promise<string>;
}

interface ElectronModule {
  clipboard: ElectronClipboard;
  dialog: ElectronDialog;
  nativeImage: ElectronNativeImageModule;
  shell: ElectronShell;
}

interface FileSystemPromises {
  writeFile(path: string, data: Uint8Array): Promise<void>;
}

export function getElectronModule(): ElectronModule {
  assertDesktop();
  return module.require("electron") as ElectronModule;
}

export function getFileSystemPromises(): FileSystemPromises {
  assertDesktop();
  return module.require("node:fs/promises") as FileSystemPromises;
}

function assertDesktop(): void {
  if (!Platform.isDesktopApp) {
    throw new Error("This operation is only available on desktop");
  }
}
