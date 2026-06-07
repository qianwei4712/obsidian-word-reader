declare module "electron" {
  export const clipboard: {
    writeImage(image: {
      isEmpty?(): boolean;
    }): void;
  };

  export const dialog: {
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
  };

  export const nativeImage: {
    createFromBuffer(buffer: Uint8Array): {
      isEmpty?(): boolean;
    };
    createFromDataURL(dataUrl: string): {
      isEmpty?(): boolean;
    };
  };

  export const shell: {
    openPath(path: string): Promise<string>;
  };
}
