import { Plugin } from "obsidian";

import { openExternalDocx } from "./commands/openExternal";
import { createNoteFromDocx } from "./commands/createNoteFromDocx";
import { WordView, VIEW_TYPE_WORD_READER } from "./WordView";

export default class WordReaderPlugin extends Plugin {
  async onload(): Promise<void> {
    this.registerView(
      VIEW_TYPE_WORD_READER,
      (leaf) => new WordView(leaf, this),
    );

    this.registerExtensions(["docx"], VIEW_TYPE_WORD_READER);

    this.addCommand({
      id: "word-reader-reload",
      name: "Reload current Word document",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void view.reload();
        }

        return true;
      },
    });

    this.addCommand({
      id: "word-reader-copy-text",
      name: "Copy text from current Word document",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void view.copyText();
        }

        return true;
      },
    });

    this.addCommand({
      id: "word-reader-create-note",
      name: "Create summary note for current Word document",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void createNoteFromDocx(this.app, view.file);
        }

        return true;
      },
    });

    this.addCommand({
      id: "word-reader-open-external",
      name: "Open current Word document externally",
      checkCallback: (checking) => {
        const view = this.getActiveWordView();
        if (!view?.file) {
          return false;
        }

        if (!checking) {
          void openExternalDocx(this.app, view.file);
        }

        return true;
      },
    });
  }

  private getActiveWordView(): WordView | null {
    return this.app.workspace.getActiveViewOfType(WordView);
  }
}
