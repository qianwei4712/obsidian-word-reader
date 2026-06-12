export type ReaderStatusKind = "idle" | "loading" | "error";

export interface ReaderStatus {
  message: string;
  kind: ReaderStatusKind;
}

export class ReaderStatusController {
  private current: ReaderStatus = {
    message: "",
    kind: "idle",
  };

  constructor(
    private readonly onChange: (status: ReaderStatus) => void,
  ) {}

  set(message: string, kind: ReaderStatusKind = "idle"): void {
    this.current = { message, kind };
    this.onChange(this.current);
  }

  refresh(): void {
    this.onChange(this.current);
  }

  get value(): ReaderStatus {
    return this.current;
  }
}
