export class ReaderLifecycle {
  private generation = 0;

  begin(): number {
    this.generation += 1;
    return this.generation;
  }

  cancel(): void {
    this.generation += 1;
  }

  isCurrent(token: number): boolean {
    return token === this.generation;
  }

  get currentToken(): number {
    return this.generation;
  }
}
