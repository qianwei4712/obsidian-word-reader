export class RetainedResourceRegistry {
  private activeResources = new Set<string>();
  private retainedResources = new Map<string, number>();
  private pendingReleases = new Set<string>();

  constructor(
    private readonly releaseResource: (resource: string) => void,
  ) {}

  replace(resources: Iterable<string>): void {
    this.releaseActive();
    this.activeResources = new Set(resources);
    for (const resource of this.activeResources) {
      this.pendingReleases.delete(resource);
    }
  }

  retain(resource: string): () => void {
    this.retainedResources.set(
      resource,
      (this.retainedResources.get(resource) ?? 0) + 1,
    );
    let released = false;

    return () => {
      if (released) {
        return;
      }
      released = true;

      const remaining = (this.retainedResources.get(resource) ?? 1) - 1;
      if (remaining > 0) {
        this.retainedResources.set(resource, remaining);
        return;
      }

      this.retainedResources.delete(resource);
      if (this.pendingReleases.delete(resource)) {
        this.releaseResource(resource);
      }
    };
  }

  releaseActive(): void {
    for (const resource of this.activeResources) {
      if ((this.retainedResources.get(resource) ?? 0) > 0) {
        this.pendingReleases.add(resource);
      } else {
        this.releaseResource(resource);
      }
    }
    this.activeResources.clear();
  }
}

export function releaseResources(
  resources: Iterable<string>,
  releaseResource: (resource: string) => void,
): void {
  for (const resource of resources) {
    releaseResource(resource);
  }
}
