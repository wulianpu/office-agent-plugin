/**
 * ArtifactStore (§15): sole owner of the ref → physical path mapping.
 * Everything else in the runtime addresses Office bytes via ArtifactRef.
 * Staging artifacts (candidates, temp copies) live inside the workspace root.
 */

import { mkdir, readdir, rm } from "node:fs/promises";
import { join, resolve } from "node:path";
import {
  newArtifactRef
} from "../../support/ids.js";
import { canonicalSourceKey, cloneFile, ensureDir, pathExists, removeQuiet } from "../../support/fsx.js";
import type { ArtifactRef, OfficeFormat } from "../../contracts/ids.js";
import { formatFromPath } from "../../contracts/ids.js";
import { OfficeRuntimeError } from "../../contracts/document.js";
import type { ArtifactStorePersistence } from "./store-persistence.js";

export interface RegisterOptions {
  format?: OfficeFormat;
}

export class ArtifactStore {
  private readonly byRef = new Map<ArtifactRef, StoredArtifact>();
  private readonly byPath = new Map<string, ArtifactRef>();
  private stagingSeq = 0;

  constructor(
    readonly workspaceRoot: string,
    private readonly persistence?: ArtifactStorePersistence
  ) {}

  static async open(workspaceRoot: string, persistence?: ArtifactStorePersistence): Promise<ArtifactStore> {
    const store = new ArtifactStore(workspaceRoot, persistence);
    await store.hydrate();
    return store;
  }

  /**
   * P0-1: hydrate from persistence — persisted byPath keys go through
   * canonicalSourceKey so restart lookups hit the same keys live code uses.
   */
  async hydrate(): Promise<void> {
    await ensureDir(this.stagingRoot);
    if (!this.persistence) return;
    for (const row of this.persistence.loadArtifacts()) {
      if (!this.byRef.has(row.ref)) {
        await this.bindArtifact(row);
      }
    }
  }

  get stagingRoot(): string {
    return join(this.workspaceRoot, "staging");
  }

  /** Register a source file (user file anywhere on disk) and get an opaque ref. */
  async register(path: string, options: RegisterOptions = {}): Promise<ArtifactRef> {
    const abs = resolve(path);
    const format = options.format ?? formatFromPath(abs);
    if (!format) {
      throw new OfficeRuntimeError("unsupported-format", `not an Office file: ${abs}`);
    }
    if (!(await pathExists(abs))) {
      throw new OfficeRuntimeError("artifact-missing", `artifact source missing: ${abs}`);
    }
    const existing = this.byPath.get(canonicalSourceKey(abs));
    if (existing) return existing;

    const ref = newArtifactRef();
    await this.bindArtifact({ ref, path: abs, kind: "source", format });
    return ref;
  }

  /** Internal-only path resolution. Never surfaced through MCP tools (INV-12). */
  resolvePath(ref: ArtifactRef): string {
    const artifact = this.byRef.get(ref);
    if (!artifact) {
      throw new OfficeRuntimeError("artifact-missing", `unknown artifact ref: ${ref}`);
    }
    return artifact.path;
  }

  formatOf(ref: ArtifactRef): OfficeFormat {
    const artifact = this.byRef.get(ref);
    if (!artifact) {
      throw new OfficeRuntimeError("artifact-missing", `unknown artifact ref: ${ref}`);
    }
    return artifact.format;
  }

  tryResolvePath(ref: ArtifactRef): string | undefined {
    return this.byRef.get(ref)?.path;
  }

  /** Reverse lookup used by crash recovery (journal records physical paths). */
  tryResolveRefByPath(path: string): ArtifactRef | undefined {
    return this.byPath.get(canonicalSourceKey(path));
  }

  /**
   * Candidate clone (§66): copy-on-write semantics via full copy fallback —
   * hard links are never used for isolation.
   */
  async createStagingCopy(sourceRef: ArtifactRef): Promise<ArtifactRef> {
    const sourcePath = this.resolvePath(sourceRef);
    const stagingRef = newArtifactRef();
    const stagingPath = this.stagingPathFor(stagingRef, this.formatOf(sourceRef));
    await ensureDir(this.stagingRoot);
    await cloneFile(sourcePath, stagingPath);
    await this.bindArtifact({ ref: stagingRef, path: stagingPath, kind: "staging", format: this.formatOf(sourceRef) });
    return stagingRef;
  }

  /**
   * P1-high: SINGLE registration point — every source/staging binding goes
   * through byRef + byPath (canonical) + persistence together.
   */
  private async bindArtifact(artifact: StoredArtifact): Promise<void> {
    this.byRef.set(artifact.ref, artifact);
    this.byPath.set(canonicalSourceKey(artifact.path), artifact.ref);
    await this.persistence?.saveArtifact(artifact);
  }

  /** Register a staging file that already exists (e.g. produced by OfficeCLI create). */
  async registerStagingFile(path: string, format: OfficeFormat): Promise<ArtifactRef> {
    const abs = resolve(path);
    const existing = this.byPath.get(canonicalSourceKey(abs));
    if (existing) return existing;
    const ref = newArtifactRef();
    await this.bindArtifact({ ref, path: abs, kind: "staging", format });
    return ref;
  }

  stagingPathFor(ref: ArtifactRef, format: OfficeFormat): string {
    return join(this.stagingRoot, `${ref}.${format}`);
  }

  /**
   * Last-resort lock releaser for staging deletion (engine daemons can hold
   * Windows locks past the rm retries). Wired to officecli close by the
   * runtime service.
   */
  setLockReleaser(releaser: (path: string) => Promise<void>): void {
    this.lockReleaser = releaser;
  }

  private lockReleaser?: (path: string) => Promise<void>;

  async release(ref: ArtifactRef): Promise<void> {
    const artifact = this.byRef.get(ref);
    if (!artifact) return;
    if (artifact.kind === "staging") {
      try {
        await removeQuiet(artifact.path);
      } catch (error) {
        if ((error as NodeJS.ErrnoException).code !== "EBUSY" && (error as NodeJS.ErrnoException).code !== "EPERM") {
          throw error;
        }
        await this.lockReleaser?.(artifact.path);
        await removeQuiet(artifact.path);
      }
    }
    this.byRef.delete(ref);
    this.byPath.delete(canonicalSourceKey(artifact.path));
    await this.persistence?.deleteArtifact(ref);
  }

  /** Best-effort cleanup of orphaned staging files (recovery pass calls this). */
  async purgeStaging(keep: Set<ArtifactRef>): Promise<number> {
    let removed = 0;
    await mkdir(this.stagingRoot, { recursive: true });
    for (const name of await readdir(this.stagingRoot)) {
      const ref = this.byPath.get(canonicalSourceKey(join(this.stagingRoot, name)));
      if (ref && keep.has(ref)) continue;
      await removeQuiet(join(this.stagingRoot, name));
      if (ref) {
        this.byRef.delete(ref);
        this.byPath.delete(join(this.stagingRoot, name));
        await this.persistence?.deleteArtifact(ref);
      }
      removed++;
    }
    return removed;
  }

  async dispose(): Promise<void> {
    this.byRef.clear();
    this.byPath.clear();
  }
}

export interface StoredArtifact {
  ref: ArtifactRef;
  path: string;
  kind: "source" | "staging";
  format: OfficeFormat;
}
