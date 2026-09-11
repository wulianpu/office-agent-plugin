/**
 * Optional host adapter contracts (§86–§88): PowerPoint/WPS provide native
 * rendering and certification through their COM automation surface. Hosts
 * stay OPTIONAL — the plugin never depends on them for its main paths.
 */

export type HostAdapterId = "powerpoint" | "wps";

export interface HostRenderOptions {
  width?: number;
  height?: number;
}

export interface HostRenderResult {
  /** PNG bytes per requested slide index (1-based), in request order. */
  pngs: Buffer[];
  slideCount: number;
}

/**
 * §88 Certification Copy: the host opens a DISPOSABLE copy of the candidate.
 * If the host repairs/normalizes the bytes, certification FAILS — the
 * candidate itself is never touched.
 */
export interface HostCertificationResult {
  status: "pass" | "fail" | "unavailable";
  /** Host rewrote the bytes on open/save (repair/normalize) — §88 FAIL. */
  repaired: boolean;
  /** First slide rendered by the host as extra evidence, when available. */
  renderedPng?: Buffer;
  note?: string;
}

export interface HostAdapter {
  readonly id: HostAdapterId;
  readonly engine: string;
  /** Cheap liveness probe (COM instantiation only). */
  probe(): Promise<boolean>;
  /** §86 host render: export slides as PNGs (headless, no window). */
  renderSlides(path: string, slideIndices: number[], options?: HostRenderOptions): Promise<HostRenderResult>;
  /** §88 certification copy workflow. */
  certifyCopy(path: string, options?: HostRenderOptions): Promise<HostCertificationResult>;
}
