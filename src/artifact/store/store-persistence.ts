/**
 * Persistence seam for ArtifactStore — implemented by the SQLite layer.
 */

import type { ArtifactRef, OfficeFormat } from "../../contracts/ids.js";
import type { StoredArtifact } from "./artifact-store.js";

export interface ArtifactStorePersistence {
  saveArtifact(row: StoredArtifact): Promise<void>;
  deleteArtifact(ref: ArtifactRef): Promise<void>;
  loadArtifacts(): StoredArtifact[];
}
