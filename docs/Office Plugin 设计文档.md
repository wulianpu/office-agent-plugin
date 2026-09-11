# Harness Office Plugin
# 完整开发与架构评审设计

**版本：** v3.0  
**状态：** Architecture Freeze Candidate  
**日期：** 2026-09-10  
**目标宿主：** Harness Desktop / Electron  
**核心格式：** DOCX / XLSX / PPTX  
**核心实现：** GenOffice + OfficeCLI  
**运行原则：** Offline First / Read Optimized / Native Editor / Agent Safe / Recoverable

---

# 1. 文档定位

本文档定义 Harness Desktop 中 Office 能力的完整开发架构。

它覆盖：

- Office Plugin 边界；
- 高频 Office 文件预览；
- 正式打开与按需编辑；
- GenOffice 可视化编辑器集成；
- OfficeCLI Agent Runtime；
- MCP Tool Plane；
- Portable MCP View；
- Document Session；
- Artifact Runtime；
- Revision / Candidate；
- Writer Lease / Fencing；
- 文件 Commit；
- Crash Recovery；
- Memory / Cache / Resource Governance；
- 性能基线；
- Verification；
- Offline Security；
- PowerPoint / WPS Optional Host Adapter；
- Testing / Benchmark；
- Release / Compatibility；
- Development Roadmap。

本文档用于：

```text
Architecture Review
Development Planning
Implementation Review
Code Review
Performance Review
Security Review
Release Review
```

---

# 2. 演进基础

本架构继承原 Harness Desktop PPT Agent v1.0 中已经验证合理的核心原则。

原设计目标是让 Harness Agent 可靠操作 PPTX，并以：

> OfficeCLI First，Native Fallback

作为核心方向，同时针对 WPS Presentation 与 Microsoft PowerPoint 建立宿主补充能力。

原架构明确拒绝建立新的：

```text
CanonicalPresentation
CanonicalSlide
CanonicalShape
CanonicalChart
```

以及第三套 PPT API，避免重复包装 OfficeCLI、WPS 和 PowerPoint 已存在的业务模型。这个原则继续保留。

原架构中的：

```text
Single Writer
Minimal Mutation
Preserve by Default
Target Verification
Recoverable Editing
Document Content Is Data
```

也继续作为新版的重要安全基础。

原设计的 Working Copy、Shadow Copy、Package Diff、Resident OfficeCLI、Verification Pipeline 等能力则被进一步泛化为：

```text
CandidateRevision
Artifact Runtime
WriterLease
VerificationReport
Commit Protocol
```

---

# 3. 产品使用模型

本项目最主要的用户行为不是：

> 长时间编辑 Office。

而是：

> **高频打开、查看、预览 Office 文件，偶尔人工编辑，偶尔要求 Agent 修改。**

因此以下操作频率预计从高到低：

```text
Preview / Peek
      ↓
Open / Read
      ↓
Selection / Copy / Search
      ↓
Human Edit
      ↓
Agent Edit
```

架构必须优先优化：

> Read Path。

而不是让所有读取都承担完整 Write Runtime 的成本。

---

# 4. 核心用户体验

用户在 Harness 中：

```text
Files
├── report.docx
├── data.xlsx
├── board.pptx
└── appendix.docx
```

用户单击：

```text
board.pptx
```

应快速出现 Preview。

用户双击：

```text
board.pptx
```

正式进入可滚动、翻页、缩放、选择的 Open 状态。

用户点击文本并输入：

```text
2026 Revenue
```

系统才原地升级为 Edit。

用户告诉 Agent：

> 帮我重新排版当前这一页。

系统不要求 Human Editor 获得 writer，而是：

```text
Agent
 ↓
OfficeCLI
 ↓
Candidate
 ↓
Verification
 ↓
Proposal
```

用户可以：

```text
Accept
Reject
Edit Proposal
```

所有操作应发生在同一个 Harness 工作台中。

---

# 5. 顶层架构原则

本项目冻结以下设计原则。

## P1 — One Logical Office Plugin

Harness 对外只认识：

```text
OfficePlugin
```

而不是：

```text
WordPlugin
ExcelPlugin
PowerPointPlugin
OfficeCLIPlugin
GenOfficePlugin
```

---

## P2 — Internal Modular Capabilities

一个逻辑 Plugin 内部保持清晰模块：

```text
Preview
Editor
Agent
Artifact
Document
Verification
Host
MCP
```

---

## P3 — Read Optimized

Preview/Open 是最高频操作。

读取 Office 不应该承担：

```text
Candidate
WriterLease
OfficeCLI
Undo
Commit
```

成本。

---

## P4 — MVCC-like Read Semantics

Reader pin immutable artifact。

Writer 产生新的 artifact。

```text
Committed 42
 ├─ Reader A
 ├─ Reader B
 └─ Candidate 43
```

Reader 不因为 Writer 而阻塞。

---

## P5 — GenOffice First for Human

GenOffice：

```text
Human Visual Editor
```

负责：

```text
Preview
View
Edit
Selection
Undo
Human Save
```

---

## P6 — OfficeCLI First for Agent

Agent 对 DOCX/XLSX/PPTX 的：

```text
inspect
query
edit
render
verify
```

优先通过 OfficeCLI。

---

## P7 — Native Editor First

Harness 自己运行时：

```text
GenOffice Editor
```

直接作为 First-party Native Plugin 挂入 Harness Renderer。

主编辑器不强制经过 MCP iframe。

---

## P8 — MCP Control Plane

MCP 用于：

```text
Agent tools
portable review UI
cross-host capability
```

不用于大型 Office Editor 的高频数据热路径。

当前 MCP Apps 实现使用 2026-01-26 协议版本并提供 inline/fullscreen/pip 等 UI 模式。

---

## P9 — Native Data Plane

高频或大型数据：

```text
XLSX ranges
images
binary previews
```

走：

```text
MessagePort
Transferable ArrayBuffer
Native Sidecar
```

Electron 的 `utilityProcess` 支持 Node 子进程并支持转移 MessagePort；`MessagePortMain` 可以建立长期双向 Channel。

---

## P10 — No Canonical Office IR

禁止建立：

```text
UniversalDocument
UniversalShape
UniversalChart
UniversalAnimation
```

统一的只是：

```text
Session
Artifact
Revision
Capability
Operation Envelope
```

不是 Office 内容模型。

---

## P11 — Single Writer Lease

同一个 DocumentSession：

```text
最多一个 Writer
```

但可以同时存在任意 Reader。

---

## P12 — Candidate Before Commit

Agent：

```text
永远不直接写 Committed Artifact
```

---

## P13 — Explicit Flush

OfficeCLI resident 中 mutation 成功不代表磁盘已经更新。

OfficeCLI 当前明确规定 resident mutation 默认首先存在于内存，`save` 用于立即 flush 至磁盘但继续保持 resident；`close` 则 flush 并释放 resident。

因此：

```text
save
=
Read Visibility Barrier

close
=
Writer Handoff Barrier
```

---

## P14 — Preserve Aggressively

GenOffice：

```text
能编辑 → 编辑
能显示 → 显示
不能显示 → Preserve
```

不能因为 Renderer 不理解就主动删除。

GenOffice 当前本身也采用 narrow patch / untouched content preservation 的方向：DOCX 只重新生成 dirty blocks，其余 ZIP entries 尽量保持原字节；其 README 明确把相同哲学扩展到 Sheets/Slides。

---

## P15 — Offline Is Default

核心 Office 能力：

```text
0 network dependency
```

联网是可选能力，而不是基本要求。

---

# 6. 总体组件架构

```text
                    ┌──────────────────────┐
                    │       Harness        │
                    │                      │
                    │ Files / Tabs / Agent │
                    └──────────┬───────────┘
                               │
                         OfficePlugin
                               │
      ┌────────────────────────┼────────────────────────┐
      │                        │                        │
      ▼                        ▼                        ▼
Artifact Runtime          Editor Runtime          Agent Runtime
      │                        │                        │
      │                        │                        ▼
      │                    GenOffice                MCP Tools
      │                                             │
      │                                             ▼
      │                                         OfficeCLI
      │
      └──────────────────┬─────────────────────────────┘
                         │
                         ▼
                 OfficeRuntimeService
                         │
             ┌───────────┼───────────┐
             ▼           ▼           ▼
          Session     Revision    Candidate
             │           │           │
             └───────────┼───────────┘
                         ▼
                  Artifact Registry
                         │
                DOCX / XLSX / PPTX
```

Optional：

```text
PowerPoint Adapter
WPS Adapter
Portable MCP Review View
```

---

# 7. 运行进程模型

推荐默认：

```text
Process 1
Harness Main

Process 2
Harness Renderer
 ├─ Workbench
 ├─ Agent UI
 └─ GenOffice Native Editor

Process 3
OfficeRuntimeService
Electron utilityProcess

Process 4...
OfficeCLI processes

Optional:
XLSX Rust Sidecar

Optional:
PowerPoint / WPS
```

---

# 8. 为什么 OfficeRuntimeService 独立

Harness Main 只应管理：

```text
Window
Permission
Process lifecycle
Plugin lifecycle
```

OfficeRuntimeService 管：

```text
DocumentSession
ArtifactRegistry
Candidate
Revision
WriterLease
Verification
Commit
Recovery
ResourceGovernor
```

这样大型：

```text
hash
ZIP scan
Package Diff
verification
```

不会污染 Harness Main。

Electron `utilityProcess` 本身提供 Node 子进程能力和 MessagePort 通讯机制，非常适合该后台 Runtime。

---

# 9. Plugin 模块布局

```text
plugins/
└── office/
    │
    ├── contracts/
    │
    ├── runtime/
    │
    ├── artifact/
    │
    ├── preview/
    │
    ├── editors/
    │   ├── docs/
    │   ├── sheets/
    │   └── slides/
    │
    ├── agent/
    │   └── officecli/
    │
    ├── verification/
    │
    ├── mcp/
    │
    ├── hosts/
    │   ├── powerpoint/
    │   └── wps/
    │
    ├── native/
    │
    └── vendor/
        └── genoffice/
```

---

# 10. 读取生命周期

读取正式分成：

```text
Preview
Open
Edit
```

而不是 Preview/Edit 二档。

---

# 11. PreviewRequest

用户：

```text
单击
Hover
Quick Look
Agent attachment preview
```

创建：

```text
PreviewRequest
```

不创建：

```text
DocumentSession
WriterLease
Candidate
OfficeCLI
```

接口：

```ts
interface PreviewRequest {
  requestId: string;
  artifactRef: ArtifactRef;

  priority:
    | "visible"
    | "prefetch"
    | "background";

  scope?: PreviewScope;
}
```

---

# 12. 正式 Open

用户双击：

```text
Open
```

创建：

```text
DocumentSession
```

但默认：

```text
read-only
```

不获得 WriterLease。

---

# 13. Edit Activation

第一次 mutation：

```text
Typing
Paste
Delete
Format
Drag Shape
Edit Cell
```

触发：

```text
Edit Promotion
```

流程：

```text
Read-only DocumentSession
        ↓
check stable source identity
        ↓
check Candidate conflict
        ↓
acquire Human WriterLease
        ↓
activate editable capability
        ↓
replay triggering operation
```

目标是：

> 同一个 Tab 原地变为 Editor。

---

# 14. OpenIntent

正式冻结：

```ts
type OpenIntent =
  | "preview"
  | "open"
  | "edit";
```

其中：

```text
preview
→ PreviewRequest

open
→ DocumentSession + Read Runtime

edit
→ DocumentSession + Editor Activation
```

---

# 15. ArtifactRef

任何物理 Office bytes 不直接用 path 暴露。

```ts
type ArtifactRef = string;
```

真实文件：

```text
C:\Users\...\report.docx
```

只由 Artifact Store 知道。

---

# 16. ArtifactVersionKey

```ts
interface ArtifactVersionKey {
  artifactRef: ArtifactRef;

  fingerprint: {
    size: bigint;
    mtimeNs: bigint;
    fileId?: string;
  };

  contentHash?: string;
}
```

---

# 17. 两级 Identity

快速路径：

```text
FileFingerprint
```

用于：

```text
Preview
Cache lookup
```

强一致性路径：

```text
SHA-256
```

用于：

```text
Edit
Agent
Verification
Commit
```

禁止：

```text
Preview 2GB XLSX
→ 必须先同步 Hash 2GB
```

---

# 18. ReadConsistency

```ts
type ReadConsistency =
  | "optimistic"
  | "stable";
```

Preview 默认：

```text
optimistic
```

执行：

```text
fingerprint before
 ↓
render
 ↓
fingerprint after
```

不一致：

```text
discard/retry
```

Edit / Agent：

```text
stable
```

---

# 19. ArtifactContext

这是整个高频读取架构的核心。

```ts
interface ArtifactContext {
  artifactRef: ArtifactRef;

  version: ArtifactVersionKey;

  format:
    | "docx"
    | "xlsx"
    | "pptx";

  consistency: ReadConsistency;

  rendererVersion: string;

  lastAccessAt: number;
}
```

它必须：

```text
Immutable document identity
Shareable
Reference counted
Evictable
```

---

# 20. ArtifactContext 不是 Canonical IR

不同格式仍然有完全不同实现。

PPTX：

```text
PptxArtifactContext
```

XLSX：

```text
XlsxArtifactContext
```

DOCX：

```text
DocxArtifactContext
```

统一的只是生命周期。

---

# 21. ArtifactRegistry

所有 ArtifactContext 获取必须经过：

```text
ArtifactRegistry
```

禁止：

```text
PreviewService new Context
Editor new Context
Verification new Context
```

统一：

```ts
interface ArtifactRegistry {
  acquire(
    key: ArtifactVersionKey
  ): Promise<ArtifactLease>;

  release(lease: ArtifactLease): void;
}
```

---

# 22. ArtifactLease

```ts
interface ArtifactLease {
  leaseId: string;
  context: ArtifactContext;

  release(): void;
}
```

用于追踪：

```text
谁还在持有 Context？
```

Debug 时可以看到：

```text
Artifact A
├─ Editor tab-12
├─ Preview request-71
└─ Verification job-9
```

---

# 23. In-flight Deduplication

非常重要。

用户：

```text
Hover
 ↓
Preview

100ms
 ↓
Open

200ms
 ↓
Edit
```

不能产生三个 parse。

ArtifactRegistry 必须同时维护：

```text
Completed Context Cache
+
In-flight Build Registry
```

即：

```text
1 Build Job
3 Consumers
```

---

# 24. Consumer-aware Cancellation

如果 Preview 取消：

```text
consumer A release
```

但 Open 仍需要：

```text
consumer B
```

Build Job 不取消。

只有：

```text
consumer count = 0
```

才真正 cancel/deprioritize。

---

# 25. Progressive Context Enrichment

ArtifactContext 内容身份 immutable。

但 cache 可以 lazy enrichment：

```text
metadata
 ↓
first page
 ↓
current section
 ↓
more parsed parts
```

允许：

```text
未解析 → 已解析
```

禁止：

```text
内容 A → 内容 B
```

内容变化必须产生新的 ArtifactVersion。

---

# 26. FormatRuntime

每个格式一个进程级 Runtime：

```text
DocsRuntime
SheetsRuntime
SlidesRuntime
```

负责共享：

```text
font registry
worker pools
renderer infrastructure
format metadata
parser infrastructure
```

而不是每个文档重新初始化。

接口：

```ts
interface FormatRuntime {
  initialize(): Promise<void>;

  createArtifactContext(
    input: ArtifactBuildInput
  ): Promise<ArtifactContext>;

  createEditor(
    input: EditorBootstrapContext
  ): Promise<EditorInstance>;

  trimMemory(
    level: MemoryTrimLevel
  ): Promise<void>;
}
```

---

# 27. Preview → Open → Edit 复用

理想流程：

```text
ArtifactContext
      │
      ├─ Preview
      │
      ├─ Open
      │
      └─ Editor Bootstrap
```

不重复读取：

```text
ZIP index
relationships
styles
font metadata
media metadata
```

能复用多少由格式实现决定。

Contract 不要求强制全量复用。

---

# 28. PPTX Read Path

推荐：

```text
PPTX
 ↓
pptx-engine
 ↓
metadata/index
 ↓
first slide
 ↓
visible thumbnails
 ↓
interactive view
```

编辑时：

```text
reuse PptxArtifactContext
 ↓
activate editing interaction
```

---

# 29. DOCX Read Path

```text
DOCX
 ↓
docx-engine
 ↓
block/index
 ↓
visible page/block
 ↓
read view
```

编辑：

```text
reuse document/block model where possible
 ↓
activate Tiptap editable runtime
```

GenOffice 当前 DOCX 设计以 top-level OOXML block 为基本编辑单元，并将未修改内容尽量保持原字节。

---

# 30. XLSX Read Path

必须保留 GenOffice 当前 sidecar 思路。

当前实现中：

```text
Renderer
→ typed bridge
→ Electron IPC
→ Rust XLSX sidecar
→ row chunks
```

Sidecar 首先读取 ZIP/workbook metadata，而不是构建完整 JS workbook snapshot；首次 range 请求才流式解析 sheet XML，renderer 只保留 viewport + buffer。

因此：

```text
Workbook Size
≠
Renderer Working Set
```

这应成为整个项目的性能参考。

---

# 31. Working Set 原则

正式冻结：

> **内存规模应尽量与 Active Working Set 成比例，而不是与 Office 文件总大小成比例。**

例如：

```text
1,000,000 row XLSX
```

不意味着：

```text
1,000,000 JS row objects
```

---

# 32. DocumentSession

正式 Open 后创建。

```ts
interface DocumentSession {
  sessionId: SessionId;
  documentId: DocumentId;

  format: OfficeFormat;

  backend:
    | "managed-file"
    | "external-host";

  lifecycle: SessionLifecycle;

  sessionEpoch: number;

  committedRevision: CommittedRevision;

  candidate?: CandidateRevision;

  writerLease?: WriterLease;

  editor?: EditorBinding;
}
```

---

# 33. SessionLifecycle

不要把所有状态压进一个大枚举。

```ts
type SessionLifecycle =
  | "opening"
  | "ready"
  | "conflict"
  | "recovery-required"
  | "closing"
  | "closed";
```

---

# 34. EditorState

独立：

```ts
type EditorState =
  | "detached"
  | "loading"
  | "clean"
  | "dirty"
  | "saving"
  | "suspended"
  | "error";
```

---

# 35. CandidateState

独立：

```ts
type CandidateState =
  | "preparing"
  | "mutating"
  | "flushing"
  | "verifying"
  | "ready"
  | "human-amended"
  | "committing"
  | "failed";
```

避免状态组合爆炸。

---

# 36. CommittedRevision

原来的 acceptedRevision 正式改名：

```ts
interface CommittedRevision {
  revisionId: RevisionId;

  sequence: number;

  artifactRef: ArtifactRef;

  contentHash: string;

  origin:
    | "human"
    | "agent"
    | "external";

  createdAt: number;
}
```

原因：

Human Ctrl+S 产生的版本同样是 committed revision，并不是“accepted”。

---

# 37. CandidateRevision

```ts
interface CandidateRevision {
  candidateId: CandidateId;

  sessionId: SessionId;

  baseRevisionId: RevisionId;
  baseHash: string;

  artifactRef: ArtifactRef;

  currentHash?: string;

  state: CandidateState;

  createdBy:
    | "agent"
    | "human";

  verification?: VerificationReport;
}
```

UI 可以叫：

```text
AI Proposal
```

Runtime 内部保持：

```text
CandidateRevision
```

---

# 38. WriterLease

```ts
interface WriterLease {
  leaseId: LeaseId;

  sessionId: SessionId;

  owner:
    | "human"
    | "agent"
    | "external";

  backend:
    | "genoffice"
    | "officecli"
    | "powerpoint"
    | "wps";

  baseRevisionId: RevisionId;

  fencingToken: bigint;

  sessionEpoch: number;
}
```

---

# 39. Fencing Token

必须：

```text
monotonic
```

例如：

```text
100
101
102
```

旧 writer：

```text
token 100
```

即使晚返回：

```text
current token = 102
```

立即拒绝。

---

# 40. MVCC 读取语义

WriterLease 不阻塞 Reader。

例如：

```text
Agent 写 Candidate 43
```

用户仍可：

```text
阅读 Revision 42
```

只有另一个 Writer 请求才发生协调。

---

# 41. Session Actor

每个 DocumentSession 是一个逻辑 Actor。

```text
Session Actor
 ├─ beginEdit
 ├─ beginAgentTask
 ├─ registerCandidate
 ├─ publishVerification
 ├─ accept
 └─ reject
```

同一 Session 的：

```text
state-changing operations
```

严格串行。

---

# 42. Actor 不运行重任务

禁止：

```text
Session Actor
→ hash 1GB
→ render
→ XML diff
```

正确：

```text
Actor
 ↓ dispatch immutable job
Worker/Scheduler
 ↓
Result
 ↓
Actor validates freshness
```

---

# 43. Async Work Staleness

```ts
interface AsyncWorkIdentity {
  sessionId: SessionId;
  sessionEpoch: number;

  artifactRef: ArtifactRef;

  candidateId?: CandidateId;

  fencingToken?: bigint;
}
```

完成后如果 identity stale：

```text
discard
```

---

# 44. EditorPlugin

Harness Core 不直接依赖 GenOffice API。

```ts
interface OfficeEditorPlugin {
  format: OfficeFormat;

  create(
    context: EditorBootstrapContext
  ): Promise<EditorInstance>;
}
```

---

# 45. EditorInstance

```ts
interface EditorInstance {
  instanceId: string;

  mount(
    container: HTMLElement
  ): Promise<void>;

  activateEdit(): Promise<void>;

  save(): Promise<EditorSaveResult>;

  reload(
    input: ReloadRequest
  ): Promise<void>;

  getSelection():
    Promise<SelectionAnchor | null>;

  suspend(): Promise<void>;

  resume(): Promise<void>;

  dispose(): Promise<void>;
}
```

---

# 46. Native Editor 默认同 Renderer

第一版：

```text
InProcessEditorHost
```

原因：

```text
最快
最低 IPC
Focus 简单
Selection 简单
```

只有 Benchmark 证明：

```text
严重内存泄漏
Editor crash isolation 需求
```

才实现：

```text
IsolatedEditorHost
```

---

# 47. EditorHost 抽象

预留：

```ts
interface EditorHost {
  mount(
    plugin: OfficeEditorPlugin
  ): Promise<EditorInstance>;
}
```

未来可实现：

```text
InProcessEditorHost
IsolatedEditorHost
```

而 Document Runtime 不变。

---

# 48. Editor Surface 隔离

GenOffice 被直接嵌入 Harness Renderer 后必须限制：

```text
CSS
Keyboard
Global Events
Portals
Drag/drop
Focus
z-index
Clipboard
```

推荐：

```text
office-editor-root
+
CSS scoping
+
plugin-owned overlay root
+
Harness FocusManager
```

---

# 49. 不建议第一版强制 Shadow DOM

Tiptap / Univer / floating overlay 等可能增加 Shadow DOM 兼容成本。

第一阶段：

```text
Scoped CSS
CSS Layers
Dedicated Overlay Root
```

优先。

---

# 50. Keyboard Ownership

当 Office Editor focus：

```text
Ctrl+B
Ctrl+Z
Delete
Arrow
```

交给 Editor。

Harness 保留：

```text
Global Agent shortcut
Command Palette
Workbench commands
```

所有 global key listener 必须进入：

```text
CommandScope
```

---

# 51. Progressive Editor Activation

Open 后默认：

```text
Read
```

第一次 mutation：

```text
request Edit Promotion
```

流程：

```text
check source
 ↓
check no competing Candidate
 ↓
ensure WriteReadiness
 ↓
acquire Human WriterLease
 ↓
activate editable runtime
 ↓
replay input
```

---

# 52. WriteReadiness

```ts
type WriteReadiness =
  | "cold"
  | "probing"
  | "ready"
  | "blocked";
```

正式 Open 后可以低优先级准备：

```text
strong hash
package manifest
permission
safe-open check
editor JS prefetch
```

但不获取 WriterLease。

---

# 53. Warmup Admission

Quick Preview 不应该自动：

```text
hash 2GB file
```

只有：

```text
正式 Open
停留
Selection
Agent referenced
```

才进入强 Write Warmup。

---

# 54. View Bookmark

Preview/Open→Edit 时保持：

```text
slide
sheet
range
scroll
zoom
selection
```

接口：

```ts
interface ViewBookmark {
  location: unknown;
  zoom?: number;
}
```

避免 Edit 激活后跳回文档开头。

---

# 55. SelectionAnchor

不能只存：

```text
/slide[3]/shape[5]
```

必须 revision-aware。

```ts
interface SelectionAnchor {
  revisionId: RevisionId;

  format: OfficeFormat;

  logicalPath?: string;

  stableId?: string;

  containerId?: string;

  fingerprint?: string;
}
```

Agent真正修改前：

```text
resolve
probe
confirm
mutate
```

继续继承原设计 Read → Probe → Mutate 的安全思想。

---

# 56. Agent Runtime

Agent 不直接运行：

```text
OfficeCLI raw path
```

而是：

```text
Agent
 ↓
MCP Office Tools
 ↓
OfficeRuntimeService
 ↓
OfficeCliAdapter
 ↓
OfficeCLI
```

---

# 57. MCP Model Tool Surface

保持小：

```text
office.inspect
office.query
office.edit
office.render
office.verify
office.capabilities
```

禁止暴露：

```text
filesystem
shell
raw process management
editor reload
internal range fetching
```

---

# 58. DocumentCapability

Agent 不获取真实 path。

```ts
interface DocumentCapability {
  documentId: DocumentId;
  sessionId: SessionId;

  permissions: {
    read: boolean;
    edit: boolean;
  };
}
```

关闭文件/撤销权限后：

```text
capability revoked
```

---

# 59. OfficeTaskContext

```ts
interface OfficeTaskContext {
  taskId: string;

  sessionId: SessionId;
  documentId: DocumentId;

  candidateId: CandidateId;

  baseRevisionId: RevisionId;

  fencingToken: bigint;

  mutationScope: MutationScope;
}
```

Agent不需要直接操作内部 revision machinery。

---

# 60. MutationScope

```ts
interface MutationScope {
  intent: string;

  allowedTargets?: string[];

  allowedParts?: string[];

  destructiveAllowed: boolean;
}
```

例如：

```text
用户：
修改第三页标题颜色

Allowed:
slide 3
target shape
font color

Not Allowed:
theme
master
other slides
```

---

# 61. Idempotent Commands

所有 mutating Tool 请求：

```ts
interface MutationCommand<T> {
  commandId: string;

  idempotencyKey: string;

  candidateId: CandidateId;

  fencingToken: bigint;

  payload: T;
}
```

同一个 idempotencyKey 重试：

```text
返回旧结果
```

而不是再次执行。

---

# 62. MutationReceipt

```ts
interface MutationReceipt {
  receiptId: string;

  candidateId: CandidateId;

  engine: "officecli";

  affectedTargets: string[];

  completedAt: number;
}
```

便于 Audit / Debug。

---

# 63. OfficeCLI Execution Policy

不要永远 Resident。

```ts
type OfficeCliExecutionMode =
  | "standalone"
  | "resident";
```

推荐：

| Task | Mode |
|---|---|
| 一次简单 batch | standalone |
| 一次性改几个对象 | standalone |
| inspect/edit/inspect 多轮循环 | resident |
| 视觉反馈持续修复 | resident |

OfficeCLI 当前 batch 在 standalone 模式一次打开、执行并保存；当前文档说明 v1.0.137+ 默认 atomic rollback。

---

# 64. OfficeCLI Resident

Resident：

```text
mutation
→ in-memory
```

外部 Reader：

```text
save
→ flush disk
```

完成 Agent Writer：

```text
close
→ flush + release
```

这是正式协议边界。

---

# 65. Resident Pool

OfficeCLI resident 是：

```text
cacheable runtime resource
```

不是 DocumentSession 必备。

ResourceGovernor 负责：

```text
max concurrency
idle eviction
memory pressure eviction
```

Candidate Review 阶段通常可关闭 resident。

---

# 66. Candidate 创建

Agent Task：

```text
Committed Artifact
 ↓
Candidate Clone
```

优先：

```text
CoW / reflink
```

fallback：

```text
normal copy
```

禁止用 hard link 作为安全隔离。

---

# 67. Candidate 生命周期

```text
PREPARING
 ↓
MUTATING
 ↓
FLUSHING
 ↓
VERIFYING
 ↓
READY
 ↓
REVIEW
```

用户：

```text
Accept
Reject
Edit Proposal
```

---

# 68. Candidate Verification Hash Binding

Verification 永远绑定：

```text
candidate contentHash
```

如果 Human Edit Proposal：

```text
hash changes
```

则旧：

```text
VerificationReport
```

立即失效。

Accept 前必须重新 Verify。

---

# 69. Proposal Review 默认 Read-only

Agent Candidate Ready 后：

```text
Review
```

默认不能直接编辑。

用户主动：

```text
Edit Proposal
```

才：

```text
OfficeCLI close
 ↓
Human Candidate WriterLease
 ↓
GenOffice editable
```

---

# 70. Ctrl+S ≠ Accept

Proposal Editor 中：

```text
Ctrl+S
=
save candidate
```

绝不能：

```text
Accept
```

Accept 是显式独立动作。

---

# 71. Accept 不应该重新 Parse

如果 GenOffice 当前正在看：

```text
Candidate hash X
```

Accept：

```text
Candidate X
→ Committed Revision N+1
```

bytes 没变。

所以：

```text
ArtifactContext X
```

继续使用。

只更新：

```text
revision binding
candidate status
source mapping
```

不 full reload。

---

# 72. Context Promotion

正式定义：

> Logical role changes do not invalidate an unchanged ArtifactContext.

即：

```text
Candidate ArtifactContext
→
Committed ArtifactContext
```

零重复 parse。

---

# 73. SelfWriteGuard

Commit source 后会收到 filesystem watcher event。

必须识别：

```text
self-originated mutation
```

避免：

```text
commit
→ watcher
→ external conflict
→ reload
```

接口：

```ts
interface SelfWriteGuard {
  commitId: string;
  expectedHash: string;
  expiresAt: number;
}
```

---

# 74. AtomicFileCommitter

所有 source replacement：

```text
Human Save
Agent Accept
External Import
```

统一经过：

```ts
interface AtomicFileCommitter {
  commit(
    input: CommitRequest
  ): Promise<CommitResult>;
}
```

---

# 75. Commit 流程

```text
verify expected source hash
 ↓
prepare sibling temp
 ↓
fsync
 ↓
replace source
 ↓
rehash
 ↓
finalize revision
```

不同 OS 需要平台实现。

尤其 Windows 上必须处理文件句柄/replace 差异。

---

# 76. CommitJournal

DB 与 filesystem 无法形成一个真正 ACID transaction。

因此：

```text
PREPARED
 ↓
TEMP_READY
 ↓
SOURCE_REPLACED
 ↓
FINALIZED
```

记录：

```ts
interface CommitJournal {
  commitId: string;

  sourceHashBefore: string;
  candidateHash: string;

  phase:
    | "prepared"
    | "temp-ready"
    | "source-replaced"
    | "finalized";
}
```

---

# 77. Crash Recovery

Recovery 真相优先级：

```text
Filesystem hash facts
      >
CommitJournal
      >
Session metadata
```

不能仅根据：

```text
state = committing
```

猜测。

---

# 78. ArtifactScanner

避免重复读大文件。

Strong scan 一次产生：

```ts
interface ArtifactScanResult {
  contentHash: string;

  size: bigint;

  packageManifest: PackageManifest;

  integrity: IntegrityResult;
}
```

供：

```text
Hash
Verification
PackageDiff
Revision metadata
```

复用。

---

# 79. FastArtifactProbe

Preview/Open 热路径只执行：

```text
stat
format
ZIP central directory
basic metadata
```

不运行：

```text
full SHA
full XML validation
full render
```

---

# 80. Package Diff

继续继承原设计 Package Diff 的思想：原设计会比较 OOXML package parts，并识别用户只修改 slide3 却额外改 theme/master/embedding 等异常情况。

新版：

```text
Raw Part Diff
 ↓
Safe Normalization
 ↓
Relationship Diff
 ↓
Risk Classification
```

分类：

```text
UNCHANGED
NORMALIZATION_ONLY
EXPECTED
UNEXPECTED_LOW_RISK
UNEXPECTED_HIGH_RISK
```

---

# 81. Verification Pipeline

不再只有：

```text
PASS / FAIL
```

定义分层 Confidence。

```text
L0 artifact integrity
L1 OOXML structural
L2 package relationships
L3 OfficeCLI issues
L4 mutation-aware diff
L5 changed-scope render
L6 GenOffice preview
L7 PowerPoint/WPS certification
```

原设计已经使用 Schema、Issues、Package Diff、Host reopen、Native render、Vision QA 的多层验证，这一思想继续保留。

---

# 82. VerificationReport

```ts
interface VerificationReport {
  candidateId: CandidateId;

  contentHash: string;

  structural: CheckResult;
  package: CheckResult;
  semantic: CheckResult;
  visual: CheckResult;

  confidence:
    | "unverified"
    | "structural"
    | "engine"
    | "visual"
    | "consumer-certified";
}
```

---

# 83. Changed Scope First

Agent只改：

```text
slide 17
```

优先：

```text
slide 17
相关 relationships
相关 media/chart
render slide 17
```

不要默认 render 全部 150 slides。

Full validation：

```text
background
high-risk
release
```

再运行。

---

# 84. ChangeImpact

```ts
interface ChangeImpact {
  scope:
    | "document"
    | "slide"
    | "sheet"
    | "range"
    | "block";

  targets?: string[];

  confidence:
    | "exact"
    | "conservative"
    | "unknown";
}
```

第一版 Editor 可以：

```text
full reload
```

未来：

```text
slide/sheet/block partial reload
```

协议不变。

---

# 85. PortableReviewView

MCP View 不再承担完整 Office Editor。

主要用途：

```text
preview
before/after
verification status
accept/reject
basic navigation
```

Generic MCP Host 可以使用。

Harness 自己：

```text
Native Editor
```

优先。

---

# 86. Optional PowerPoint / WPS Adapter

Native Host 不再是主路径。

职责：

```text
Selection
Host-specific advanced features
Certification
External attached workflow
```

原 Attached Mode 的 Native Selection、SaveCopyAs Shadow、Probe、Native edit、Host render 等能力仍然可以保留为 Optional Adapter。

---

# 87. External Host Backend

```ts
type DocumentBackend =
  | "managed-file"
  | "external-powerpoint"
  | "external-wps";
```

PowerPoint unsaved in-memory Presentation：

不能伪装成普通 committed filesystem revision。

External Host 保持独立语义。

---

# 88. Certification Copy

PowerPoint/WPS Verification 永远使用：

```text
Disposable Certification Copy
```

而不是 Candidate 本体。

如果 Host：

```text
Repair
Normalize
Save metadata
```

都不会改变 Candidate。

自动 Repair：

```text
Certification FAIL
```

---

# 89. Resource 模型

整个 Office Plugin 内存正式拆成：

```text
Artifact Layer
Visual Layer
Interaction Layer
Native Layer
```

---

# 90. Artifact Layer

包括：

```text
ZIP index
relationships
styles
metadata
sheet chunks
slide parse
block index
```

尽量：

```text
shareable
immutable
evictable
```

---

# 91. Visual Layer

包括：

```text
thumbnail
decoded image
page bitmap
canvas texture
preview
```

高度可回收。

---

# 92. Interaction Layer

包括：

```text
DOM
selection
undo
editing commands
caret
overlay
```

主要属于：

```text
EditorInstance
```

通常不能跨文档共享。

---

# 93. Native Layer

包括：

```text
OfficeCLI resident
Rust sidecar
worker
utility process
```

由 ResourceGovernor 管理。

---

# 94. Residency

统一：

```ts
type Residency =
  | "hot"
  | "warm"
  | "cold"
  | "evicted";
```

用于：

```text
ArtifactContext
Editor
Preview
OfficeCLI resident
Cache
```

---

# 95. 多 Tab 策略

例如 10 个 Tab：

```text
1 HOT
2 WARM
7 COLD
```

而不是：

```text
10 full GenOffice editors
```

DocumentSession 可以一直存在。

EditorInstance 可以被 suspend/dispose。

---

# 96. Hot

当前 Tab：

```text
full editor
visible assets
selection
```

---

# 97. Warm

最近 Tab：

```text
ArtifactContext
limited layout cache
possibly suspended editor
```

---

# 98. Cold

```text
DocumentSession
ViewBookmark
Artifact metadata
```

Editor Runtime 已释放。

---

# 99. Cache 类型必须分开

不要一个 GlobalCache。

正式区分：

```text
ArtifactCache
VisualPreviewCache
DecodedAssetCache
EditorRuntimeCache
```

生命周期不同。

---

# 100. Cache Admission

不仅要：

```text
怎么 Evict
```

还要：

```text
值不值得 Cache
```

规则：

```text
small + expensive + reused
→ cache

huge + one-shot
→ skip

decoded image
→ hot only

thumbnail
→ disk cache friendly
```

---

# 101. Cache 不按对象数限制

禁止：

```text
100 images
```

应使用：

```text
byte budget
```

因为 decoded bitmap 成本和 encoded file size 差别很大。

---

# 102. Encoded / Decoded Asset

图片拆：

```text
EncodedAssetStore
DecodedVisualCache
```

同一 Logo 在不同 slide/revision 可共享 encoded bytes。

Decoded representation 只在 visible hot scope 保留。

---

# 103. FontEnvironment

系统字体信息、fallback、metrics：

```text
Process-wide shared
```

不要每个文档重复初始化。

定义：

```text
fontEnvironmentId
```

用于：

```text
layout cache
preview cache
renderer invalidation
```

---

# 104. Preview Cache Key

```ts
interface PreviewCacheKey {
  contentHash: string;

  rendererVersion: string;

  fontEnvironmentId: string;

  previewProfile: string;
}
```

避免 renderer/font 改变后仍使用旧图。

---

# 105. ResourceGovernor

成为 P0 基础设施。

```ts
interface ResourceGovernor {
  acquire(
    request: ResourceRequest
  ): Promise<ResourceLease>;

  getPressure(): ResourcePressure;

  release(
    lease: ResourceLease
  ): void;
}
```

---

# 106. Resource Budget

至少分：

```text
Memory
CPU
I/O
Render/GPU
Native Process
Disk Cache
```

不能只看 JS heap。

---

# 107. Memory Pressure

```text
normal
elevated
high
critical
```

回收：

```text
expired preview
 ↓
decoded visuals
 ↓
prefetch
 ↓
idle residents
 ↓
warm editors
 ↓
cold artifact cache
```

永远不自动丢：

```text
unsaved Human state
Candidate
Commit Journal
```

---

# 108. FormatRuntime.trimMemory

格式自己知道该怎么释放。

```ts
trimMemory(
  "light" | "moderate" | "aggressive"
)
```

例如 Sheets：

```text
light
→ shrink viewport buffer

moderate
→ drop inactive sheet chunks

aggressive
→ destroy idle sidecar
```

---

# 109. Session Actor 不做大任务

正式 Performance Invariant：

```text
Session Actor
=
Control Only
```

所有：

```text
hash
render
package scan
verification
```

进入 Scheduler。

---

# 110. Scheduler Priority

建议：

```text
INTERACTIVE
VISIBLE_PREVIEW
EDIT_PROMOTION
AGENT_FOREGROUND
VERIFICATION
PREFETCH
BACKGROUND_INDEX
CACHE_BUILD
```

前台永远高于后台。

---

# 111. CPU / IO / Render 分预算

避免：

```text
4 个 hash job
```

把：

```text
当前 slide render
```

堵住。

逻辑上至少：

```text
IOScheduler
RenderScheduler
NativeProcessGovernor
```

---

# 112. Backpressure

队列不能无限增长。

例如：

```text
Agent render 100 slides
```

新 Revision 已经产生：

```text
旧 preview work
→ cancel/deprioritize
```

---

# 113. Binary IPC

控制：

```text
typed object
```

大型数据：

```text
ArrayBuffer
```

优先 Transfer。

禁止：

```text
Base64
```

作为大型本地 IPC 常规路径。

---

# 114. Buffer Ownership

```ts
interface BinaryPayload {
  buffer: ArrayBuffer;

  ownership:
    | "transfer"
    | "copy";
}
```

默认大数据：

```text
transfer
```

---

# 115. Persistence

Session Runtime Metadata 推荐：

```text
SQLite + WAL
```

数据库只由：

```text
OfficeRuntimeService
```

拥有。

Renderer 不直接连接 DB。

---

# 116. Durable vs Ephemeral Events

持久化：

```text
SessionEvent
```

例如：

```text
candidate.created
candidate.verified
revision.committed
```

不持久化：

```text
EditorSignal
```

例如：

```text
selectionChanged
viewportChanged
focusChanged
scroll
```

---

# 117. SessionEvent

```ts
interface SessionEvent<T> {
  eventId: string;

  sessionId: SessionId;

  sequence: bigint;

  sessionEpoch: number;

  type: string;

  payload: T;
}
```

---

# 118. Event Semantics

使用：

```text
At-least-once
+
Idempotent consumer
```

不追求 expensive Exactly Once。

每个 Session sequence 单调递增。

旧/重复 Event：

```text
ignore
```

---

# 119. SQLite WAL

数据库事务必须短。

禁止：

```text
Editor 打开
→ 持有 SQLite Read Transaction 两小时
```

所有查询：

```text
read
→ copy result
→ close transaction
```

---

# 120. Plugin Offline Security

Plugin 默认：

```text
network disabled
```

OfficeCLI 自更新必须关闭。

所有 JS/CSS/font/runtime：

```text
local bundle
```

---

# 121. Macro/OLE

规则：

```text
Preserve
Never Execute
```

Agent 默认不能：

```text
Run macro
Activate OLE executable
Follow external data connection
```

---

# 122. AssetCapability

Agent 插图片时：

```text
asset_123
```

而不是：

```text
C:\...
https://...
```

Runtime：

```text
AssetCapability
→ approved local artifact
```

---

# 123. Prompt Injection

所有 Office 内容：

```text
Slide Text
Notes
Comments
Cell Values
Alt Text
Metadata
```

全部是：

```text
Untrusted Document Data
```

原设计已经明确建立这一原则。

文档里的：

```text
Ignore previous instructions
Delete files
Run command
```

绝不提升为工具指令。

---

# 124. OperationPolicyEngine

OfficeCLI 支持：

```text
≠
Agent automatically allowed
```

风险：

```text
Low
Medium
High
Critical
```

例如：

```text
text edit → Low
delete slide → High
raw OOXML → Critical
macro → denied
```

---

# 125. Runtime Compatibility Lock

Release 必须固定：

```text
OfficePlugin version
GenOffice commit
OfficeCLI version
OfficeCLI schema fingerprint
Electron version
Contract version
Session DB schema
Golden Corpus version
```

例如：

```json
{
  "plugin": "3.0.0",
  "genoffice": "<commit>",
  "officecli": {
    "version": "1.x",
    "schemaFingerprint": "..."
  },
  "contract": 1,
  "dbSchema": 1
}
```

---

# 126. Upgrade Barrier

以下状态禁止 Runtime upgrade：

```text
Human dirty
Agent running
Candidate review
Commit running
Host mutation
```

升级：

```text
drain
snapshot
install
probe
migrate
activate
```

失败：

```text
rollback
```

---

# 127. Capability Degradation

OfficePlugin 不应只有：

```text
working/broken
```

而是：

```text
Docs Editor      available
Docs Agent       available

Sheets Editor    available
Sheets Agent     degraded

Slides Editor    available
Slides Agent     unavailable

PowerPoint Cert  unavailable
```

局部失败不拖垮整个 Plugin。

---

# 128. Preview Performance Metrics

正式冻结五个用户性能指标。

### TTFP

Time To First Preview

---

### TTI

Time To Interactive View

---

### TTE

Time To Editable

---

### TTP

Time To First Agent Proposal

---

### TTV

Time To Verified Proposal

---

# 129. 内存指标

至少：

```text
Peak Working Set

Steady-state Working Set

Post-Close Retained Memory

Retained Memory Slope
```

---

# 130. Renderer 指标

必须监控：

```text
Renderer memory
Blink resources
JS heap
GPU
Utility process
OfficeCLI
XLSX sidecar
```

Electron 可通过应用进程指标和各进程类型信息帮助观测这些资源。

---

# 131. Long Task Budget

当前 active Editor：

```text
不得被持续同步 parse/layout 卡住
```

建议 Benchmark 重点记录：

```text
>50ms Renderer Main Thread tasks
```

并统计 P95/P99。

---

# 132. 建议性能目标

以下为 PoC 设计目标，不是现有引擎承诺：

| 场景 | Initial Goal |
|---|---:|
| Cached Preview | <100ms |
| 常规 PPT/DOC First Useful Preview | <500ms |
| 常规 XLSX First Viewport | <700ms |
| Warm Open → Edit | <300ms |
| Cold Open → Edit | <1s |
| Agent Commit → Proposal Visible | 尽量 <500ms，不含模型推理 |
| Candidate Accept | 不发生不必要 Full Reload |
| UI Main Thread | 避免连续 Long Tasks |

真正阈值以 Benchmark 调整。

---

# 133. Benchmark Corpus

必须包含：

```text
small / medium / large DOCX
small / medium / large XLSX
small / medium / large PPTX
image-heavy PPT
chart-heavy PPT
formula-heavy XLSX
enterprise templates
CJK fonts
comments
tracked changes
embedded media
unsupported extension objects
```

---

# 134. 高频 Preview Benchmark

```text
连续 Preview 100 个文件
```

测：

```text
TTFP
I/O
CPU
memory
cache hit rate
cancellation
```

验收：

```text
memory 不应接近线性增长
```

---

# 135. Open/Close Benchmark

```text
Open
Close
×100
```

测：

```text
T+0
T+2s
T+10s
T+30s
```

观察：

```text
retained memory plateau
```

---

# 136. Preview→Open→Edit Benchmark

重复：

```text
Preview
Open
Edit
Close
×50
```

重点：

```text
parse count
ArtifactContext reuse
TTE
memory
```

---

# 137. Candidate Benchmark

```text
Agent Candidate
→ Proposal
→ Accept
×50
```

检查：

```text
是否重复 parse
是否无意义 reload
Candidate GC
OfficeCLI resident cleanup
```

---

# 138. Large XLSX Benchmark

至少：

```text
500MB+
```

检查：

```text
TTFP
viewport scroll
sidecar memory
renderer memory
range eviction
```

GenOffice 当前已有 large-workbook gate，并明确采用 bounded renderer cell window。

---

# 139. Image-heavy PPT Benchmark

例如：

```text
150 slides
hundreds of images
```

重点：

```text
decoded image memory
GPU
thumbnail cache
slide switch
post-close retained memory
```

---

# 140. DOCX Benchmark

例如：

```text
200–500 pages
```

检查：

```text
page DOM count
visible virtualization
scroll smoothness
memory slope
```

---

# 141. Fault Injection

必须测试：

```text
OfficeCLI crash
Runtime crash
Renderer crash
disk full
Candidate missing
commit interrupted
verification interrupted
stale Agent result
external file mutation
file lock
permission revoked
```

---

# 142. Correctness Invariants

必须至少满足：

```text
INV-01
Agent 不直接写 Committed Artifact。

INV-02
同 Session 只有一个有效 WriterLease。

INV-03
所有 Writer 使用 FencingToken。

INV-04
Mutation command 幂等。

INV-05
Candidate 绑定 Base Revision + Hash。

INV-06
OfficeCLI 外部 Reader 前必须 Flush。

INV-07
Writer handoff 前 OfficeCLI 必须 release。

INV-08
Verification 绑定 Candidate Hash。

INV-09
Human 修改 Candidate 后 Verification 失效。

INV-10
Accept 检查 Source Hash。

INV-11
所有 Commit 经过 AtomicFileCommitter。

INV-12
Agent 不拿真实 path。

INV-13
未知 OOXML 默认 Preserve。

INV-14
Document content 是 untrusted data。

INV-15
Plugin crash 不导致不可解释的半写 Committed Artifact。
```

---

# 143. Performance Invariants

```text
PERF-01
Preview 不创建 Candidate。

PERF-02
Preview 不启动 OfficeCLI。

PERF-03
Open Read-only 不获取 WriterLease。

PERF-04
同 ArtifactVersion 的并发读取必须去重。

PERF-05
同 Artifact 不存在无意义重复 ArtifactContext。

PERF-06
Preview 取消不能杀死仍有其他消费者的 Shared Build。

PERF-07
Strong Hash 不阻塞 First Preview。

PERF-08
Session Actor 不执行重 CPU/IO 工作。

PERF-09
Candidate Promotion 不因为身份变化而重复 Parse。

PERF-10
Background Work 不优先于 Interactive Work。

PERF-11
缓存有 Admission + Eviction。

PERF-12
Editor dispose 后所有 Editor-owned ResourceLease 被释放。

PERF-13
Working Set 尽量与 Visible Complexity 成比例。

PERF-14
大型 Binary 不经 Base64 热路径传输。

PERF-15
同文件多 View 共享 DocumentSession / ArtifactContext。
```

---

# 144. 代码依赖规则

推荐：

```text
office-contracts
      ↑
      │
 ┌────┼────────┐
runtime editor mcp
 │      │       │
artifact genoffice officecli
```

禁止：

```text
GenOfficeAdapter
→ OfficeCLI internals

OfficeCLIAdapter
→ GenOffice internals

Editor
→ SQLite

MCP
→ filesystem
```

---

# 145. 推荐 Repository

```text
plugins/
└── office/
    ├── manifest/
    ├── contracts/
    │   ├── artifact.ts
    │   ├── document.ts
    │   ├── revision.ts
    │   ├── candidate.ts
    │   ├── lease.ts
    │   ├── events.ts
    │   ├── editor.ts
    │   └── verification.ts
    │
    ├── runtime/
    │   ├── service/
    │   ├── sessions/
    │   ├── revisions/
    │   ├── candidates/
    │   ├── scheduler/
    │   ├── resources/
    │   ├── commit/
    │   ├── recovery/
    │   └── persistence/
    │
    ├── artifact/
    │   ├── registry/
    │   ├── store/
    │   ├── scanner/
    │   └── cache/
    │
    ├── preview/
    │
    ├── editors/
    │   ├── common/
    │   ├── docs/
    │   ├── sheets/
    │   └── slides/
    │
    ├── agent/
    │   └── officecli/
    │
    ├── mcp/
    │   ├── tools/
    │   └── portable-review/
    │
    ├── verification/
    │
    ├── hosts/
    │   ├── powerpoint/
    │   └── wps/
    │
    ├── native/
    │
    └── vendor/
        └── genoffice/
```

---

# 146. GenOffice Fork Policy

尽量不修改：

```text
docx-engine core
pptx-engine core
pptx-render core
xlsx fidelity core
```

主要增加：

```text
EditorAdapter
ArtifactContext integration
HostDocumentBridge
Focus/Keyboard integration
CSS isolation
NativeDataBridge
AI/cloud removal
```

当前 GenOffice 本身已经采用多 Electron app + shared engine layer 的组织方式，而且 Sheets 明确通过 `WorkbookAdapter` 避免产品代码直接依赖 Univer API。

这与本设计的 Adapter 策略高度一致。

---

# 147. 开发阶段

## Phase 0 — Contracts

实现：

```text
ArtifactRef
ArtifactVersionKey
ArtifactContext
ArtifactLease
DocumentSession
CommittedRevision
CandidateRevision
WriterLease
SessionEvent
EditorSignal
VerificationReport
```

不接 GenOffice/OfficeCLI。

---

## Phase 1 — Artifact Runtime

实现：

```text
ArtifactStore
ArtifactRegistry
In-flight Dedup
PreviewRequest
FastArtifactProbe
StrongArtifactScan
```

先用简单测试 Renderer。

---

## Phase 2 — PPTX Read PoC

接：

```text
GenOffice pptx-engine
pptx-render
```

实现：

```text
Preview
Open
```

重点测：

```text
TTFP
Memory
Repeated open
```

---

## Phase 3 — Slides Edit

实现：

```text
EditorAdapter
Progressive Edit Activation
SelectionAnchor
Human Save
```

---

## Phase 4 — Agent Candidate

接：

```text
OfficeCLI
MCP Tools
Candidate
WriterLease
Flush
Verification
Proposal
Accept/Reject
```

---

## Phase 5 — Runtime Hardening

实现：

```text
Fencing
Idempotency
CommitJournal
SelfWriteGuard
Recovery
SQLite
```

---

## Phase 6 — Performance Governance

实现：

```text
ResourceGovernor
Residency
Cache Admission
Memory Telemetry
Scheduler
```

---

## Phase 7 — DOCX

接 GenOffice Docs。

---

## Phase 8 — XLSX

保留：

```text
Rust sidecar
viewport loading
bounded renderer window
```

---

## Phase 9 — Optional Native Hosts

实现：

```text
PowerPointAdapter
WpsAdapter
```

---

# 148. MVP 范围

MVP 不要求：

```text
Part-level Cache
Object-level Reload
IsolatedEditorHost
Multi-document Group Commit
CRDT
Per-operation Accept
Full PowerPoint Certification
```

MVP 必须：

```text
Preview
Open
Edit
Save

Agent Candidate
Review
Accept
Reject

Crash-safe Commit
Basic Verification
```

---

# 149. P2 优化

只有 Benchmark 证明需要才做：

```text
Slide-level partial reload
Sheet-level partial reload
DOCX block reload

Part-level OOXML cache

IsolatedEditorHost

Process-pooled Preview Renderer

Advanced decoded asset dedup
```

---

# 150. 不允许提前实现

禁止因为“以后可能需要”提前增加：

```text
Canonical Office IR
Custom Office DOM
Universal Shape API
Universal Animation API
Complex cross-plugin RPC
CRDT
Multi-user synchronization
```

除非有新的 ADR。

---

# 151. 架构 Review Checklist

每一个新功能必须回答：

### Read Path

是否增加 Preview/Open 延迟？

是否重复 Parse？

是否不必要获得 WriterLease？

---

### Memory

是否创建新的长期 cache？

谁拥有？

何时释放？

是否有 byte budget？

---

### Write Safety

是否绕过 Candidate？

是否绕过 Lease？

是否绕过 Atomic Commit？

---

### Agent

是否泄漏真实 path？

是否扩大 Tool 权限？

是否支持 idempotency？

---

### Compatibility

是否触碰未知 OOXML？

是否能 Preserve？

---

### Runtime

是否在 Session Actor 中执行重任务？

---

### Renderer

是否增加 global listener？

是否能完全 dispose？

---

# 152. Performance Review Checklist

新增热点必须说明：

```text
CPU cost
I/O cost
memory cost
GPU cost
IPC payload
cache lifecycle
cancellation behavior
```

不允许只说：

> 应该不会慢。

必须 Benchmark。

---

# 153. Memory Review Checklist

每个重对象必须有：

```text
Owner
Size estimate
Residency
Admission
Eviction
Dispose
Metrics
```

如果回答不了：

> 谁最终释放这个对象？

不能进入主分支。

---

# 154. Naming Baseline

正式冻结以下名称：

```text
OfficePlugin

OfficeRuntimeService

PreviewRequest

ArtifactRef

ArtifactVersionKey

ArtifactContext

ArtifactRegistry

ArtifactLease

FormatRuntime

DocumentSession

CommittedRevision

CandidateRevision

WriterLease

FencingToken

SessionEpoch

EditorInstance

EditorHost

DocumentCapability

AssetCapability

SessionEvent

EditorSignal

ChangeImpact

VerificationReport

ResourceGovernor
```

原则上不再改名。

---

# 155. Architecture Decision Records

建议正式建立：

```text
ADR-001 One Office Plugin

ADR-002 No Canonical Office IR

ADR-003 Native Editor First

ADR-004 MCP Control Plane

ADR-005 Native Data Plane

ADR-006 Read Optimized Architecture

ADR-007 MVCC-like Immutable Reads

ADR-008 ArtifactRegistry Owns Read Contexts

ADR-009 Preview Does Not Create DocumentSession

ADR-010 Progressive Editor Activation

ADR-011 Single Writer Lease + Fencing

ADR-012 Candidate Before Commit

ADR-013 Explicit OfficeCLI Flush Barrier

ADR-014 Candidate Verification Hash Binding

ADR-015 Atomic Commit + Journal

ADR-016 GenOffice Adapter Boundary

ADR-017 ResourceGovernor Owns Global Budgets

ADR-018 Session Actor Is Control-only

ADR-019 Portable MCP View Is Not Main Editor

ADR-020 PowerPoint/WPS Are Optional Host Adapters

ADR-021 Offline Is Default

ADR-022 Artifact Context Promotion Avoids Reload

ADR-023 Performance Changes Require Benchmark

ADR-024 Isolated Editor Requires Benchmark Evidence
```

---

# 156. 最终开发架构

```text
                         HARNESS
                            │
                       OfficePlugin
                            │
          ┌─────────────────┼─────────────────┐
          │                 │                 │
          ▼                 ▼                 ▼
   Artifact Runtime   Document Runtime   Agent Runtime
          │                 │                 │
   ArtifactRegistry   DocumentSession       MCP
          │                 │                 │
   ArtifactContext    Revision/Candidate      ▼
          │                 │             OfficeCLI
   ┌──────┴──────┐          │
   ▼             ▼          │
Preview         Open        │
                  │          │
                  ▼          │
             GenOffice       │
                  │          │
                  └────┬─────┘
                       ▼
                 Artifact Store
                       │
              DOCX / XLSX / PPTX
```

外围：

```text
ResourceGovernor
Scheduler
Verification
SQLite
CommitJournal
Recovery

Optional:
PowerPoint
WPS
PortableReviewView
```

---

# 157. 最终 Read Path

```text
USER SELECTS FILE
        │
        ▼
PreviewRequest
        │
        ▼
ArtifactRegistry
        │
    ┌───┴─────────────┐
    │                 │
cached?          build active?
    │                 │
   YES               YES
    │                 │
    └──────┬──────────┘
           │
           ▼
    ArtifactContext
           │
       First Preview
           │
      User Open?
       /      \
     NO        YES
                │
         DocumentSession
                │
         Interactive Read
                │
         User Mutation?
          /          \
        NO            YES
                      │
               Write Readiness
                      │
                Human Lease
                      │
               GenOffice Edit
```

---

# 158. 最终 Agent Write Path

```text
USER REQUEST
      │
      ▼
OfficeTaskContext
      │
      ▼
Acquire WriterLease
      │
      ▼
Create Candidate
      │
      ▼
OfficeCLI
      │
      ├─ inspect
      ├─ query
      └─ atomic batch/edit
      │
      ▼
Explicit SAVE Barrier
      │
      ▼
ArtifactScanner
      │
      ▼
Verification
      │
      ▼
Candidate Ready
      │
      ▼
OfficeCLI CLOSE
      │
      ▼
Human Review
      │
 ┌────┴──────────┐
 │               │
Reject       Edit Proposal
 │               │
 │         Human Candidate Lease
 │               │
 │             Save
 │               │
 │          Re-Verify
 │               │
 └───────┬───────┘
         ▼
       Accept
         │
     Check Fence
     Check Base
     Check Hash
     Check Verify
         │
         ▼
     CommitJournal
         │
         ▼
    Atomic Commit
         │
         ▼
Committed Revision
         │
         ▼
Context Promotion
```

---

# 159. 最终性能哲学

整个系统最终只有两个核心方向：

## Read Path

```text
Immutable
Shared
Progressive
Parallel
Mostly Lock-free
Cacheable
Evictable
```

## Write Path

```text
Exclusive
Leased
Candidate-based
Idempotent
Flushed
Verified
Journaled
Recoverable
```

两者不应该强行共享完全相同的生命周期。

---

# 160. Definition of Done — Architecture

本架构进入正式冻结，需要满足：

```text
核心 Contract 已代码化

ArtifactRegistry PoC 成功

PPT Preview/Open/Edit 跑通

OfficeCLI Candidate Workflow 跑通

Candidate Accept 不发生无意义 full reload

WriterLease / Fencing 测试通过

Crash Commit Recovery 测试通过

100 次 Preview/Open/Close 不出现明显线性内存增长

大 XLSX 保持 bounded renderer working set

Offline Network Gate 通过

Golden Corpus 基线建立
```

---

# 161. Definition of Done — Office Task

一个 Agent Office Task 只有在：

```text
操作目标完成

MutationScope 未越界

Candidate 已 Flush

Verification 与 Candidate Hash 匹配

Source Base Hash 未变化

Writer Fencing 有效

Commit 完成

Committed Revision 可恢复

Editor 指向正确 Revision
```

之后才属于：

```text
DONE
```

命令 exit code = 0：

> 不等于任务完成。

这一点继续继承原设计的最终完成定义。

---

# 162. 最终结论

本项目最终定位不是：

> 一个 GenOffice Fork。

也不是：

> 一个 OfficeCLI Wrapper。

而是：

> **一个本地、事务型、Read-Optimized 的 Office Runtime Plugin。**

其中：

```text
GenOffice
=
Human Visual Runtime

OfficeCLI
=
Agent Mutation Runtime

ArtifactRuntime
=
High-frequency Read Runtime

DocumentRuntime
=
Revision / Transaction Runtime

Harness
=
Product Host
```

系统真正长期稳定的资产不是具体 Editor 或 Agent Engine，而是：

```text
Artifact Contract
Document Contract
Editor Contract
Agent Contract
Commit Contract
Resource Contract
```

这意味着未来：

```text
GenOffice → Renderer B
```

或者：

```text
OfficeCLI → Engine B
```

不会迫使 Harness 重做文档生命周期。

最终技术原则可以浓缩成：

> **Read often, initialize progressively, share immutable work, write rarely and transactionally.**

中文即：

> **高频读取要轻、要共享、要渐进；低频写入要独占、要验证、要可恢复。**

这是本架构 v3.0 的最终基线。