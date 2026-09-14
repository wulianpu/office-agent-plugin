# Harness Office Plugin

> 一个本地、事务型、Read-Optimized 的 Office Runtime Plugin（DOCX / XLSX / PPTX）。
> 依据 [docs/Office Plugin 设计文档.md](docs/Office%20Plugin%20设计文档.md) v3.0（Architecture Freeze Candidate）完整实现。

**运行原则**：Offline First · Read Optimized · Native Editor · Agent Safe · Recoverable
**核心实现**：GenOffice（vendor @d35d770，引擎已接入）+ OfficeCLI 1.x（真实引擎）+ WPS KWPP COM（L7 认证）
**运行时依赖**：无网络、无外部服务——全部为本地 npm 包与预编译原生库（jszip / fast-xml-parser / opentype.js / bidi-js / utif2 / sharp + 本地 `genoffice-vendor` bundle），Node ≥ 24

---

## 快速开始

```bash
npm install          # 开发依赖：typescript / vitest / @types/node
npm run build        # tsc → dist/
npm test             # 56 个测试（8 单元 + 7 集成，含 INV-01…15 全量断言）
npm run corpus       # 生成基准语料（small/medium 三格式 + 73MB 大 XLSX）
npm run bench        # §134–§140 基准套件（5 项判定）
npm run demo -- <file.docx|xlsx|pptx>   # 端到端走通 §157 读取路径 + §158 Agent 写路径
```

### 一次 Agent 变更的完整生命周期（§158）

```ts
import { OfficePlugin } from "@harness/office-plugin";

const plugin = await OfficePlugin.create({ workspaceRoot: ".office-runtime" });
const ref = await plugin.registerArtifact("board.pptx");        // §15 不透明 ArtifactRef
await plugin.preview({ artifactRef: ref, priority: "visible" }); // §157 无 session/lease/engine
const session = await plugin.openSession(ref);                   // §12 只读 MVCC 会话

const task = await plugin.beginAgentTask(session.sessionId, {    // §59 任务上下文
  intent: "改第三页标题颜色",
  allowedTargets: ["/slide[3]"],                                  // §60 MutationScope
  destructiveAllowed: false
});
await plugin.executeAgentMutation(task, {                        // §61 幂等命令
  commandId, idempotencyKey: "recolor-1",
  payload: [{ command: "set", path: "/slide[3]/shape[1]", props: { color: "#ff0000" } }]
});
await plugin.flushAgentCandidate(task);      // §64 save+close 屏障 → 权威哈希
await plugin.verifyAgentCandidate(task, ["/slide[3]"]);  // §81 L0–L5 流水线
await plugin.finalizeAgentTask(task);        // §64 writer handoff（租约释放）
await plugin.acceptCandidate(session.sessionId);  // §75 原子提交 → 修订 N+1（零重复 parse）
```

---

## 架构映射（设计文档 § → 实现）

| 设计文档章节 | 契约（§154 冻结命名） | 实现 |
|---|---|---|
| §15–§27 Artifact 层 | `ArtifactRef/VersionKey/Context/Lease/Registry` | `src/artifact/registry/` — in-flight 去重、消费者感知取消、引用计数 |
| §21 §15 路径隔离 | Agent 永不接触真实路径 | `src/artifact/store/` 唯一持有 ref→path 映射（INV-12） |
| §26 FormatRuntime | 每格式进程级 Runtime | `src/artifact/runtime/`（basic 引擎；GenOffice 接入点） |
| §28–§31 读取路径 | DOCX/XLSX/PPTX 渐进渲染 | `src/preview/` — 全部流式、有界窗口（Working Set ∝ 可视复杂度） |
| §32–§43 会话层 | `DocumentSession/SessionActor/AsyncWorkIdentity` | `src/runtime/sessions/` — 严格串行 Actor、控制专用（重活走 Scheduler） |
| §13 §51 Edit Promotion | 只读→编辑原地升级 | `src/runtime/sessions/edit-promotion.ts` — 源稳定性/候选冲突/就绪/租约四道闸 |
| §35–§37 §66–§72 候选 | `CandidateRevision` 状态机 | `src/runtime/candidates/` — 验证哈希绑定 + 人工修订失效 |
| §38–§40 租约 | `WriterLease/FencingToken` | `src/runtime/sessions/lease-manager.ts` — 单写者、单调令牌、跨重启持久 |
| §73–§77 提交 | `AtomicFileCommitter/CommitJournal/SelfWriteGuard` | `src/runtime/commit/` — temp+fsync+replace、四相日志、哈希事实恢复 |
| §56 §59–§65 Agent | `OfficeTaskContext/MutationCommand/MutationReceipt` | `src/agent/officecli/` — 真实引擎适配、resident 池、幂等表 |
| §60 §121 §124 安全 | `MutationScope/OperationPolicyEngine` | 低/中/高/危/拒 风险阶梯 + 目标/部件围栏 + 宏全拒 |
| §80–§84 验证 | `VerificationReport/PackageDiff` | `src/verification/` — L0–L5、规范化识别、风险分类、changed-scope-first |
| §57–§62 §85 MCP | 六工具 `office.*` + 便携评审 | `src/mcp/` — stdio JSON-RPC（2026-01-26）、输出路径消毒 |
| §89–§112 资源 | `ResourceGovernor/Residency/缓存四类` | `src/runtime/resources/` + 字节预算 LRU + §107 逐出阶梯 + §110 优先级 + §112 背压 |
| §115–§119 持久 | SQLite WAL、短事务、事件 at-least-once | `src/runtime/persistence/` — `node:sqlite`，bigint 走 TEXT |
| §125 §127 兼容 | Runtime Compatibility Lock、能力降级 | `src/plugin/` — 矩阵如实上报（编辑器 degraded=待 GenOffice、Agent=officecli） |
| §7 §8 进程模型 | utilityProcess 宿主 | `src/native/bootstrap.ts` — RPC 分发表，Electron fork 或纯 Node 通用 |
| §146 编辑器边界 | EditorAdapter、不 fork 引擎核心 | `src/editors/` — 契约完整 + headless basic 实现（挂载/激活/保存/重载/挂起/释放） |

## Definition of Done 核对（§160）

| 架构 DoD 条目 | 证据 |
|---|---|
| 核心 Contract 已代码化 | `src/contracts/`（§154 全部冻结命名）+ tsc 通过 |
| ArtifactRegistry PoC 成功 | `test/unit/artifact-registry.test.ts`（1 build→N consumers、取消语义） |
| PPT Preview/Open/Edit 跑通 | `test/integration/read-path|human-edit.test.ts` + demo CLI |
| OfficeCLI Candidate Workflow 跑通 | `test/integration/agent-write-path.test.ts`（真引擎全链路） |
| Accept 不发生无意义 full reload | `contextPromoted=true` 断言（§71–§72 零重复 parse） |
| WriterLease / Fencing 测试通过 | `test/integration/lease-fencing.test.ts`（INV-02/03） |
| Crash Commit Recovery 测试通过 | `test/integration/crash-recovery.test.ts`（三相位 + 冲突 + INV-15） |
| 100 次 Preview/Open/Close 无线性内存增长 | bench §134：heap ±0.3MB / retained 7MB 平台 |
| 大 XLSX 保持 bounded working set | bench §138：73MB 文件、视口 60 行、堆增量 0 |
| Offline Network Gate 通过 | 零运行时依赖；officecli 子进程 `NO_UPDATE=1`；`offlineGate()` 上报 |
| Golden Corpus 基线建立 | `npm run corpus` → `.corpus/manifest.json` |

## 正确性/性能不变量覆盖（§142–§143）

**INV**：01 候选隔离 · 02 单写者 · 03 fencing · 04 幂等 · 05 基线绑定 · 06 flush 屏障 · 07 handoff 屏障 · 08 验证哈希绑定 · 09 人工修订失效 · 10 源哈希校验 · 11 全量过 AtomicFileCommitter · 12 无路径泄漏 · 13 未知 OOXML 保序（diff 只分类不丢弃） · 14 文档内容=不可信数据（策略引擎） · 15 无半写提交 —— 全部有对应断言。

**PERF**：01/02/03 preview/open 零写路径成本 · 04/05/06 并发去重与取消 · 07 强哈希不阻首屏 · 08/10 Actor 控制专用 + 前台优先 · 11 缓存准入+逐出 · 12 dispose 全释放 · 13 有界工作集 · 15 同文件共享上下文 —— 均有测试或基准判定。

## 基准与证据的边界（如实声明）

`npm run bench` 的 5 项判定是**基础设施绿灯**（内存有界、TTFP/TTE 分布、有界视口），
阈值是 PoC 级而非 release gate：Editor 内存基线不代表最终 GenOffice 视觉编辑器的
真实占用（canvas/GPU/字体不在 JS heap），Editor 集成后需重新校准阈值并纳入 CI。

## 引擎交互备忘（实测 officecli 1.0.148）

- `close` 的 flush 字节与 `save` 不同（终写语义）→ 权威哈希在 save+close 之后采样（`AgentRuntime.flushCandidate`）。
- 读命令可能异步拉起 resident 守护进程持锁 → 读后补 `close`（no-op 快）；删除/改名带 EPERM/EBUSY 重试 + 锁释放兜底。
- Windows npm shim 的 JS 入口解析（`%dp0%` 展开）避免 `.cmd` spawn 限制与引号转义。

## 生产加固清单（第二轮）

- **恢复解决**（§77）：`service.resolveRecoveredSession()` — `recovery-required` 会话按文件系统哈希事实恢复 `ready` 或落 `conflict`（含持久事件），非纯标记。
- **外部变异监听接线**（§73）：`openSession` 即启动 `SourceWatcher`，外部写入 → `session.conflict` 事件；自写事件经 `SelfWriteGuard` 抑制。
- **跨会话提交串行化**（P11）：同源文件的 commit 按物理路径互斥 —— 多会话并发 accept 恰好一个落地（有并发测试证据）。
- **MCP stdio 服务器可执行**：`office-mcp` bin（`dist/mcp/launch.js`），协议层 8 项线测试（initialize/ping/tools/list/tools/call/未知方法 -32601/坏 JSON -32700/isError 语义）。


## 渲染宿主方案（「需要 Electron」的解法）

设计 P7 要求的是**一个渲染进程容器**，Electron 只是选项之一。本项目落地了三层：

1. **零宿主（已就绪）**：pptx-render 的 RenderTree（纯数据绘制列表）→ 自有 SVG 适配器 → `sharp` 栅格化 PNG——缩略图与 L5 变更范围像素验证全程无 DOM/无浏览器。
2. **localhost Web 宿主（已就绪）**：`npm run host -- <files...>` 启动 `PreviewHost`——任意浏览器打开即为渲染进程：画廊页、SVG 幻灯片翻页、PNG 缩略、Rust sidecar 支撑的表格视口、JSON API；响应零物理路径（INV-12）。
3. **Electron 包装器（随取随用）**：`tools/host/electron.mjs`——BrowserWindow 指向同一 localhost origin，约 60 行桌面壳，对应 §7 进程模型（main/renderer/runtime 三进程）。`npm i -D electron && npx electron tools/host/electron.mjs <files...>`。交互编辑画布（Univer/Konva）在该容器内按 GenOffice app 层接入。

## 边界与后续（§149 P2 / §146）

- **GenOffice 引擎已全量 vendor 接入（§146）**：`vendor/genoffice`（genspark-ai/genoffice@d35d770，Apache-2.0，引擎源码零修改）。
  - **pptx/docx**：解析/预览走真实引擎数据模型（deck / blocks）；**pptx 视觉预览为无头 SVG**——pptx-render 的 RenderTree（纯数据绘制列表）经自有适配器序列化为独立 SVG 字符串（零 DOM/零 canvas，Node 与 utility process 皆可产出），预览模型携带 `svgSlides`。
  - **xlsx（§30 字面实现）**：Rust sidecar（calamine + IronCalc）已用 cargo 编译（`npm run sidecar:build`），stdio JSON 协议客户端接入；100k 行工作簿的有界视口读取经 sidecar 完成，工作簿内存留在 sidecar 进程（§31）。未编译时自动回退进程内流式解析。
- **WPS 宿主适配器已实现（§86–§88）**：`WpsHostAdapter` 经 KWPP COM（PowerShell 桥，零新依赖）提供宿主真实渲染（slides→PNG）与一次性认证副本（copy→打开→SaveCopyAs→二次渲染；重序列化≠修复，损坏/拒开才 FAIL）。**验证阶梯实测可达 `consumer-certified`**。COM 为单用途服务器：套件与全量并行会互斥，故按 §86 Optional 定位单独回归：`npx vitest run test/integration/wps-host.test.ts`（5 用例）。PowerPoint 为残留 COM 注册（服务器启动失败），如实 unavailable。
- Univer/Konva 交互编辑画布：SVG/PNG/表格视口的只读面已通过 Web 宿主交付；像素级交互编辑 UI 属 GenOffice app 层，在上述宿主容器内接入（见「渲染宿主方案」）。

## 维护流程（main required CI gate）

`main` 受 Repository Ruleset `main-required-ci-gate` 保护：所有变更（包括
定时修复自动化）一律走 **branch → PR → `linux` + `windows` required
checks 全绿 → merge**；直接 push、force-push 与分支删除均被平台拒绝，
无任何 bypass 身份。临时 break-glass 需另行创建专用身份并显式授权。
