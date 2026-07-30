import { mkdirSync, readdirSync, readFileSync, statSync, writeFileSync } from 'fs';
import { join, relative, resolve } from 'path';
import { config } from '../config.js';
import { gql } from './client.js';
import type {
  AssetExportJobResponse,
  DecodeFile,
  DfmImageMap,
  DfmResultEntry,
  DfmResults,
  DfmViolationImage,
  ExportProductResponse,
  PipelineResult,
  PriceResponse,
} from './types.js';

// Field selection for the exportProduct query. This requests the full
// public export payload (specs, BOM, costing scenarios); trim it if you
// only need a subset. The authoritative shape is the Decode API's
// published GraphQL schema.
const EXPORT_PRODUCT_FIELDS = `
    jobId
    status
    warnings { code message }
    data {
        meta { schemaVersion generatedAt supplierName }
        kind
        productId
        partNumber
        specs {
            industrySector
            ipcClassIII
            bareBoard {
                customerSupplied
                boardThickness { value unit }
                numOfLayers
                boardWidth { value unit }
                boardLength { value unit }
                copperThickness { value unit }
                layersCuThickness { value unit }
                material materialOtherText
                finishing finishingOtherText
                hdi controlledDepthDrilling goldFingers customStackup viaPlugging
                solderMaskColor solderMaskColorOtherText
                silkscreenColor silkscreenColorOtherText
                layerFilesNotRequired electricalTesting controlledImpedance blindAndBuriedVias
                panel {
                    panelWidth { value unit }
                    panelLength { value unit }
                    designFilesArePanel
                    boardsPerAsmPanel
                }
            }
            assemblyOptions {
                conformalCoating functionalTesting xRayInspection firstArticleInspection
                bedOfNailsTesting inBoardProgramming automatedOpticalInspection other otherText
            }
            projectBom {
                id partNumber pricedPartNumber partManufacturer
                internalPartNumber customerPartNumber mfrIdentifier customerIdentifier
                mountType description notes potentialPackage sourcingOption substitutions
                referenceDesignators { designator dnp }
                quantityPerBoard customerSupplied
                lifeCycleStatus rohsStatus reachStatus
                lifeCycleStatusRaw rohsStatusRaw reachStatusRaw
                rohsCompliant reachCompliant
                countryOfOrigin minimumOrderQuantity itemClass leadCount lastPricingUpdate
            }
            stackup {
                type fileType typeIndex circuitBoardLayerNumber
                copperThickness { value unit }
                dielectricType
                dielectricThickness { value unit }
                plated drillSpanStart drillSpanEnd
            }
            padTraceData {
                layers {
                    layerName fileName circularPadCount rectangularPadCount totalPadCount
                    minPadSurfaceArea { value unit }
                    maxPadSurfaceArea { value unit }
                    totalPadSurfaceArea { value unit }
                    traceCount
                    minTraceWidth { value unit }
                    maxTraceWidth { value unit }
                    uniqueTraceWidths
                    minSpacing { value unit }
                    copperArea { value unit }
                }
                totalPadCount totalTraceCount
                minSpacing { value unit }
                totalCopperArea { value unit }
            }
            drillData {
                drillFiles {
                    drillDescription fileName holeCount
                    minHoleDiameter { value unit }
                    maxHoleDiameter { value unit }
                    totalHoleSurfaceArea { value unit }
                }
                totalHoleCount
                holeDensity { value unit }
            }
            fileNotes { text cleanedText foundIn { fileName } }
            fileTables { table foundIn { fileName } }
            provenance {
                areas { area entries { source message sourceFile } }
                totalDecisions
            }
        }
        costing {
            scenarios {
                quantity
                assemblyUnitPrice { value formatted currency }
                pricingLineItems {
                    itemId group name
                    unitPrice { value formatted currency }
                    total { value formatted currency }
                    specs {
                        partNumber distributor clickUrl stock leadTime
                        tariffCost { value formatted currency }
                        additionalDetails {
                            source_name source_id id purchase_quant
                            is_backordered is_uncertain_lead
                            best_case_lead expected_ships_in expected_transit_time
                            unit_price { value formatted currency }
                            total_price { value formatted currency }
                            landed_price { value formatted currency }
                            part_quote_id returnable_until release_deadline package
                            availability { availability_type stock lead multiple moq }
                        }
                    }
                }
                calculations { name floatValue booleanValue }
                profitProjections {
                    type margin
                    grossProfit { value formatted currency }
                }
                validationDetails {
                    assemblyWithSubstitutions { primaryPartNumber subPartNumber refDes }
                }
                subtotal unitPrice pcbSubtotal assemblySubtotal bomSubtotal
            }
        }
    }
    errors {
        ... on IError { code message }
        ... on ErrorExportProductNotFound { code message }
        ... on ErrorJobNotFound { code message }
        ... on ErrorAccessDenied { code message }
    }
`;

export function isArchiveFile(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return lower.endsWith('.zip') || lower.endsWith('.tar') || lower.endsWith('.tar.gz') || lower.endsWith('.tgz');
}

export function classifyFile(fileName: string): DecodeFile['type'] {
  if (isArchiveFile(fileName)) return 'Archive';
  const isBomName = /bom/i.test(fileName) || /bill[\s_-]?of[\s_-]?materials?/i.test(fileName);
  if (isBomName && /\.(xlsx|xls|csv)$/i.test(fileName)) return 'Bill Of Material';
  if (/\.(pdf|rpt|doc|docx|md)$/i.test(fileName)) return 'Other Attachment';
  return 'Gerber/Drill File';
}

function isMacOsJunk(name: string): boolean {
  // AppleDouble resource forks ("._foo.gbr") and DS_Store. The `__MACOSX`
  // directory itself is excluded one level up in the directory walk.
  return name.startsWith('._') || name === '.DS_Store';
}

export function scanDirectory(dir: string): DecodeFile[] {
  const rootDir = resolve(dir);
  const files: DecodeFile[] = [];
  const seenNames = new Map<string, DecodeFile>();

  const walk = (current: string) => {
    for (const entry of readdirSync(current)) {
      if (entry.startsWith('.')) continue;
      if (entry === '__MACOSX') continue;
      if (isMacOsJunk(entry)) continue;
      const full = join(current, entry);
      const st = statSync(full);
      if (st.isDirectory()) { walk(full); continue; }
      if (!st.isFile() || st.size === 0) continue;
      // Archives (.zip/.tar/.tar.gz/.tgz) are uploaded as-is. Decode
      // extracts them server-side and classifies the inner files.
      if (seenNames.has(entry)) {
        const existing = seenNames.get(entry)!;
        console.warn(`  WARN: duplicate filename '${entry}' (${relative(rootDir, full)}) — keeping '${relative(rootDir, existing.absPath)}', skipping this one.`);
        continue;
      }
      const record: DecodeFile = {
        absPath: full,
        name: entry,
        type: classifyFile(entry),
        size: st.size,
      };
      seenNames.set(entry, record);
      files.push(record);
    }
  };

  walk(rootDir);
  return files;
}

async function createProject(name: string): Promise<string> {
  const data = await gql<{ createProject: { project?: { _id: string }; errors?: any[] } }>(
    `mutation createProject($name: String, $designerEmail: String) {
      createProject(name: $name, designerEmail: $designerEmail) {
        project { _id }
        errors { ... on IError { code message } ... on ErrorProjectNotCreated { code message } }
      }
    }`,
    { name, designerEmail: config.outlook.userEmail },
  );
  const id = data.createProject.project?._id;
  if (!id) throw new Error(`createProject failed: ${JSON.stringify(data.createProject.errors)}`);
  return id;
}

type Policy = { fileName: string; filePath?: string; policy: { url: string; fields: [string, string][] } };

async function getUploadPolicies(projectId: string, files: DecodeFile[]): Promise<{ requestId: string; policies: Policy[] }> {
  const data = await gql<{ getFileUploadPolicies: { writePolicies?: { requestId: string; policies: Policy[] }; errors?: any[] } }>(
    `query getFileUploadPolicies($projectId: String!, $files: [ProjectFile!]!) {
      getFileUploadPolicies(projectId: $projectId, files: $files) {
        writePolicies { requestId policies { fileName filePath policy { url fields } } }
        errors { ... on IError { code message } }
      }
    }`,
    // Omit `type` for Archive uploads — the docs explicitly say "omit
    // type and the server will classify each [extracted] file."
    { projectId, files: files.map(f => f.type === 'Archive' ? { path: f.name } : { path: f.name, type: f.type }) },
  );
  const r = data.getFileUploadPolicies;
  if (r.errors?.length) throw new Error(`getFileUploadPolicies errors: ${JSON.stringify(r.errors)}`);
  if (!r.writePolicies?.policies?.length) throw new Error('No upload policies returned');
  return r.writePolicies;
}

async function uploadFileToS3(policy: Policy, fileData: Buffer): Promise<void> {
  const formData = new FormData();
  for (const [field, value] of policy.policy.fields) formData.append(field, value);
  formData.append('file', new Blob([new Uint8Array(fileData)]), policy.fileName);
  const res = await fetch(policy.policy.url, { method: 'POST', body: formData });
  if (!res.ok && res.status !== 204) {
    const text = await res.text().catch(() => '');
    throw new Error(`S3 upload failed for ${policy.fileName} (${res.status}): ${text}`);
  }
}

async function analyzeProject(projectId: string): Promise<{ analysis?: { complete?: boolean }; errors?: any[] }> {
  const data = await gql<{ analyzeProject: { analysis?: { complete?: boolean }; errors?: any[] } }>(
    `mutation analyzeProject($projectId: String!) {
      analyzeProject(projectId: $projectId) {
        analysis { complete }
        errors { ... on IError { code message } ... on ErrorProjectNotFound { code message } ... on ErrorLimitExceeded { code message } }
      }
    }`,
    { projectId },
  );
  return data.analyzeProject;
}

async function getProject(projectId: string): Promise<{ _id: string; name?: string; boardFileType?: string } | null> {
  const data = await gql<{ getProject: { project?: any; errors?: any[] } }>(
    `query getProject($projectId: String!) {
      getProject(projectId: $projectId) {
        project { _id name boardFileType }
        errors { ... on IError { code message } ... on ErrorProjectNotFound { code message } }
      }
    }`,
    { projectId },
  );
  return data.getProject.project ?? null;
}

async function exportProduct(productId: string, quantity: number): Promise<ExportProductResponse> {
  const data = await gql<{ exportProduct: ExportProductResponse }>(
    `query exportProduct($productId: String!, $quantity: Int) {
      exportProduct(productId: $productId, quantity: $quantity) { ${EXPORT_PRODUCT_FIELDS} }
    }`,
    { productId, quantity },
  );
  return data.exportProduct;
}

async function getExportProductJob(jobId: string): Promise<ExportProductResponse> {
  const data = await gql<{ getExportProductJob: ExportProductResponse }>(
    `query getExportProductJob($jobId: String!) {
      getExportProductJob(jobId: $jobId) { ${EXPORT_PRODUCT_FIELDS} }
    }`,
    { jobId },
  );
  return data.getExportProductJob;
}

// ----------------------------------------------------------------------
// SpeedDFM (getDFM + retrieveDFM). Same async/poll pattern as
// exportProduct.
//
// The demo renders its OWN DFM report (src/reports/dfm-pdf.ts) from the
// `results` + `summary` payload rather than selecting `reportUrl` (the
// API's server-rendered PDF). Per-violation image tiles for that report
// are fetched separately and on demand via the top-level
// `dfmEntryImage(productId, x, y)` query — see fetchDfmViolationImages
// below — which avoids paying for a render of every marker up front.
//
// The error union carries the full set of members getDFM can surface;
// retrieveDFM uses a subset, but selecting extra union members is
// harmless, so DFM_FIELDS carries the full set for both.
// ----------------------------------------------------------------------

const DFM_FIELDS = `
    jobId
    status
    data {
        results {
            ruleId
            ruleName
            ruleType
            severity
            layerName
            layerType
            x
            y
            measuredValue
            threshold
            unit
            description
        }
        summary {
            fabricationErrorCount
            fabricationWarningCount
            assemblyErrorCount
            assemblyWarningCount
            generalCount
            analysisStatus
        }
    }
    errors {
        ... on IError { code message }
        ... on ErrorAnalysisNotComplete { code message }
        ... on ErrorDfmFailed { code message }
        ... on ErrorDfmNotAvailable { code message }
        ... on ErrorExportProductNotFound { code message }
        ... on ErrorProjectNotFound { code message }
        ... on ErrorExportFailed { code message }
        ... on ErrorAccessDenied { code message }
        ... on ErrorJobNotFound { code message }
    }
`;

type DfmJobResponse = {
  jobId?: string;
  status?: string;
  data?: DfmResults | null;
  errors?: { code?: string; message?: string }[];
};

async function getDfm(productId: string): Promise<DfmJobResponse> {
  const data = await gql<{ getDFM: DfmJobResponse }>(
    `query getDFM($productId: String!) {
      getDFM(productId: $productId) { ${DFM_FIELDS} }
    }`,
    { productId },
  );
  return data.getDFM;
}

async function retrieveDfm(jobId: string): Promise<DfmJobResponse> {
  const data = await gql<{ retrieveDFM: DfmJobResponse }>(
    `query retrieveDFM($jobId: String!) {
      retrieveDFM(jobId: $jobId) { ${DFM_FIELDS} }
    }`,
    { jobId },
  );
  return data.retrieveDFM;
}

async function retrieveAsset(jobId: string): Promise<AssetExportJobResponse> {
  const data = await gql<{ retrieveAsset: AssetExportJobResponse }>(
    `query retrieveAsset($jobId: String!) {
      retrieveAsset(jobId: $jobId) {
        jobId status url contentType
        errors {
          ... on IError { code message }
          ... on ErrorAssetFailed { code message }
          ... on ErrorAssetNotAvailable { code message }
          ... on ErrorJobNotFound { code message }
          ... on ErrorAccessDenied { code message }
        }
      }
    }`,
    { jobId },
  );
  return data.retrieveAsset;
}

async function getProjectPrices(projectId: string, quantity: number): Promise<PriceResponse[]> {
  const data = await gql<{ getProjectPrices: PriceResponse[] }>(
    `query getProjectPrices($projectId: String!, $quantity: Int) {
      getProjectPrices(projectId: $projectId, quantity: $quantity) {
        mfrId
        mfrDisplayName
        price { unitPrice subtotal quantity currency projectType }
        errors {
          ... on IError { code message }
          ... on ErrorProjectNotFound { code message }
          ... on ErrorCannotPrice { code message }
          ... on ErrorQuoteRequired { code message }
          ... on ErrorPricingDisabled { code message }
          ... on ErrorPricingInProgress { code message }
          ... on ErrorMfrNotFound { code message }
        }
      }
    }`,
    { projectId, quantity },
  );
  return data.getProjectPrices ?? [];
}

// Orchestrates DFM end-to-end. Returns null when DFM isn't available
// (e.g. ErrorDfmNotAvailable when the API plan behind the key doesn't
// include SpeedDFM). Logs errors but does NOT throw — DFM is
// informational; a missing DFM shouldn't kill the whole email reply.
async function runDfmAnalysis(productId: string): Promise<DfmResults | null> {
  let response: DfmJobResponse;
  try {
    response = await getDfm(productId);
  } catch (err) {
    console.warn(`  WARN: getDFM call failed: ${(err as Error).message}`);
    return null;
  }
  console.log(`  getDFM: status=${response.status}${response.jobId ? ', jobId=' + response.jobId : ''}`);

  if (response.errors?.length) {
    console.warn(`  WARN: getDFM returned errors: ${JSON.stringify(response.errors)}`);
    return null;
  }

  if (response.status === 'processing' && response.jobId) {
    const jobId = response.jobId;
    const MAX = 60;
    for (let i = 1; i <= MAX; i++) {
      await sleep(5000);
      response = await retrieveDfm(jobId);
      if (i === 1 || i % 5 === 0 || response.status === 'completed' || response.status === 'failed') {
        console.log(`  DFM poll ${i}: status=${response.status}`);
      }
      if (response.status === 'completed' || response.status === 'failed') break;
    }
  }

  if (response.status !== 'completed' || !response.data) {
    console.warn(`  WARN: DFM did not complete (status=${response.status}); errors=${JSON.stringify(response.errors)}`);
    return null;
  }
  return response.data;
}

// Display ordering shared by the DFM report and the image-tile picker:
// errors before warnings before info; within a severity, fabrication
// before assembly before general.
const DFM_SEVERITY_ORDER: Record<string, number> = { error: 0, warning: 1, info: 2 };
const DFM_TYPE_ORDER: Record<string, number> = { fabrication: 0, assembly: 1, general: 2 };

export function compareDfmEntries(a: DfmResultEntry, b: DfmResultEntry): number {
  const sa = DFM_SEVERITY_ORDER[a.severity] ?? 3;
  const sb = DFM_SEVERITY_ORDER[b.severity] ?? 3;
  if (sa !== sb) return sa - sb;
  return (DFM_TYPE_ORDER[a.ruleType] ?? 3) - (DFM_TYPE_ORDER[b.ruleType] ?? 3);
}

// Kicks off (or reuses) a composite-image render of one DFM violation,
// identified by (x, y). Returns the asset-export job id to poll, or null
// when the coordinates are invalid / DFM isn't enabled for this vendor.
async function getDfmEntryImage(productId: string, x: number, y: number): Promise<string | null> {
  const data = await gql<{ dfmEntryImage: { assetExportJobId: string } | null }>(
    `query dfmEntryImage($productId: String!, $x: Float!, $y: Float!) {
      dfmEntryImage(productId: $productId, x: $x, y: $y) { assetExportJobId }
    }`,
    { productId, x, y },
  );
  return data.dfmEntryImage?.assetExportJobId ?? null;
}

// Polls retrieveAsset until the presigned URL lands (or the job fails).
async function pollAssetUntilReady(jobId: string): Promise<AssetExportJobResponse | null> {
  const MAX = 36;
  let asset: AssetExportJobResponse | null = null;
  for (let i = 1; i <= MAX; i++) {
    asset = await retrieveAsset(jobId);
    if (asset.status === 'completed' || asset.status === 'failed') break;
    await sleep(2500);
  }
  if (!asset || asset.status !== 'completed' || !asset.url) return null;
  return asset;
}

// Materializes per-violation image tiles for the DFM report. Renders at
// most `limit` tiles (highest-severity first), each via
// dfmEntryImage → retrieveAsset → download. Writes the PNGs to
// `<outputDir>/dfm-images/` as artifacts and returns a map keyed by the
// DfmResultEntry object reference so the report can look each tile up.
// Best-effort: a failed tile is logged and skipped, never thrown.
async function fetchDfmViolationImages(
  productId: string,
  dfm: DfmResults,
  opts: { limit: number; outputDir: string },
): Promise<DfmImageMap> {
  const map: DfmImageMap = new Map();
  if (opts.limit <= 0) return map;

  const withCoords = (dfm.results || []).filter(e => e.x != null && e.y != null);
  if (!withCoords.length) {
    console.log('  No DFM violations carry coordinates — skipping image tiles.');
    return map;
  }
  const selected = [...withCoords].sort(compareDfmEntries).slice(0, opts.limit);
  console.log(`  Rendering ${selected.length} DFM violation tile(s) (of ${withCoords.length} with coordinates)...`);

  const imagesDir = resolve(opts.outputDir, 'dfm-images');
  mkdirSync(imagesDir, { recursive: true });

  await Promise.all(selected.map(async (entry, i) => {
    const tag = `#${i + 1} (${entry.ruleId})`;
    try {
      const jobId = await getDfmEntryImage(productId, entry.x as number, entry.y as number);
      if (!jobId) { console.warn(`  WARN: DFM image ${tag}: no asset job returned.`); return; }
      const asset = await pollAssetUntilReady(jobId);
      if (!asset?.url) { console.warn(`  WARN: DFM image ${tag}: asset not ready.`); return; }
      const res = await fetch(asset.url);
      if (!res.ok) { console.warn(`  WARN: DFM image ${tag}: download failed (${res.status}).`); return; }
      const buf = Buffer.from(await res.arrayBuffer());
      const contentType = asset.contentType || 'image/png';
      const ext = contentType.includes('png') ? 'png' : contentType.includes('jpeg') ? 'jpg' : 'img';
      const path = resolve(imagesDir, `violation-${String(i + 1).padStart(2, '0')}.${ext}`);
      writeFileSync(path, buf);
      const img: DfmViolationImage = {
        path,
        dataUri: `data:${contentType};base64,${buf.toString('base64')}`,
        contentType,
      };
      map.set(entry, img);
    } catch (err) {
      console.warn(`  WARN: DFM image ${tag} failed: ${(err as Error).message}`);
    }
  }));
  console.log(`  Rendered ${map.size}/${selected.length} DFM violation tile(s) to ${imagesDir}`);
  return map;
}

const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));

export async function runDecodePipeline(opts: {
  inputDir: string;
  projectName: string;
  quantity: number;
  // Output directory for assets we materialize during the pipeline
  // (the DFM violation image tiles). The orchestrator passes its
  // per-message `out/` folder so artifacts land next to the reports.
  outputDir: string;
  // Max number of per-violation DFM image tiles to render. Defaults to
  // 20 when omitted. 0 disables image rendering entirely.
  imageLimit?: number;
}): Promise<PipelineResult> {
  const files = scanDirectory(opts.inputDir);
  if (!files.length) throw new Error(`No usable files found in ${opts.inputDir}`);

  const countsByType: Record<string, number> = {};
  for (const f of files) countsByType[f.type] = (countsByType[f.type] || 0) + 1;
  console.log(`  Files: ${files.length} (${Object.entries(countsByType).map(([k, v]) => `${k}: ${v}`).join(', ')})`);

  const projectId = await createProject(opts.projectName);
  console.log(`  Project ID: ${projectId}`);

  const { policies } = await getUploadPolicies(projectId, files);
  await Promise.all(policies.map(async p => {
    const match = files.find(f => f.name === p.fileName);
    if (!match) throw new Error(`No local file matches policy fileName=${p.fileName}`);
    const data = readFileSync(match.absPath);
    await uploadFileToS3(p, data);
  }));
  console.log(`  Uploaded ${policies.length} file(s) to S3`);
  await sleep(1000);

  const MAX_POLLS = 260;
  const POLL_INTERVAL_MS = 3000;
  const startMs = Date.now();
  let complete = false;
  for (let i = 1; i <= MAX_POLLS; i++) {
    const result = await analyzeProject(projectId);
    const errs = result?.errors?.length ? result.errors : [];
    const project = await getProject(projectId);
    const elapsed = ((Date.now() - startMs) / 1000).toFixed(1);
    const parts = [`t+${elapsed}s`, `complete=${result?.analysis?.complete}`, `boardFileType=${project?.boardFileType ?? 'n/a'}`];
    if (errs.length) parts.push(`errors=${JSON.stringify(errs)}`);
    if (i === 1 || i % 5 === 0 || result?.analysis?.complete) console.log(`  Analyze poll ${i}: ${parts.join(', ')}`);
    if (result?.analysis?.complete) { complete = true; break; }
    await sleep(POLL_INTERVAL_MS);
  }
  if (!complete) {
    throw new Error(`Analysis did not complete within ${(MAX_POLLS * POLL_INTERVAL_MS / 1000).toFixed(0)}s for project ${projectId}`);
  }
  console.log(`  Analysis complete in ${((Date.now() - startMs) / 1000).toFixed(1)}s`);

  let result = await exportProduct(projectId, opts.quantity);
  console.log(`  exportProduct: status=${result.status}${result.jobId ? ', jobId=' + result.jobId : ''}`);
  if (result.status === 'processing' && result.jobId) {
    const jobId = result.jobId;
    const MAX = 60;
    for (let i = 1; i <= MAX; i++) {
      await sleep(5000);
      result = await getExportProductJob(jobId);
      if (i === 1 || i % 5 === 0 || result.status === 'completed' || result.status === 'failed') {
        console.log(`  Export poll ${i}: status=${result.status}`);
      }
      if (result.status === 'completed' || result.status === 'failed') break;
    }
  }

  if (result.status !== 'completed') {
    throw new Error(`exportProduct did not complete (status=${result.status}); errors=${JSON.stringify(result.errors)}`);
  }
  if (!result.data) throw new Error('exportProduct returned status=completed but no data');

  const exportData = result.data;

  const project = await getProject(projectId);
  const productName = project?.name || opts.projectName;

  // DFM always runs. Failures are logged and produce a null DFM
  // section rather than killing the email.
  console.log('  Running SpeedDFM analysis...');
  const dfm = await runDfmAnalysis(projectId);
  if (dfm) {
    const s = dfm.summary;
    console.log(`  DFM complete: ${s.fabricationErrorCount}/${s.fabricationWarningCount} fab err/warn, ${s.assemblyErrorCount}/${s.assemblyWarningCount} asm err/warn, ${s.generalCount} general (${s.analysisStatus})`);
  }

  // Per-violation image tiles for the demo's own DFM report (the report
  // PDF itself is rendered by the orchestrator, src/pipeline.ts, next to
  // the BOM Excel and Exec PDF). dfmEntryImage → retrieveAsset, capped to
  // imageLimit and highest-severity-first.
  const dfmImages: DfmImageMap = dfm
    ? await fetchDfmViolationImages(projectId, dfm, { limit: opts.imageLimit ?? 20, outputDir: opts.outputDir })
    : new Map();

  // Cross-manufacturer pricing. Errors per-manufacturer are kept in the
  // response so the report can show which ones declined; only a hard
  // network failure is suppressed here.
  console.log('  Fetching manufacturer pricing comparison...');
  let prices: PriceResponse[] = [];
  try {
    prices = await getProjectPrices(projectId, opts.quantity);
    const okCount = prices.filter(p => !p.errors?.length && p.price).length;
    console.log(`  getProjectPrices returned ${prices.length} manufacturer(s) (${okCount} priced).`);
  } catch (err) {
    console.warn(`  WARN: getProjectPrices failed: ${(err as Error).message}`);
  }

  return {
    projectId,
    productName,
    quantity: opts.quantity,
    exportData,
    dfm,
    dfmImages,
    prices,
    files,
  };
}
