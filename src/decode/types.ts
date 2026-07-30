export type DecodeFile = {
  absPath: string;
  name: string;
  // 'Archive' is the demo-internal marker for .zip/.tar/.tar.gz/.tgz
  // uploads. We send `type: undefined` to Decode for these so its
  // server-side extractor classifies the contained files itself.
  type: 'Bill Of Material' | 'Other Attachment' | 'Gerber/Drill File' | 'Archive';
  size: number;
};

// Shape of the Decode API's public exportProduct payload. The
// authoritative definitions are the API's published GraphQL schema;
// these types only need to cover the fields the demo selects in
// EXPORT_PRODUCT_FIELDS (src/decode/pipeline.ts).

export type Money = {
  value: number | null;
  formatted: string;
  currency: string;
};

export type ValueWithUnit = {
  value: number | null;
  unit: string;
};

export type TextWithUnit = {
  value: string | null;
  unit: string;
};

export type RefDesignator = {
  designator: string;
  dnp: boolean;
};

export type ExportMeta = {
  schemaVersion: string;
  generatedAt: string;
  supplierName?: string | null;
};

export type ExportAssemblyOptions = {
  conformalCoating?: boolean;
  functionalTesting?: boolean;
  xRayInspection?: boolean;
  firstArticleInspection?: boolean;
  bedOfNailsTesting?: boolean;
  inBoardProgramming?: boolean;
  automatedOpticalInspection?: boolean;
  other?: boolean;
  otherText?: string | null;
};

export type ExportPanel = {
  panelWidth?: ValueWithUnit | null;
  panelLength?: ValueWithUnit | null;
  designFilesArePanel?: boolean | null;
  boardsPerAsmPanel?: number | null;
};

export type ExportBareBoard = {
  customerSupplied?: boolean | null;
  boardThickness?: ValueWithUnit | null;
  numOfLayers?: number | null;
  boardWidth?: ValueWithUnit | null;
  boardLength?: ValueWithUnit | null;
  copperThickness?: ValueWithUnit | null;
  layersCuThickness?: TextWithUnit | null;
  material?: string | null;
  materialOtherText?: string | null;
  finishing?: string | null;
  finishingOtherText?: string | null;
  hdi?: boolean | null;
  controlledDepthDrilling?: boolean | null;
  goldFingers?: boolean | null;
  customStackup?: boolean | null;
  viaPlugging?: string | null;
  solderMaskColor?: string | null;
  solderMaskColorOtherText?: string | null;
  silkscreenColor?: string | null;
  silkscreenColorOtherText?: string | null;
  layerFilesNotRequired?: string[] | null;
  electricalTesting?: boolean | null;
  controlledImpedance?: boolean | null;
  blindAndBuriedVias?: boolean | null;
  panel?: ExportPanel | null;
};

export type ExportProjectBomItem = {
  id?: string | null;
  partNumber: string;
  pricedPartNumber?: string | null;
  partManufacturer?: string | null;
  internalPartNumber?: string | null;
  customerPartNumber?: string | null;
  mfrIdentifier?: string | null;
  customerIdentifier?: string | null;
  mountType?: string | null;
  description: string;
  notes: string[];
  potentialPackage?: string | null;
  sourcingOption?: string | null;
  substitutions?: string | null;
  referenceDesignators: RefDesignator[];
  quantityPerBoard: number;
  customerSupplied: boolean;
  lifeCycleStatus?: string | null;
  rohsStatus?: string | null;
  reachStatus?: string | null;
  // The "Raw" fields carry the distributor's unnormalized status string;
  // the "*Compliant" booleans are Decode's normalized verdict; the rest
  // are sourcing/classification metadata. All optional — older projects
  // won't have them populated.
  lifeCycleStatusRaw?: string | null;
  rohsStatusRaw?: string | null;
  reachStatusRaw?: string | null;
  rohsCompliant?: boolean | null;
  reachCompliant?: boolean | null;
  countryOfOrigin?: string | null;
  minimumOrderQuantity?: number | null;
  itemClass?: string | null;
  leadCount?: number | null;
  lastPricingUpdate?: string | null;
};

export type ExportStackupLayer = {
  type: string;
  fileType: string;
  typeIndex: number;
  circuitBoardLayerNumber?: number | null;
  copperThickness?: ValueWithUnit | null;
  dielectricType?: string | null;
  dielectricThickness?: ValueWithUnit | null;
  plated?: boolean | null;
  drillSpanStart?: number | null;
  drillSpanEnd?: number | null;
};

export type ExportLayerPadTraceSummary = {
  layerName: string;
  fileName: string;
  circularPadCount: number;
  rectangularPadCount: number;
  totalPadCount: number;
  minPadSurfaceArea?: ValueWithUnit | null;
  maxPadSurfaceArea?: ValueWithUnit | null;
  totalPadSurfaceArea?: ValueWithUnit | null;
  traceCount: number;
  minTraceWidth?: ValueWithUnit | null;
  maxTraceWidth?: ValueWithUnit | null;
  uniqueTraceWidths: number;
  minSpacing?: ValueWithUnit | null;
  copperArea?: ValueWithUnit | null;
};

export type ExportPadTraceData = {
  layers: ExportLayerPadTraceSummary[];
  totalPadCount: number;
  totalTraceCount: number;
  minSpacing?: ValueWithUnit | null;
  totalCopperArea?: ValueWithUnit | null;
};

export type ExportDrillFileSummary = {
  drillDescription: string;
  fileName: string;
  holeCount: number;
  minHoleDiameter?: ValueWithUnit | null;
  maxHoleDiameter?: ValueWithUnit | null;
  totalHoleSurfaceArea?: ValueWithUnit | null;
};

export type ExportDrillData = {
  drillFiles: ExportDrillFileSummary[];
  totalHoleCount: number;
  holeDensity?: ValueWithUnit | null;
};

export type ExportFileNote = {
  text: string;
  cleanedText?: string | null;
  foundIn: { fileName: string }[];
};

export type ExportFileTable = {
  table: string[][];
  foundIn: { fileName: string }[];
};

export type ExportProvenanceEntry = {
  source: string;
  message: string;
  sourceFile?: string | null;
};

export type ExportProvenanceArea = {
  area: string;
  entries: ExportProvenanceEntry[];
};

export type ExportProvenance = {
  areas: ExportProvenanceArea[];
  totalDecisions: number;
};

export type ExportProductSpecs = {
  industrySector?: string | null;
  ipcClassIII?: boolean | null;
  bareBoard?: ExportBareBoard | null;
  assemblyOptions?: ExportAssemblyOptions | null;
  projectBom?: ExportProjectBomItem[] | null;
  stackup?: ExportStackupLayer[] | null;
  padTraceData?: ExportPadTraceData | null;
  drillData?: ExportDrillData | null;
  fileNotes?: ExportFileNote[] | null;
  fileTables?: ExportFileTable[] | null;
  provenance?: ExportProvenance | null;
};

// Nested per-source availability inside ExportBomAdditionalDetails.
// Populated by Decode when a row has distributor quotes broken out
// per source rather than rolled up to the top-level specs fields.
export type ExportBomAvailability = {
  availability_type?: string | null;
  stock?: number | null;
  lead?: number | null;
  multiple?: number | null;
  moq?: number | null;
};

export type ExportBomAdditionalDetails = {
  source_name?: string | null;
  source_id?: string | null;
  id?: string | null;
  purchase_quant?: number | null;
  is_backordered?: boolean | null;
  is_uncertain_lead?: boolean | null;
  best_case_lead?: number | null;
  expected_ships_in?: number | null;
  // Added in the export reshape: per-source money + timing + return-policy
  // fields broken out alongside the existing availability block.
  unit_price?: Money | null;
  total_price?: Money | null;
  landed_price?: Money | null;
  part_quote_id?: string | null;
  expected_transit_time?: number | null;
  returnable_until?: string | null;
  release_deadline?: string | null;
  package?: string | null;
  availability?: ExportBomAvailability | null;
};

// Per-row engineering/sourcing/compliance metadata attached to a
// pricingLineItem whose group === 'BOM'. The demo's reports read the
// top-level rollup fields (stock, distributor, leadTime, etc.) when
// populated, and fall back to additionalDetails[].availability when
// the top-level fields are null but per-source data is present.
export type ExportPricingLineItemSpecs = {
  partNumber?: string | null;
  distributor?: string | null;
  clickUrl?: string | null;
  stock?: number | null;
  leadTime?: number | null;
  tariffCost?: Money | null;
  additionalDetails?: ExportBomAdditionalDetails[] | null;
};

export type ExportPricingLineItem = {
  itemId?: string | null;
  group: string;
  name: string;
  unitPrice?: Money | null;
  total?: Money | null;
  specs?: ExportPricingLineItemSpecs | null;
};

export type ExportPricingCalculation = {
  name: string;
  floatValue?: number | null;
  booleanValue?: boolean | null;
};

export type ExportProfitProjection = {
  type: string;
  margin?: number | null;
  grossProfit?: Money | null;
};

export type ExportAssemblySubstitution = {
  primaryPartNumber: string;
  subPartNumber: string;
  refDes?: string | null;
};

export type ExportValidationDetails = {
  assemblyWithSubstitutions?: ExportAssemblySubstitution[] | null;
};

export type ExportCostingScenario = {
  quantity: number;
  assemblyUnitPrice?: Money | null;
  pricingLineItems: ExportPricingLineItem[];
  calculations: ExportPricingCalculation[];
  profitProjections?: ExportProfitProjection[] | null;
  validationDetails?: ExportValidationDetails | null;
  subtotal?: number | null;
  unitPrice?: number | null;
  pcbSubtotal?: number | null;
  assemblySubtotal?: number | null;
  bomSubtotal?: number | null;
};

export type ExportCosting = {
  scenarios: ExportCostingScenario[];
};

export type ExportProductData = {
  meta: ExportMeta;
  kind: string;
  productId: string;
  partNumber: string;
  specs?: ExportProductSpecs | null;
  costing?: ExportCosting | null;
};

export type ExportProductResponse = {
  jobId?: string;
  status?: 'completed' | 'processing' | 'failed' | string;
  data?: ExportProductData | null;
  errors?: { message?: string; code?: string }[];
};

// ----------------------------------------------------------------------
// SpeedDFM (manufacturing/assembly rule checks).
// ----------------------------------------------------------------------

export type DfmRuleType = 'fabrication' | 'assembly' | 'general' | string;
export type DfmSeverity = 'error' | 'warning' | 'info' | string;
export type DfmAnalysisStatus = 'complete' | 'processing' | 'not_run' | string;

export type AssetExportJobReference = {
  assetExportJobId: string;
};

export type AssetExportJobResponse = {
  jobId: string;
  status: 'completed' | 'processing' | 'failed' | string;
  url?: string | null;
  contentType?: string | null;
  errors?: { code?: string; message?: string }[];
};

export type DfmResultEntry = {
  ruleId: string;
  ruleName: string;
  ruleType: DfmRuleType;
  severity: DfmSeverity;
  layerName?: string | null;
  layerType?: string | null;
  x?: number | null;
  y?: number | null;
  measuredValue?: number | null;
  threshold?: number | null;
  unit: string;
  description: string;
};

export type DfmSummary = {
  fabricationErrorCount: number;
  fabricationWarningCount: number;
  assemblyErrorCount: number;
  assemblyWarningCount: number;
  generalCount: number;
  analysisStatus: DfmAnalysisStatus;
};

export type DfmResults = {
  results: DfmResultEntry[];
  summary: DfmSummary;
  reportUrl?: AssetExportJobReference | null;
};

// A rendered per-violation tile the demo materializes via
// dfmEntryImage → retrieveAsset. `dataUri` is the base64-inlined PNG so
// the DFM report HTML embeds without needing file/network access at
// render time; `path` is the same image written to disk as an artifact.
export type DfmViolationImage = {
  path: string;
  dataUri: string;
  contentType: string;
};

// Keyed by the DfmResultEntry object reference from `dfm.results` so the
// report generator can look up a violation's tile regardless of how it
// re-sorts the list.
export type DfmImageMap = Map<DfmResultEntry, DfmViolationImage>;

// ----------------------------------------------------------------------
// Cross-manufacturer project pricing (getProjectPrices).
// `Price` carries totals only — no per-BOM stock/distributor fields.
// ----------------------------------------------------------------------

export type Price = {
  unitPrice: number;
  subtotal: number;
  quantity: number;
  currency?: string | null;
  projectType?: string | null;
};

export type PriceResponse = {
  mfrId: number;
  mfrDisplayName?: string | null;
  price?: Price | null;
  errors?: { code?: string; message?: string }[];
};

// ----------------------------------------------------------------------
// Aggregated pipeline result.
// ----------------------------------------------------------------------

export type PipelineResult = {
  projectId: string;
  productName: string;
  quantity: number;
  exportData: ExportProductData;
  dfm: DfmResults | null;
  // Per-violation tiles fetched for the DFM report. Empty when DFM is
  // unavailable or no violations carried renderable coordinates. The DFM
  // report PDF itself is generated by the orchestrator (src/pipeline.ts),
  // alongside the BOM Excel and Exec PDF.
  dfmImages: DfmImageMap;
  prices: PriceResponse[];
  files: DecodeFile[];
};
