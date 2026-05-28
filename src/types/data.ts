import type { ToleranceRow } from '../lib/tolerance'

export interface IvPoint {
  volts: number
  amps: number
  watts: number
}

export interface MeasurementSummary {
  pmaxW: number | null
  vmppV: number | null
  imppA: number | null
  vocV: number | null
  iscA: number | null
}

export interface MeasurementRecord {
  sourcePath: string
  relativePath: string
  measurementDate: string | null
  measurementTime: string | null
  measurementNote: string | null
  station: string | null
  arrayLocation: string | null
  systemGroup: string | null
  inverter: string | null
  stringName: string | null
  irradianceWm2: number | null
  cellTemperatureC: number | null
  latitude: number | null
  longitude: number | null
  timeZone: number | null
  moduleManufacturer: string | null
  moduleModel: string | null
  modulesInString: number | null
  stringsInParallel: number | null
  wireGaugeMm2: number | null
  wireLengthM: number | null
  measurementSummary: MeasurementSummary
  ivMeasurements: IvPoint[]
}

export interface DataImportError {
  filePath: string
  message: string
}

export interface DataImportResult {
  folderPath: string
  totalFiles: number
  records: MeasurementRecord[]
  errors: DataImportError[]
}

export interface WordExportPayload {
  folderPath: string | null
  projectInfo?: ProjectInfo | null
  pvModule: PvModule | null
  record: MeasurementRecord
  recordImages?: RecordImageMap
  records?: MeasurementRecord[]
  pvModules?: PvModule[]
  selectedPvModuleId?: string
  toleranceRows?: ToleranceRow[]
}

export interface RecordImagePoint {
  x: number
  y: number
}

export type RecordImageFit = 'contain' | 'cover' | 'stretch'
export type RecordImageShapeType = 'polygon' | 'rectangle'

export interface RecordImageShape {
  color: string
  points: RecordImagePoint[]
  type: RecordImageShapeType
}

export interface RecordImageSlot {
  dataUrl: string
  fit: RecordImageFit
  name: string
  offsetX: number
  offsetY: number
  polygon: RecordImagePoint[]
  shapes: RecordImageShape[]
  strokeWidth: number
  zoom: number
}

export type RecordImageMap = Record<string, Array<RecordImageSlot | null>>

export interface WordExportResult {
  canceled: boolean
  elapsedMs?: number
  filePath?: string
  warning?: string
}

export interface WordExportProgress {
  completed?: boolean
  elapsedMs?: number
  message: string
  percent: number
}

export interface ProjectInfo {
  projectName: string
  investorName: string
  investorNameEnglish: string
  factoryOwnerName: string
  factoryOwnerNameEnglish: string
  reportTitle: string
  measurementTitle: string
  companyName: string
  companyAddress: string
  preparedBy: string
  checkedBy: string
  approvedBy: string
  ownerApproval: string
  consultantApproval: string
  contractorEpcApproval: string
  testerApproval: string
  applicableStandards: string
}

export interface ProjectInfoLibraryItem extends ProjectInfo {
  id: string
  label: string
  updatedAt: string
}

export interface PvModule {
  id: string
  model: string
  ratedMaximumPowerW: number | null
  openCircuitVoltageV: number | null
  maximumPowerVoltageV: number | null
  shortCircuitCurrentA: number | null
  maximumPowerCurrentA: number | null
  moduleEfficiencyPercent: number | null
  powerTolerance: string
  firstYearDegradationPercent: number | null
  annualDegradationPercent: number | null
  temperatureCoefficientIscPercentPerC: number | null
  temperatureCoefficientVocPercentPerC: number | null
  temperatureCoefficientPmaxPercentPerC: number | null
  createdAt: string
  updatedAt: string
}
