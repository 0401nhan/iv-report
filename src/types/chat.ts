import type { ToleranceRow } from '../lib/tolerance'

export type ChatCommandAction =
  | 'set_view'
  | 'select_data_folder'
  | 'select_record'
  | 'move_record'
  | 'select_module'
  | 'upsert_pv_module'
  | 'set_tolerance'
  | 'export_docx'
  | 'open_exported_file'

export interface ChatPvModulePayload {
  model: string
  ratedMaximumPowerW: number | null
  openCircuitVoltageV: number | null
  maximumPowerVoltageV: number | null
  shortCircuitCurrentA: number | null
  maximumPowerCurrentA: number | null
  moduleEfficiencyPercent: number | null
  powerTolerance: string | null
  firstYearDegradationPercent: number | null
  annualDegradationPercent: number | null
  temperatureCoefficientIscPercentPerC: number | null
  temperatureCoefficientVocPercentPerC: number | null
  temperatureCoefficientPmaxPercentPerC: number | null
}

export interface ChatCommand {
  action: ChatCommandAction
  view: 'home' | 'project-info' | 'pv-module' | null
  direction: 'next' | 'previous' | null
  systemGroup: string | null
  inverter: string | null
  stringName: string | null
  recordQuery: string | null
  moduleId: string | null
  moduleModel: string | null
  rowLabel: 'Tol+' | 'Tol-' | null
  columnKey: string | null
  numericValue: number | null
  pvModule: ChatPvModulePayload | null
}

export interface ChatImageAttachment {
  id: string
  name: string
  mimeType: string
  dataUrl: string
  size: number
}

export interface ChatHistoryItem {
  role: 'user' | 'assistant'
  content: string
  images?: ChatImageAttachment[]
}

export interface ChatContextRecord {
  systemGroup: string
  inverter: string
  stringName: string
  relativePath: string
  status: string
  pfPercent: number | null
  deviationPercent: number | null
}

export interface ChatContextModule {
  id: string
  model: string
}

export interface ChatContextSummaryRow {
  inverter: string
  totalStrings: number
  passCount: number
  failCount: number
  degradationRatePercent: number
}

export interface ChatAppContext {
  activeView: 'home' | 'project-info' | 'pv-module'
  hasImportedData: boolean
  folderPath: string | null
  selected: {
    systemGroup: string | null
    inverter: string | null
    stringName: string | null
    moduleModel: string | null
    recordIndex: number
    totalRecords: number
  }
  navigation: {
    canMoveNext: boolean
    canMovePrevious: boolean
  }
  systems: string[]
  invertersBySystem: Record<string, string[]>
  records: ChatContextRecord[]
  recordsTruncated: boolean
  modules: ChatContextModule[]
  modulesTruncated: boolean
  toleranceRows: ToleranceRow[]
  summaryRows: ChatContextSummaryRow[]
}

export interface ChatCompletionRequest {
  message: string
  history: ChatHistoryItem[]
  context: ChatAppContext
  images: ChatImageAttachment[]
}

export interface ChatCompletionResult {
  reply: string
  commands: ChatCommand[]
  source: 'codex' | 'openai' | 'local' | 'error'
  model: string | null
  error: string | null
}

export interface CodexSetupStatus {
  bundledAvailable: boolean
  globalAvailable: boolean
  installed: boolean
  version: string | null
  hasApiKey: boolean
  hasStoredLogin: boolean
  binaryPath: string | null
  workspacePath: string
  workspaceDefaultPath: string
  workspaceReady: boolean
  workspaceError: string | null
  error: string | null
}

export interface CodexSetupActionResult {
  ok: boolean
  message: string
  output: string
}
