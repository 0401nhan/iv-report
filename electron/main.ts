import type { BrowserWindow as BrowserWindowType, OpenDialogOptions, SaveDialogOptions } from 'electron'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'
import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises'
import { spawn } from 'node:child_process'
import path from 'node:path'
import { Codex } from '@openai/codex-sdk'
import type { Input, ModelReasoningEffort, Thread, ThreadOptions } from '@openai/codex-sdk'
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeightRule,
  HeadingLevel,
  ImageRun,
  Math as DocxMath,
  MathFraction,
  MathRoundBrackets,
  MathRun,
  MathSubScript,
  Packer,
  PageBreak,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  UnderlineType,
  VerticalAlign,
  WidthType,
  type MathComponent,
} from 'docx'
import type {
  DataImportResult,
  IvPoint,
  MeasurementRecord,
  MeasurementSummary,
  ProjectInfo,
  PvModule,
  RecordImageMap,
  RecordImageShape,
  RecordImageSlot,
  WordExportPayload,
  WordExportProgress,
  WordExportResult,
} from '../src/types/data'
import type {
  ChatCommand,
  ChatCommandAction,
  ChatImageAttachment,
  ChatCompletionRequest,
  ChatCompletionResult,
  ChatPvModulePayload,
  CodexSetupActionResult,
  CodexSetupStatus,
} from '../src/types/chat'
import {
  buildPvModuleReferenceMeasurements,
  convertMeasurementSummaryToStc,
  convertIvMeasurementsToStc,
  findPvModuleForRecord,
  getStcConversionProblem,
} from '../src/lib/stc'
import {
  formatTolerancePercent,
  getToleranceCellValue,
  normalizeToleranceRows,
  TOLERANCE_COLUMNS,
  TOLERANCE_TOTAL_KEY,
  type ToleranceRow,
} from '../src/lib/tolerance'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const require = createRequire(import.meta.url)
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron') as typeof import('electron')

// The built directory structure
//
// ├─┬─┬ dist
// │ │ └── index.html
// │ │
// │ ├─┬ dist-electron
// │ │ ├── main.js
// │ │ └── preload.mjs
// │
process.env.APP_ROOT = path.join(__dirname, '..')

// 🚧 Use ['ENV_NAME'] avoid vite:define plugin - Vite@2.x
export const VITE_DEV_SERVER_URL = process.env['VITE_DEV_SERVER_URL']
export const MAIN_DIST = path.join(process.env.APP_ROOT, 'dist-electron')
export const RENDERER_DIST = path.join(process.env.APP_ROOT, 'dist')

process.env.VITE_PUBLIC = VITE_DEV_SERVER_URL ? path.join(process.env.APP_ROOT, 'public') : RENDERER_DIST

let win: BrowserWindowType | null

const IGNORED_DIRECTORIES = new Set(['.git', 'dist', 'dist-electron', 'node_modules'])
const DOCX_REPORT_FONT = 'Times New Roman'
const DOCX_REPORT_FONT_ATTRIBUTES = {
  ascii: DOCX_REPORT_FONT,
  cs: DOCX_REPORT_FONT,
  eastAsia: DOCX_REPORT_FONT,
  hAnsi: DOCX_REPORT_FONT,
}
const DOCX_A4_PAGE_WIDTH_DXA = 11906
const DOCX_A4_PAGE_HEIGHT_DXA = 16838
const DOCX_VERTICAL_MARGIN_DXA = 1440
const DOCX_HORIZONTAL_MARGIN_DXA = 720
const DOCX_OVERVIEW_VERTICAL_MARGIN_DXA = 720
const DOCX_OVERVIEW_HORIZONTAL_MARGIN_DXA = 360
const DOCX_BODY_WIDTH_DXA = DOCX_A4_PAGE_WIDTH_DXA - DOCX_HORIZONTAL_MARGIN_DXA * 2
const DOCX_OVERVIEW_BODY_WIDTH_DXA = DOCX_A4_PAGE_WIDTH_DXA - DOCX_OVERVIEW_HORIZONTAL_MARGIN_DXA * 2
const DOCX_REPORT_COLORS = {
  border: '000000',
  headerBlue: '1F4E79',
  headerLight: 'EAF2F8',
  headerSlate: 'E2E8F0',
  headerSlateDark: '475569',
  ink: '0F172A',
  pass: '2E7D32',
  passLight: 'D9EAD3',
  fail: 'C62828',
  failLight: 'F4CCCC',
  neutral: 'F8FAFC',
  neutralBand: 'F1F5F9',
  neutralDark: 'CBD5E1',
  amber: 'FEF3C7',
  yellow: 'FEFCE8',
  green: 'DCFCE7',
  orange: 'FFEDD5',
  red: 'FEE2E2',
  sky: 'DBEAFE',
}
const DEFAULT_PROJECT_INFO: ProjectInfo = {
  projectName: '',
  investorName: '',
  investorNameEnglish: '',
  factoryOwnerName: '',
  factoryOwnerNameEnglish: '',
  reportTitle: 'PV STRING: I - V CURVE MEASUREMENTS TEST REPORT',
  measurementTitle: 'PV STRINGS I-V CURVE MEASUREMENTS',
  companyName: 'ELECTRIC BIRD HIGH TECHNOLOGY CO., LTD',
  companyAddress: '72 Le Thanh Ton Street, Sai Gon Ward, Ho Chi Minh City, Vietnam.',
  preparedBy: 'Nguyen Trong Nhan',
  checkedBy: 'Nguyen Le Nhat Trung',
  approvedBy: 'Nguyen Le Nhat Trung',
  ownerApproval: '',
  consultantApproval: '',
  contractorEpcApproval: '',
  testerApproval: '',
  applicableStandards: 'IEC 62446-1\nTCVN 11855 - 1:2017',
}
const DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT = 2
const DEFAULT_ANNUAL_DEGRADATION_PERCENT = 0.55
const naturalCollator = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: 'base',
})

let codexClient: Codex | null = null
let codexThread: Thread | null = null

const DEFAULT_PV_MODULES: PvModule[] = [
  {
    id: 'default-longi-lr4-72hph-445',
    model: 'Longi LR4-72HPH-445',
    ratedMaximumPowerW: 445,
    openCircuitVoltageV: 49.1,
    maximumPowerVoltageV: 41.3,
    shortCircuitCurrentA: 11.53,
    maximumPowerCurrentA: 10.78,
    moduleEfficiencyPercent: 20.5,
    powerTolerance: '0~+5W',
    firstYearDegradationPercent: DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent: DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: 0.048,
    temperatureCoefficientVocPercentPerC: -0.27,
    temperatureCoefficientPmaxPercentPerC: -0.35,
    createdAt: '2026-04-25T00:00:00.000Z',
    updatedAt: '2026-04-25T00:00:00.000Z',
  },
  {
    id: 'default-longi-lr4-72hph-450',
    model: 'Longi LR4-72HPH-450',
    ratedMaximumPowerW: 450,
    openCircuitVoltageV: 49.3,
    maximumPowerVoltageV: 41.5,
    shortCircuitCurrentA: 11.6,
    maximumPowerCurrentA: 10.85,
    moduleEfficiencyPercent: 20.7,
    powerTolerance: '0~+5W',
    firstYearDegradationPercent: DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent: DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: 0.048,
    temperatureCoefficientVocPercentPerC: -0.27,
    temperatureCoefficientPmaxPercentPerC: -0.35,
    createdAt: '2026-04-25T00:00:00.000Z',
    updatedAt: '2026-04-25T00:00:00.000Z',
  },
  {
    id: 'default-jam72s30-540-mr-1500v',
    model: 'JAM72S30-540/MR/1500V',
    ratedMaximumPowerW: 540,
    openCircuitVoltageV: 49.6,
    maximumPowerVoltageV: 41.64,
    shortCircuitCurrentA: 13.86,
    maximumPowerCurrentA: 12.97,
    moduleEfficiencyPercent: 20.9,
    powerTolerance: '0~+5W',
    firstYearDegradationPercent: DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent: DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: 0.045,
    temperatureCoefficientVocPercentPerC: -0.275,
    temperatureCoefficientPmaxPercentPerC: -0.35,
    createdAt: '2026-04-25T00:00:00.000Z',
    updatedAt: '2026-04-25T00:00:00.000Z',
  },
  {
    id: 'default-jam72s30-550-mr',
    model: 'JAM72S30-550 MR',
    ratedMaximumPowerW: 550,
    openCircuitVoltageV: 49.9,
    maximumPowerVoltageV: 41.96,
    shortCircuitCurrentA: 14,
    maximumPowerCurrentA: 13.11,
    moduleEfficiencyPercent: 21.3,
    powerTolerance: '0~+5W',
    firstYearDegradationPercent: DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent: DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: 0.045,
    temperatureCoefficientVocPercentPerC: -0.275,
    temperatureCoefficientPmaxPercentPerC: -0.35,
    createdAt: '2026-04-25T00:00:00.000Z',
    updatedAt: '2026-04-25T00:00:00.000Z',
  },
  {
    id: 'default-jam72d10-440-mb',
    model: 'JAM72D10-440 MB',
    ratedMaximumPowerW: 440,
    openCircuitVoltageV: 53.01,
    maximumPowerVoltageV: 44.31,
    shortCircuitCurrentA: 10.32,
    maximumPowerCurrentA: 9.82,
    moduleEfficiencyPercent: 19.7,
    powerTolerance: '0~+5W',
    firstYearDegradationPercent: DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent: DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: 0.044,
    temperatureCoefficientVocPercentPerC: -0.272,
    temperatureCoefficientPmaxPercentPerC: -0.354,
    createdAt: '2026-04-25T00:00:00.000Z',
    updatedAt: '2026-04-25T00:00:00.000Z',
  },
]

interface AppDataFile {
  version: 1
  pvModules: PvModule[]
  aiSettings?: AiSettings
}

interface AiSettings {
  workspacePath: string | null
}

interface DocxReportMetrics {
  deviationPercent: number | null
  ffPercent: number | null
  iscPercent: number | null
  pfPercent: number | null
  status: 'PASS' | 'FAIL' | 'N/A'
  tolMinusPercent: number | null
  tolPlusPercent: number | null
  vocPercent: number | null
}

interface DocxSummaryRow {
  inverter: string
  totalStrings: number
  passCount: number
  failCount: number
  degradationRatePercent: number
}

interface DocxDetailRow extends DocxReportMetrics {
  inverter: string
  stringName: string
  moduleNumber: number | null
  note: string
}

interface DocxSystemReportData extends DocxReportMetrics {
  date: string | null
  impMeasuredA: number | null
  impNominalA: number | null
  impTranslatedA: number | null
  inverter: string | null
  iscMeasuredA: number | null
  iscNominalA: number | null
  iscTranslatedA: number | null
  irradianceWm2: number | null
  model: string | null
  moduleCount: number | null
  pmaxMeasuredW: number | null
  pmaxNominalW: number | null
  pmaxTranslatedW: number | null
  stringName: string | null
  temperatureC: number | null
  time: string | null
  vmpMeasuredV: number | null
  vmpNominalV: number | null
  vmpTranslatedV: number | null
  vocMeasuredV: number | null
  vocNominalV: number | null
  vocTranslatedV: number | null
}

type DocxProgressReporter = (progress: WordExportProgress) => void

function createWindow() {
  win = new BrowserWindow({
    autoHideMenuBar: true,
    icon: path.join(process.env.VITE_PUBLIC, 'electron-vite.svg'),
    minHeight: 720,
    minWidth: 1024,
    show: false,
    title: 'Build IV Report',
    titleBarStyle: 'default',
    width: 1440,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, 'preload.mjs'),
    },
  })

  win.once('ready-to-show', () => {
    win?.maximize()
    win?.show()
  })

  if (VITE_DEV_SERVER_URL) {
    win.loadURL(VITE_DEV_SERVER_URL)
  } else {
    // win.loadFile('dist/index.html')
    win.loadFile(path.join(RENDERER_DIST, 'index.html'))
  }
}

ipcMain.handle('data:select-folder', async () => {
  const aiWorkspacePath = await getAiWorkspacePath()
  const options: OpenDialogOptions = {
    defaultPath: aiWorkspacePath,
    title: 'Select data folder',
    properties: ['openDirectory'],
  }

  const result = win
    ? await dialog.showOpenDialog(win, options)
    : await dialog.showOpenDialog(options)

  if (result.canceled || result.filePaths.length === 0) {
    return null
  }

  return importDataFolder(result.filePaths[0])
})

ipcMain.handle('data:import-ai-workspace', async () => {
  const aiWorkspacePath = await getAiWorkspacePath()

  return importDataFolder(aiWorkspacePath)
})

ipcMain.handle('data:import-folder', async (_event, folderPath: string) => {
  const normalizedFolderPath = typeof folderPath === 'string' ? folderPath.trim() : ''

  if (!normalizedFolderPath) {
    return null
  }

  return importDataFolder(normalizedFolderPath)
})

ipcMain.handle('data:export-docx', async (event, payload: WordExportPayload): Promise<WordExportResult> => {
  const { record } = payload
  const startedAt = Date.now()
  const sendProgress = (percent: number, message: string, completed = false) => {
    event.sender.send('data:export-docx-progress', {
      completed,
      elapsedMs: Date.now() - startedAt,
      message,
      percent,
    } satisfies WordExportProgress)
  }
  const options: SaveDialogOptions = {
    title: 'Export DOCX report',
    defaultPath: path.join(app.getPath('desktop'), `${buildReportFileName(payload.folderPath, record)}.docx`),
    filters: [{ name: 'Word Document', extensions: ['docx'] }],
  }

  sendProgress(5, 'Choosing save location...')

  const result = win
    ? await dialog.showSaveDialog(win, options)
    : await dialog.showSaveDialog(options)

  if (result.canceled || !result.filePath) {
    sendProgress(100, 'Export canceled.', true)
    return { canceled: true }
  }

  sendProgress(10, 'Preparing Word export data...')
  const document = createDocxReport(payload, (progress) => {
    sendProgress(progress.percent, progress.message)
  })
  sendProgress(86, 'Packaging Word document...')
  const buffer = await Packer.toBuffer(document)

  sendProgress(95, 'Writing Word file...')
  const writeResult = await writeReportFileWithFallback(result.filePath, buffer)

  const elapsedMs = Date.now() - startedAt
  const warning = writeResult.renamed
    ? `Original file was locked, so the report was saved as ${path.basename(writeResult.filePath)}.`
    : undefined

  sendProgress(100, warning ?? 'Word export completed.', true)

  return { canceled: false, elapsedMs, filePath: writeResult.filePath, warning }
})

ipcMain.handle('file:open', async (_event, filePath: string): Promise<string | null> => {
  if (!filePath) {
    return 'No file path provided.'
  }

  const errorMessage = await shell.openPath(filePath)

  return errorMessage || null
})

ipcMain.handle('pv-modules:list', async (): Promise<PvModule[]> => {
  return loadPvModules()
})

ipcMain.handle('pv-modules:save', async (_event, modules: PvModule[]): Promise<PvModule[]> => {
  return savePvModules(normalizePvModules(modules))
})

ipcMain.handle('chat:complete', async (_event, request: ChatCompletionRequest): Promise<ChatCompletionResult> => {
  return completeChatRequest(request)
})

ipcMain.handle('codex:status', async (): Promise<CodexSetupStatus> => {
  return getCodexSetupStatus()
})

ipcMain.handle('codex:install', async (): Promise<CodexSetupActionResult> => {
  return installCodexCli()
})

ipcMain.handle('codex:login', async (): Promise<CodexSetupActionResult> => {
  return openCodexLoginTerminal()
})

ipcMain.handle('codex:reset-thread', async (): Promise<CodexSetupActionResult> => {
  codexThread = null

  return {
    ok: true,
    message: 'Codex thread reset.',
    output: '',
  }
})

ipcMain.handle('codex:select-workspace', async (event): Promise<CodexSetupActionResult> => {
  return selectAiWorkspaceFolder(event.sender)
})

ipcMain.handle('codex:set-workspace', async (_event, folderPath: string): Promise<CodexSetupActionResult> => {
  return setAiWorkspaceFolder(folderPath)
})

ipcMain.handle('codex:open-workspace', async (): Promise<CodexSetupActionResult> => {
  return openAiWorkspaceFolder()
})

ipcMain.handle('codex:reset-workspace', async (): Promise<CodexSetupActionResult> => {
  return resetAiWorkspaceFolder()
})

async function importDataFolder(folderPath: string): Promise<DataImportResult> {
  const csvFiles = (await findCsvFiles(folderPath)).sort((left, right) => {
    return left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' })
  })
  const records: MeasurementRecord[] = []
  const errors: DataImportResult['errors'] = []

  for (const filePath of csvFiles) {
    try {
      records.push(await parseMeasurementCsv(folderPath, filePath))
    } catch (error) {
      errors.push({
        filePath,
        message: error instanceof Error ? error.message : 'Unable to parse file',
      })
    }
  }

  return {
    folderPath,
    totalFiles: csvFiles.length,
    records,
    errors,
  }
}

async function completeChatRequest(request: ChatCompletionRequest): Promise<ChatCompletionResult> {
  try {
    return await completeChatRequestWithCodex(request)
  } catch (error) {
    const fallback = completeChatRequestLocally(request, getUnknownErrorMessage(error))

    return {
      ...fallback,
      reply: `${fallback.reply}\n\nCodex SDK error: ${getUnknownErrorMessage(error)}`,
      source: fallback.commands.length > 0 ? 'local' : 'error',
      error: getUnknownErrorMessage(error),
    }
  }
}

async function completeChatRequestWithCodex(request: ChatCompletionRequest): Promise<ChatCompletionResult> {
  const thread = await getCodexThread()
  const input = await createCodexChatInput(request)
  const turn = await thread.run(input, { outputSchema: CHAT_RESPONSE_SCHEMA })
  const modelOutput = normalizeChatModelOutput(parseJsonResponse(turn.finalResponse))

  return {
    ...modelOutput,
    source: 'codex',
    model: getCodexModel(),
    error: null,
  }
}

async function getCodexThread() {
  if (!codexThread) {
    const workspacePath = await getAiWorkspacePath()

    codexThread = getCodexClient().startThread(createCodexThreadOptions(workspacePath))
  }

  return codexThread
}

function getCodexClient() {
  if (!codexClient) {
    const apiKey = process.env.CODEX_API_KEY || process.env.OPENAI_API_KEY
    const codexPathOverride = getCodexPathOverride()
    const options: NonNullable<ConstructorParameters<typeof Codex>[0]> = {
      env: createCodexEnvironment(),
      config: {
        sandbox_workspace_write: {
          network_access: false,
        },
      },
    }

    if (codexPathOverride) {
      options.codexPathOverride = codexPathOverride
    }

    if (apiKey) {
      options.apiKey = apiKey
    }

    codexClient = new Codex(options)
  }

  return codexClient
}

function createCodexThreadOptions(workingDirectory: string): ThreadOptions {
  const model = getCodexModel()
  const reasoningEffort = getCodexReasoningEffort()

  return {
    approvalPolicy: 'never',
    model: model ?? undefined,
    modelReasoningEffort: reasoningEffort,
    networkAccessEnabled: false,
    sandboxMode: 'workspace-write',
    skipGitRepoCheck: true,
    webSearchMode: 'disabled',
    workingDirectory,
  }
}

function createCodexEnvironment() {
  const env = Object.entries(process.env).reduce<Record<string, string>>((values, [key, value]) => {
    if (value !== undefined && key !== 'ELECTRON_RUN_AS_NODE') {
      values[key] = value
    }

    return values
  }, {})
  const apiKey = process.env.CODEX_API_KEY || process.env.OPENAI_API_KEY

  if (apiKey) {
    env.CODEX_API_KEY = apiKey
    env.OPENAI_API_KEY = apiKey
  }

  return env
}

function getCodexModel() {
  return process.env.CODEX_MODEL || process.env.OPENAI_MODEL || process.env.OPENAI_CHAT_MODEL || null
}

function getCodexReasoningEffort(): ModelReasoningEffort {
  const value = process.env.CODEX_REASONING_EFFORT

  return value === 'minimal' || value === 'low' || value === 'medium' || value === 'high' || value === 'xhigh'
    ? value
    : 'low'
}

async function getCodexSetupStatus(): Promise<CodexSetupStatus> {
  const bundledPath = getCodexPathOverride() ?? getLocalCodexBinPath()
  const workspaceStatus = await getAiWorkspaceStatus()
  const globalVersion = await runProcess('codex', ['--version'], { timeoutMs: 15000 })
  const bundledVersion = bundledPath
    ? await runProcess(bundledPath, ['--version'], { timeoutMs: 15000 })
    : null
  const packageVersion = getBundledCodexPackageVersion()
  const hasStoredLogin = await hasCodexStoredLogin()
  const globalAvailable = Boolean(globalVersion?.ok)
  const bundledAvailable = Boolean(bundledVersion?.ok || packageVersion)
  const output = globalVersion?.ok
    ? globalVersion.stdout || globalVersion.stderr
    : bundledVersion?.ok
      ? bundledVersion.stdout || bundledVersion.stderr
      : null

  return {
    binaryPath: bundledPath,
    bundledAvailable,
    error: globalVersion?.ok || bundledVersion?.ok || packageVersion ? null : globalVersion?.stderr ?? null,
    globalAvailable,
    hasApiKey: Boolean(process.env.CODEX_API_KEY || process.env.OPENAI_API_KEY),
    hasStoredLogin,
    installed: globalAvailable || bundledAvailable,
    version: normalizeCodexVersion(output) ?? packageVersion,
    workspaceDefaultPath: workspaceStatus.defaultPath,
    workspaceError: workspaceStatus.error,
    workspacePath: workspaceStatus.path,
    workspaceReady: workspaceStatus.ready,
  }
}

async function selectAiWorkspaceFolder(
  sender: Parameters<typeof BrowserWindow.fromWebContents>[0],
): Promise<CodexSetupActionResult> {
  const currentPath = await getAiWorkspacePath()
  const options: OpenDialogOptions = {
    title: 'Select AI folder',
    defaultPath: currentPath,
    properties: ['openDirectory', 'createDirectory'],
  }
  const parentWindow = BrowserWindow.fromWebContents(sender)
  const result = parentWindow
    ? await dialog.showOpenDialog(parentWindow, options)
    : await dialog.showOpenDialog(options)

  if (result.canceled || result.filePaths.length === 0) {
    return {
      ok: false,
      message: 'AI folder selection canceled.',
      output: currentPath,
    }
  }

  return setAiWorkspaceFolder(result.filePaths[0])
}

async function setAiWorkspaceFolder(folderPath: string): Promise<CodexSetupActionResult> {
  const workspacePath = normalizeWorkspacePath(folderPath)

  if (!workspacePath) {
    return {
      ok: false,
      message: 'AI folder path is empty.',
      output: '',
    }
  }

  await mkdir(workspacePath, { recursive: true })
  await saveAiSettings({ workspacePath })
  resetCodexRuntime()

  return {
    ok: true,
    message: 'AI folder updated.',
    output: workspacePath,
  }
}

async function openAiWorkspaceFolder(): Promise<CodexSetupActionResult> {
  const workspacePath = await getAiWorkspacePath()
  const errorMessage = await shell.openPath(workspacePath)

  return {
    ok: !errorMessage,
    message: errorMessage ? 'Unable to open AI folder.' : 'Opened AI folder.',
    output: errorMessage || workspacePath,
  }
}

async function resetAiWorkspaceFolder(): Promise<CodexSetupActionResult> {
  await saveAiSettings({ workspacePath: null })
  const workspacePath = await getAiWorkspacePath()

  resetCodexRuntime()

  return {
    ok: true,
    message: 'AI folder reset to default.',
    output: workspacePath,
  }
}

async function getAiWorkspaceStatus() {
  const defaultPath = getDefaultAiWorkspacePath()

  try {
    const workspacePath = await getAiWorkspacePath()

    return {
      defaultPath,
      error: null,
      path: workspacePath,
      ready: true,
    }
  } catch (error) {
    return {
      defaultPath,
      error: getUnknownErrorMessage(error),
      path: (await loadAiWorkspacePathSetting()) ?? defaultPath,
      ready: false,
    }
  }
}

async function getAiWorkspacePath() {
  const workspacePath = (await loadAiWorkspacePathSetting()) ?? getDefaultAiWorkspacePath()

  await mkdir(workspacePath, { recursive: true })

  return workspacePath
}

async function loadAiWorkspacePathSetting() {
  const appData = await loadAppDataFile()
  const aiSettings = normalizeAiSettings(appData.aiSettings)

  return aiSettings?.workspacePath ?? null
}

async function saveAiSettings(aiSettings: AiSettings) {
  const currentAppData = await loadAppDataFile()
  const pvModules = normalizePvModules(readPvModulesFromAppData(currentAppData))

  await saveAppDataFile({
    version: 1,
    aiSettings,
    pvModules,
  })
}

function getDefaultAiWorkspacePath() {
  return path.join(app.getPath('documents'), 'AI-Wordspace')
}

function normalizeWorkspacePath(value: unknown) {
  return typeof value === 'string' && value.trim() ? path.resolve(value.trim()) : null
}

function resetCodexRuntime() {
  codexClient = null
  codexThread = null
}

async function installCodexCli(): Promise<CodexSetupActionResult> {
  const result = await runProcess('npm', ['install', '-g', '@openai/codex'], { timeoutMs: 10 * 60 * 1000 })
  const output = [result.stdout, result.stderr].filter(Boolean).join('\n')

  return {
    ok: result.ok,
    message: result.ok ? 'Codex CLI installed.' : 'Codex CLI install failed.',
    output,
  }
}

async function openCodexLoginTerminal(): Promise<CodexSetupActionResult> {
  const commandPath = getCodexPathOverride() ?? getLocalCodexBinPath() ?? 'codex'

  if (process.platform === 'win32') {
    const command = `& ${JSON.stringify(commandPath)} login`
    const child = spawn(
      'powershell.exe',
      ['-NoExit', '-ExecutionPolicy', 'Bypass', '-Command', command],
      {
        detached: true,
        stdio: 'ignore',
        windowsHide: false,
      },
    )

    child.unref()

    return {
      ok: true,
      message: 'Opened Codex login terminal.',
      output: command,
    }
  }

  const result = await runProcess(commandPath, ['login'], { timeoutMs: 5 * 60 * 1000 })

  return {
    ok: result.ok,
    message: result.ok ? 'Codex login completed.' : 'Codex login failed.',
    output: [result.stdout, result.stderr].filter(Boolean).join('\n'),
  }
}

function getCodexPathOverride() {
  if (process.env.CODEX_PATH) {
    return process.env.CODEX_PATH
  }

  if (!app.isPackaged) {
    return null
  }

  const platformPackage = getCodexPlatformPackage()
  const targetTriple = getCodexTargetTriple()

  if (!platformPackage || !targetTriple) {
    return null
  }

  return path.join(
    process.resourcesPath,
    'app.asar.unpacked',
    'node_modules',
    platformPackage,
    'vendor',
    targetTriple,
    'codex',
    process.platform === 'win32' ? 'codex.exe' : 'codex',
  )
}

function getLocalCodexBinPath() {
  const binName = process.platform === 'win32' ? 'codex.cmd' : 'codex'
  const candidatePath = path.join(process.env.APP_ROOT, 'node_modules', '.bin', binName)

  return candidatePath
}

function getBundledCodexPackageVersion() {
  try {
    const packageJson = require('@openai/codex/package.json') as { version?: unknown }

    return typeof packageJson.version === 'string' ? packageJson.version : null
  } catch {
    return null
  }
}

function getCodexPlatformPackage() {
  if (process.platform === 'win32') {
    return process.arch === 'arm64' ? '@openai/codex-win32-arm64' : '@openai/codex-win32-x64'
  }

  if (process.platform === 'darwin') {
    return process.arch === 'arm64' ? '@openai/codex-darwin-arm64' : '@openai/codex-darwin-x64'
  }

  if (process.platform === 'linux') {
    return process.arch === 'arm64' ? '@openai/codex-linux-arm64' : '@openai/codex-linux-x64'
  }

  return null
}

function getCodexTargetTriple() {
  if (process.platform === 'win32') {
    return process.arch === 'arm64' ? 'aarch64-pc-windows-msvc' : 'x86_64-pc-windows-msvc'
  }

  if (process.platform === 'darwin') {
    return process.arch === 'arm64' ? 'aarch64-apple-darwin' : 'x86_64-apple-darwin'
  }

  if (process.platform === 'linux') {
    return process.arch === 'arm64' ? 'aarch64-unknown-linux-musl' : 'x86_64-unknown-linux-musl'
  }

  return null
}

async function hasCodexStoredLogin() {
  try {
    const authPath = path.join(app.getPath('home'), '.codex', 'auth.json')
    const content = await readFile(authPath, 'utf8')

    return content.trim().length > 0
  } catch {
    return false
  }
}

function normalizeCodexVersion(output: string | null) {
  const text = output?.trim() ?? ''

  if (!text) {
    return null
  }

  const match = text.match(/\d+\.\d+\.\d+(?:[-.\w]+)?/)

  return match?.[0] ?? text.split(/\r?\n/)[0] ?? null
}

function runProcess(
  command: string,
  args: string[],
  options: {
    timeoutMs: number
  },
): Promise<{
  code: number | null
  ok: boolean
  stderr: string
  stdout: string
}> {
  return new Promise((resolve) => {
    const child: ReturnType<typeof spawn> = spawn(command, args, {
      env: createCodexEnvironment() as NodeJS.ProcessEnv,
      shell: process.platform === 'win32',
      windowsHide: true,
    })
    const stdout: Buffer[] = []
    const stderr: Buffer[] = []
    const timer = setTimeout(() => {
      child.kill()
    }, options.timeoutMs)

    child.stdout?.on('data', (chunk: Buffer) => stdout.push(chunk))
    child.stderr?.on('data', (chunk: Buffer) => stderr.push(chunk))
    child.on('error', (error: Error) => {
      clearTimeout(timer)
      resolve({
        code: null,
        ok: false,
        stdout: Buffer.concat(stdout).toString('utf8'),
        stderr: error.message,
      })
    })
    child.on('close', (code: number | null) => {
      clearTimeout(timer)
      resolve({
        code,
        ok: code === 0,
        stdout: Buffer.concat(stdout).toString('utf8'),
        stderr: Buffer.concat(stderr).toString('utf8'),
      })
    })
  })
}

async function createCodexChatInput(request: ChatCompletionRequest): Promise<Input> {
  const workspacePath = await getAiWorkspacePath()
  const textInput = `${createChatInstructions()}\n\nCurrent request payload:\n${JSON.stringify(
    createChatTextInputPayload(request, workspacePath),
  )}`
  const imagePaths = await writeCodexImageAttachments(request.images, workspacePath)

  if (imagePaths.length === 0) {
    return textInput
  }

  return [
    { type: 'text', text: textInput },
    ...imagePaths.map((imagePath) => ({ type: 'local_image' as const, path: imagePath })),
  ]
}

async function writeCodexImageAttachments(images: ChatImageAttachment[], workspacePath: string) {
  if (images.length === 0) {
    return []
  }

  const imageDirectory = path.join(workspacePath, '.ai-control', 'attachments')

  await mkdir(imageDirectory, { recursive: true })

  return Promise.all(images.map((image) => writeCodexImageAttachment(image, imageDirectory)))
}

async function writeCodexImageAttachment(image: ChatImageAttachment, directory: string) {
  const parsed = parseImageDataUrl(image.dataUrl)
  const extension = getImageExtension(parsed.mimeType)
  const filePath = path.join(directory, `${Date.now()}-${sanitizeFileName(image.name || image.id)}.${extension}`)

  await writeFile(filePath, parsed.buffer)

  return filePath
}

function parseImageDataUrl(dataUrl: string) {
  const match = dataUrl.match(/^data:([^;,]+);base64,(.+)$/)

  if (!match) {
    throw new Error('Invalid image data URL.')
  }

  const mimeType = match[1]
  const buffer = Buffer.from(match[2], 'base64')

  if (!mimeType.startsWith('image/') || buffer.length === 0) {
    throw new Error('Invalid image attachment.')
  }

  return { buffer, mimeType }
}

function getImageExtension(mimeType: string) {
  if (mimeType === 'image/jpeg') {
    return 'jpg'
  }

  if (mimeType === 'image/png') {
    return 'png'
  }

  if (mimeType === 'image/webp') {
    return 'webp'
  }

  if (mimeType === 'image/gif') {
    return 'gif'
  }

  return 'img'
}

function parseJsonResponse(text: string) {
  const trimmedText = text.trim()

  try {
    return JSON.parse(trimmedText) as unknown
  } catch {
    const match = trimmedText.match(/\{[\s\S]*\}/)

    if (!match) {
      throw new Error('Codex response did not include JSON output.')
    }

    return JSON.parse(match[0]) as unknown
  }
}

function createChatTextInputPayload(request: ChatCompletionRequest, workspacePath: string) {
  return {
    aiWorkspace: {
      path: workspacePath,
      permissions: 'Codex can read and write inside this working directory.',
    },
    appContext: request.context,
    conversation: request.history.map((message) => ({
      role: message.role,
      content: message.content,
      images: message.images?.map((image) => ({
        name: image.name,
        mimeType: image.mimeType,
        size: image.size,
      })) ?? [],
    })),
    userMessage: request.message,
    attachedImages: request.images.map((image) => ({
      name: image.name,
      mimeType: image.mimeType,
      size: image.size,
    })),
  }
}

function completeChatRequestLocally(
  request: ChatCompletionRequest,
  error: string | null = null,
): ChatCompletionResult {
  const normalizedMessage = normalizeChatText(request.message)
  const commands: ChatCommand[] = []

  if (hasAnyText(normalizedMessage, ['pv module', 'thu vien pv', 'module library', 'thu vien module'])) {
    commands.push(createChatCommand('set_view', { view: 'pv-module' }))
  } else if (hasAnyText(normalizedMessage, ['information', 'project info', 'project information', 'thong tin du an'])) {
    commands.push(createChatCommand('set_view', { view: 'project-info' }))
  } else if (hasAnyText(normalizedMessage, ['home', 'trang chu', 'man hinh chinh'])) {
    commands.push(createChatCommand('set_view', { view: 'home' }))
  }

  if (hasAnyText(normalizedMessage, ['add data', 'them data', 'them du lieu', 'chon folder', 'chon thu muc', 'import'])) {
    commands.push(createChatCommand('select_data_folder'))
  }

  const moveDirection = getLocalMoveDirection(normalizedMessage)
  if (moveDirection) {
    commands.push(createChatCommand('move_record', { direction: moveDirection }))
  }

  const selectedModule = findMentionedModule(request)
  if (selectedModule) {
    commands.push(
      createChatCommand('select_module', {
        moduleId: selectedModule.id,
        moduleModel: selectedModule.model,
      }),
    )
  }

  const selectionCommand = createLocalRecordSelectionCommand(request, normalizedMessage)
  if (selectionCommand) {
    commands.push(selectionCommand)
  }

  const toleranceCommand = createLocalToleranceCommand(normalizedMessage)
  if (toleranceCommand) {
    commands.push(toleranceCommand)
  }

  if (hasAnyText(normalizedMessage, ['export', 'docx', 'word', 'xuat bao cao', 'xuat file'])) {
    commands.push(createChatCommand('export_docx'))
  }

  if (hasAnyText(normalizedMessage, ['open file', 'mo file', 'mo bao cao'])) {
    commands.push(createChatCommand('open_exported_file'))
  }

  const uniqueCommands = dedupeChatCommands(commands)

  return {
    reply:
      uniqueCommands.length > 0
        ? 'Tôi sẽ thực hiện các lệnh phù hợp trong app.'
        : request.images.length > 0
          ? 'Tôi đã nhận ảnh, nhưng Codex SDK chưa chạy được nên chưa thể phân tích nội dung ảnh. Parser local chỉ xử lý lệnh chữ.'
          : 'Tôi chưa hiểu lệnh này. Bạn có thể thử: mở PV Module, chọn X3 inverter 4 string 4.9.2, next string, hoặc export DOCX.',
    commands: uniqueCommands,
    source: 'local',
    model: null,
    error,
  }
}

function createChatInstructions() {
  return [
    'You are the in-app controller for an Electron React app named Build IV Report.',
    'Always respond in Vietnamese.',
    'Return only JSON matching the schema. Do not include Markdown.',
    'Use commands only from the allowed action list.',
    'The app imports IV curve CSV data, selects system groups, inverters, strings, PV modules, tolerance rows, and exports DOCX reports.',
    'You can read, create, edit, delete, and organize files inside aiWorkspace.path when the user asks you to work with AI folder data.',
    'Use set_view for Home, Information, or PV Module page changes.',
    'Use select_data_folder when the user asks to add/import/select a data folder. This imports aiWorkspace.path directly without opening a folder picker.',
    'Use select_record for choosing an X/system, inverter, string, or CSV record. Use labels from appContext when possible.',
    'Use move_record for next/previous string navigation.',
    'Use select_module only when the user names a PV module model or clearly asks to select a module.',
    'Use upsert_pv_module when the user asks to add/import/create/update PV module specs or when an attached datasheet/table image contains PV module electrical specifications.',
    'For a datasheet table with multiple model columns, return one upsert_pv_module command per model column.',
    'For upsert_pv_module, fill pvModule.model and every electrical value that is visible. Use null only for fields that are not visible or cannot be confidently read.',
    'Map datasheet labels as: Pmax/W -> ratedMaximumPowerW, Voc/V -> openCircuitVoltageV, Vmp/V or Vmpp/V -> maximumPowerVoltageV, Isc/A -> shortCircuitCurrentA, Imp/A or Impp/A -> maximumPowerCurrentA, module efficiency/% -> moduleEfficiencyPercent, power tolerance -> powerTolerance, temperature coefficient Isc/Voc/Pmax -> temperatureCoefficientIscPercentPerC/temperatureCoefficientVocPercentPerC/temperatureCoefficientPmaxPercentPerC.',
    'If a module model range is shown, expand it into concrete model commands only when each model column and its values are visible.',
    'Use set_tolerance only when row, column, and numeric value are clear.',
    'Use export_docx only when the user explicitly asks to export, save, create Word, or create DOCX.',
    'Use open_exported_file only when the user asks to open the saved report.',
    'If images are attached, inspect them before deciding whether they imply an app command. A screenshot of the app can be used to infer visible selection, errors, or requested UI target.',
    'If the user asks a question, answer briefly in reply and leave commands empty unless an app action is needed.',
  ].join('\n')
}

const CHAT_COMMAND_REQUIRED_KEYS: Array<keyof ChatCommand> = [
  'action',
  'view',
  'direction',
  'systemGroup',
  'inverter',
  'stringName',
  'recordQuery',
  'moduleId',
  'moduleModel',
  'rowLabel',
  'columnKey',
  'numericValue',
  'pvModule',
]

const CHAT_RESPONSE_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['reply', 'commands'],
  properties: {
    reply: { type: 'string' },
    commands: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: CHAT_COMMAND_REQUIRED_KEYS,
        properties: {
          action: {
            type: 'string',
            enum: [
              'set_view',
              'select_data_folder',
              'select_record',
              'move_record',
              'select_module',
              'upsert_pv_module',
              'set_tolerance',
              'export_docx',
              'open_exported_file',
            ],
          },
          view: { type: ['string', 'null'], enum: ['home', 'project-info', 'pv-module', null] },
          direction: { type: ['string', 'null'], enum: ['next', 'previous', null] },
          systemGroup: { type: ['string', 'null'] },
          inverter: { type: ['string', 'null'] },
          stringName: { type: ['string', 'null'] },
          recordQuery: { type: ['string', 'null'] },
          moduleId: { type: ['string', 'null'] },
          moduleModel: { type: ['string', 'null'] },
          rowLabel: { type: ['string', 'null'], enum: ['Tol+', 'Tol-', null] },
          columnKey: { type: ['string', 'null'] },
          numericValue: { type: ['number', 'null'] },
          pvModule: {
            anyOf: [
              { type: 'null' },
              {
                type: 'object',
                additionalProperties: false,
                required: [
                  'model',
                  'ratedMaximumPowerW',
                  'openCircuitVoltageV',
                  'maximumPowerVoltageV',
                  'shortCircuitCurrentA',
                  'maximumPowerCurrentA',
                  'moduleEfficiencyPercent',
                  'powerTolerance',
                  'firstYearDegradationPercent',
                  'annualDegradationPercent',
                  'temperatureCoefficientIscPercentPerC',
                  'temperatureCoefficientVocPercentPerC',
                  'temperatureCoefficientPmaxPercentPerC',
                ],
                properties: {
                  model: { type: 'string' },
                  ratedMaximumPowerW: { type: ['number', 'null'] },
                  openCircuitVoltageV: { type: ['number', 'null'] },
                  maximumPowerVoltageV: { type: ['number', 'null'] },
                  shortCircuitCurrentA: { type: ['number', 'null'] },
                  maximumPowerCurrentA: { type: ['number', 'null'] },
                  moduleEfficiencyPercent: { type: ['number', 'null'] },
                  powerTolerance: { type: ['string', 'null'] },
                  firstYearDegradationPercent: { type: ['number', 'null'] },
                  annualDegradationPercent: { type: ['number', 'null'] },
                  temperatureCoefficientIscPercentPerC: { type: ['number', 'null'] },
                  temperatureCoefficientVocPercentPerC: { type: ['number', 'null'] },
                  temperatureCoefficientPmaxPercentPerC: { type: ['number', 'null'] },
                },
              },
            ],
          },
        },
      },
    },
  },
}

function normalizeChatModelOutput(value: unknown): Pick<ChatCompletionResult, 'reply' | 'commands'> {
  const output = asRecord(value)
  const reply = typeof output?.reply === 'string' && output.reply.trim() ? output.reply.trim() : 'Đã xử lý.'
  const commandSources = Array.isArray(output?.commands) ? output.commands : []
  const commands = commandSources
    .map((commandSource) => normalizeChatCommand(commandSource))
    .filter((command): command is ChatCommand => command !== null)

  return { reply, commands }
}

function normalizeChatCommand(value: unknown): ChatCommand | null {
  const source = asRecord(value)

  if (!source) {
    return null
  }

  const action = typeof source?.action === 'string' ? source.action : ''

  if (!isChatCommandAction(action)) {
    return null
  }

  return createChatCommand(action, {
    view: source.view === 'home' || source.view === 'project-info' || source.view === 'pv-module'
      ? source.view
      : null,
    direction: source.direction === 'next' || source.direction === 'previous' ? source.direction : null,
    systemGroup: normalizeOptionalString(source.systemGroup),
    inverter: normalizeOptionalString(source.inverter),
    stringName: normalizeOptionalString(source.stringName),
    recordQuery: normalizeOptionalString(source.recordQuery),
    moduleId: normalizeOptionalString(source.moduleId),
    moduleModel: normalizeOptionalString(source.moduleModel),
    rowLabel: source.rowLabel === 'Tol+' || source.rowLabel === 'Tol-' ? source.rowLabel : null,
    columnKey: normalizeOptionalString(source.columnKey),
    numericValue: typeof source.numericValue === 'number' && Number.isFinite(source.numericValue)
      ? source.numericValue
      : null,
    pvModule: normalizeChatPvModulePayload(source.pvModule),
  })
}

function createChatCommand(action: ChatCommandAction, values: Partial<ChatCommand> = {}): ChatCommand {
  return {
    action,
    view: null,
    direction: null,
    systemGroup: null,
    inverter: null,
    stringName: null,
    recordQuery: null,
    moduleId: null,
    moduleModel: null,
    rowLabel: null,
    columnKey: null,
    numericValue: null,
    pvModule: null,
    ...values,
  }
}

function normalizeChatPvModulePayload(value: unknown): ChatPvModulePayload | null {
  const source = asRecord(value)
  const model = normalizeOptionalString(source?.model)

  if (!source || !model) {
    return null
  }

  return {
    model,
    ratedMaximumPowerW: normalizeOptionalNumber(source.ratedMaximumPowerW),
    openCircuitVoltageV: normalizeOptionalNumber(source.openCircuitVoltageV),
    maximumPowerVoltageV: normalizeOptionalNumber(source.maximumPowerVoltageV),
    shortCircuitCurrentA: normalizeOptionalNumber(source.shortCircuitCurrentA),
    maximumPowerCurrentA: normalizeOptionalNumber(source.maximumPowerCurrentA),
    moduleEfficiencyPercent: normalizeOptionalNumber(source.moduleEfficiencyPercent),
    powerTolerance: normalizeOptionalString(source.powerTolerance),
    firstYearDegradationPercent: normalizeOptionalNumber(source.firstYearDegradationPercent),
    annualDegradationPercent: normalizeOptionalNumber(source.annualDegradationPercent),
    temperatureCoefficientIscPercentPerC: normalizeOptionalNumber(source.temperatureCoefficientIscPercentPerC),
    temperatureCoefficientVocPercentPerC: normalizeOptionalNumber(source.temperatureCoefficientVocPercentPerC),
    temperatureCoefficientPmaxPercentPerC: normalizeOptionalNumber(source.temperatureCoefficientPmaxPercentPerC),
  }
}

function createLocalRecordSelectionCommand(
  request: ChatCompletionRequest,
  normalizedMessage: string,
): ChatCommand | null {
  const systemGroup = findMentionedSystemGroup(request, normalizedMessage)
  const inverter = findMentionedInverter(request, normalizedMessage, systemGroup)
  const record = findMentionedRecord(request, normalizedMessage, systemGroup, inverter)
  const stringName = record?.stringName ?? getRegexMatch(normalizedMessage, /(?:string|chuoi)\s*([a-z0-9._-]+)/)

  if (!systemGroup && !inverter && !stringName && !record) {
    return null
  }

  return createChatCommand('select_record', {
    systemGroup,
    inverter,
    stringName,
    recordQuery: record?.relativePath ?? stringName,
  })
}

function createLocalToleranceCommand(normalizedMessage: string): ChatCommand | null {
  if (!hasAnyText(normalizedMessage, ['tol+', 'tol-', 'tolerance', 'sai so'])) {
    return null
  }

  const rowLabel = normalizedMessage.includes('tol-') ? 'Tol-' : normalizedMessage.includes('tol+') ? 'Tol+' : null
  const columnKey = findToleranceColumnKey(normalizedMessage)
  const numericValue = readLastNumber(normalizedMessage)

  if (!rowLabel || !columnKey || numericValue === null) {
    return null
  }

  return createChatCommand('set_tolerance', { rowLabel, columnKey, numericValue })
}

function findMentionedSystemGroup(request: ChatCompletionRequest, normalizedMessage: string) {
  const mentionedSystem = request.context.systems.find((system) => normalizedMessage.includes(normalizeChatText(system)))

  if (mentionedSystem) {
    return mentionedSystem
  }

  const systemNumber = getRegexMatch(normalizedMessage, /\bx\s*([0-9]+)/)

  return systemNumber ? `X${systemNumber}` : null
}

function findMentionedInverter(
  request: ChatCompletionRequest,
  normalizedMessage: string,
  systemGroup: string | null,
) {
  const candidateInverters = systemGroup
    ? request.context.invertersBySystem[systemGroup] ?? []
    : Object.values(request.context.invertersBySystem).flat()
  const mentionedInverter = candidateInverters.find((inverter) =>
    normalizedMessage.includes(normalizeChatText(inverter)),
  )

  if (mentionedInverter) {
    return mentionedInverter
  }

  const inverterNumber = getRegexMatch(normalizedMessage, /(?:inverter|inv)\s*([0-9]+)/)

  return inverterNumber ? `Inverter${inverterNumber}` : null
}

function findMentionedRecord(
  request: ChatCompletionRequest,
  normalizedMessage: string,
  systemGroup: string | null,
  inverter: string | null,
) {
  return request.context.records.find((record) => {
    if (systemGroup && normalizeChatText(record.systemGroup) !== normalizeChatText(systemGroup)) {
      return false
    }

    if (inverter && normalizeChatText(record.inverter) !== normalizeChatText(inverter)) {
      return false
    }

    return (
      normalizedMessage.includes(normalizeChatText(record.stringName)) ||
      normalizedMessage.includes(normalizeChatText(record.relativePath))
    )
  }) ?? null
}

function findMentionedModule(request: ChatCompletionRequest) {
  const normalizedMessage = normalizeChatText(request.message)

  if (!hasAnyText(normalizedMessage, ['module', 'tam pin', 'pv'])) {
    return null
  }

  return (
    request.context.modules.find((moduleItem) => normalizedMessage.includes(normalizeChatText(moduleItem.model))) ??
    null
  )
}

function findToleranceColumnKey(normalizedMessage: string) {
  const aliases: Record<string, string[]> = {
    temperatureSensor: ['temperature', 'temp', 'nhiet'],
    radiationSensor: ['radiation', 'irradiance', 'buc xa'],
    pmaxStcTolerance: ['pmax', 'stc'],
    initialModuleDegradation: ['initial', 'ban dau'],
    dcCableLoss: ['dc cable', 'cable', 'day'],
    otherLoss: ['other', 'khac'],
    degradationPerTime: ['per time', 'data sheet', 'suy giam'],
  }

  for (const [columnKey, columnAliases] of Object.entries(aliases)) {
    if (hasAnyText(normalizedMessage, columnAliases)) {
      return columnKey
    }
  }

  const matchedColumn = TOLERANCE_COLUMNS.find((column) => normalizedMessage.includes(normalizeChatText(column.label)))

  return matchedColumn?.key ?? null
}

function getLocalMoveDirection(normalizedMessage: string): 'next' | 'previous' | null {
  if (hasAnyText(normalizedMessage, ['next', 'tiep', 'sau', 'ke tiep'])) {
    return 'next'
  }

  if (hasAnyText(normalizedMessage, ['previous', 'prev', 'truoc', 'lui'])) {
    return 'previous'
  }

  return null
}

function dedupeChatCommands(commands: ChatCommand[]) {
  const seen = new Set<string>()

  return commands.filter((command) => {
    const key = JSON.stringify(command)

    if (seen.has(key)) {
      return false
    }

    seen.add(key)
    return true
  })
}

function isChatCommandAction(value: string): value is ChatCommandAction {
  return [
    'set_view',
    'select_data_folder',
    'select_record',
    'move_record',
    'select_module',
    'upsert_pv_module',
    'set_tolerance',
    'export_docx',
    'open_exported_file',
  ].includes(value)
}

function hasAnyText(source: string, values: string[]) {
  return values.some((value) => source.includes(normalizeChatText(value)))
}

function normalizeChatText(value: string) {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
}

function getRegexMatch(value: string, pattern: RegExp) {
  const match = value.match(pattern)

  return match?.[1] ?? null
}

function readLastNumber(value: string) {
  const matches = value.match(/-?\d+(?:[.,]\d+)?/g)
  const rawValue = matches?.at(-1)

  if (!rawValue) {
    return null
  }

  const parsed = Number(rawValue.replace(',', '.'))

  return Number.isFinite(parsed) ? parsed : null
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' && !Array.isArray(value) ? (value as Record<string, unknown>) : null
}

function normalizeOptionalString(value: unknown) {
  return typeof value === 'string' && value.trim() ? value.trim() : null
}

function normalizeOptionalNumber(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function getUnknownErrorMessage(error: unknown) {
  return error instanceof Error ? error.message : typeof error === 'string' ? error : 'Unknown error'
}

async function loadPvModules(): Promise<PvModule[]> {
  const bundledModules = await loadBundledPvModules()

  try {
    const appData = await loadAppDataFile()
    const rawModules = readPvModulesFromAppData(appData)
    const modules = normalizePvModules(rawModules)

    if (modules.length > 0) {
      const modulesWithDefaults = ensureRequiredPvModules(modules, bundledModules)

      if (modulesNeedDegradationMigration(rawModules) || modulesWithDefaults.length !== modules.length) {
        return savePvModules(modulesWithDefaults)
      }

      return modulesWithDefaults
    }
  } catch (error) {
    if (!isMissingFileError(error)) {
      console.warn('Unable to load PV module library:', error)
    }
  }

  const legacyModules = await loadLegacyPvModules()

  if (legacyModules.length > 0) {
    return savePvModules(ensureRequiredPvModules(legacyModules, bundledModules))
  }

  return savePvModules(ensureRequiredPvModules([], bundledModules))
}

async function savePvModules(modules: PvModule[]): Promise<PvModule[]> {
  const currentAppData = await loadAppDataFile()
  const aiSettings = normalizeAiSettings(currentAppData.aiSettings)
  const appData: AppDataFile = {
    version: 1,
    ...(aiSettings ? { aiSettings } : {}),
    pvModules: modules,
  }

  await saveAppDataFile(appData)

  return modules
}

async function loadAppDataFile(): Promise<Partial<AppDataFile>> {
  try {
    const content = await readFile(getAppDataFilePath(), 'utf8')
    const parsed = JSON.parse(content) as unknown

    if (Array.isArray(parsed)) {
      return {
        version: 1,
        pvModules: normalizePvModules(parsed),
      }
    }

    return asRecord(parsed) ? (parsed as Partial<AppDataFile>) : {}
  } catch (error) {
    if (!isMissingFileError(error)) {
      console.warn('Unable to load app data:', error)
    }

    return {}
  }
}

async function saveAppDataFile(appData: AppDataFile) {
  const filePath = getAppDataFilePath()

  await mkdir(path.dirname(filePath), { recursive: true })
  await writeFile(filePath, JSON.stringify(appData, null, 2), 'utf8')
}

function ensureRequiredPvModules(modules: PvModule[], bundledModules: PvModule[]): PvModule[] {
  const moduleModels = new Set(modules.map((module) => normalizeModuleModel(module.model)))
  const nextModules = [...modules]

  for (const module of [...bundledModules, ...DEFAULT_PV_MODULES]) {
    const moduleModel = normalizeModuleModel(module.model)

    if (!moduleModels.has(moduleModel)) {
      moduleModels.add(moduleModel)
      nextModules.push(module)
    }
  }

  return nextModules
}

async function loadBundledPvModules(): Promise<PvModule[]> {
  try {
    const content = await readFile(getBundledAppDataFilePath(), 'utf8')
    const parsed = JSON.parse(content) as unknown

    return normalizePvModules(readPvModulesFromAppData(parsed))
  } catch (error) {
    if (!isMissingFileError(error)) {
      console.warn('Unable to load bundled PV module library:', error)
    }

    return []
  }
}

async function loadLegacyPvModules(): Promise<PvModule[]> {
  try {
    const content = await readFile(getLegacyPvModulesFilePath(), 'utf8')
    const parsed = JSON.parse(content) as unknown

    return normalizePvModules(Array.isArray(parsed) ? parsed : [])
  } catch (error) {
    if (!isMissingFileError(error)) {
      console.warn('Unable to load legacy PV module library:', error)
    }

    return []
  }
}

function normalizeModuleModel(model: string) {
  return model.trim().toLowerCase()
}

function readPvModulesFromAppData(parsed: unknown) {
  if (Array.isArray(parsed)) {
    return parsed
  }

  if (parsed && typeof parsed === 'object' && 'pvModules' in parsed) {
    const pvModules = (parsed as { pvModules?: unknown }).pvModules
    return Array.isArray(pvModules) ? pvModules : []
  }

  return []
}

function normalizeAiSettings(value: unknown): AiSettings | undefined {
  const source = asRecord(value)

  if (!source) {
    return undefined
  }

  return {
    workspacePath: normalizeWorkspacePath(source.workspacePath),
  }
}

function modulesNeedDegradationMigration(modules: unknown[]) {
  return modules.some((module) => {
    if (!module || typeof module !== 'object') {
      return true
    }

    const source = module as Partial<PvModule>

    return (
      normalizeNullableNumber(source.firstYearDegradationPercent) === null ||
      normalizeNullableNumber(source.annualDegradationPercent) === null
    )
  })
}

function getAppDataFilePath() {
  return path.join(app.getPath('userData'), 'app-data.json')
}

function getBundledAppDataFilePath() {
  return path.join(process.env.APP_ROOT, 'data', 'app-data.seed.json')
}

function getLegacyPvModulesFilePath() {
  return path.join(app.getPath('userData'), 'pv-modules.json')
}

function normalizePvModules(modules: unknown[]): PvModule[] {
  return modules
    .map((module) => normalizePvModule(module))
    .filter((module): module is PvModule => module !== null)
}

function normalizePvModule(module: unknown): PvModule | null {
  if (!module || typeof module !== 'object') {
    return null
  }

  const source = module as Partial<PvModule>
  const model = typeof source.model === 'string' ? source.model.trim() : ''

  if (!model) {
    return null
  }

  const now = new Date().toISOString()

  return {
    id: typeof source.id === 'string' && source.id.trim() ? source.id : createPvModuleId(model),
    model,
    ratedMaximumPowerW: normalizeNullableNumber(source.ratedMaximumPowerW),
    openCircuitVoltageV: normalizeNullableNumber(source.openCircuitVoltageV),
    maximumPowerVoltageV: normalizeNullableNumber(source.maximumPowerVoltageV),
    shortCircuitCurrentA: normalizeNullableNumber(source.shortCircuitCurrentA),
    maximumPowerCurrentA: normalizeNullableNumber(source.maximumPowerCurrentA),
    moduleEfficiencyPercent: normalizeNullableNumber(source.moduleEfficiencyPercent),
    powerTolerance: typeof source.powerTolerance === 'string' ? source.powerTolerance.trim() : '',
    firstYearDegradationPercent:
      normalizeNullableNumber(source.firstYearDegradationPercent) ?? DEFAULT_FIRST_YEAR_DEGRADATION_PERCENT,
    annualDegradationPercent:
      normalizeNullableNumber(source.annualDegradationPercent) ?? DEFAULT_ANNUAL_DEGRADATION_PERCENT,
    temperatureCoefficientIscPercentPerC: normalizeNullableNumber(source.temperatureCoefficientIscPercentPerC),
    temperatureCoefficientVocPercentPerC: normalizeNullableNumber(source.temperatureCoefficientVocPercentPerC),
    temperatureCoefficientPmaxPercentPerC: normalizeNullableNumber(source.temperatureCoefficientPmaxPercentPerC),
    createdAt: typeof source.createdAt === 'string' ? source.createdAt : now,
    updatedAt: typeof source.updatedAt === 'string' ? source.updatedAt : now,
  }
}

function normalizeNullableNumber(value: unknown) {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }

  if (typeof value === 'string' && value.trim()) {
    const parsed = Number(value)
    return Number.isFinite(parsed) ? parsed : null
  }

  return null
}

function createPvModuleId(model: string) {
  return `${sanitizeFileName(model).toLowerCase()}-${Date.now()}`
}

function isMissingFileError(error: unknown) {
  return error instanceof Error && 'code' in error && error.code === 'ENOENT'
}

async function findCsvFiles(folderPath: string): Promise<string[]> {
  const entries = await readdir(folderPath, { withFileTypes: true })
  const files = await Promise.all(
    entries.map(async (entry) => {
      const entryPath = path.join(folderPath, entry.name)

      if (entry.isDirectory()) {
        return IGNORED_DIRECTORIES.has(entry.name) ? [] : findCsvFiles(entryPath)
      }

      return entry.isFile() && entry.name.toLowerCase().endsWith('.csv') ? [entryPath] : []
    }),
  )

  return files.flat()
}

async function parseMeasurementCsv(rootPath: string, filePath: string): Promise<MeasurementRecord> {
  const content = await readFile(filePath, 'utf8')
  const lines = content.replace(/^\uFEFF/, '').split(/\r?\n/)
  const relativePath = path.relative(rootPath, filePath)
  const arrayLocation = findJoinedValue(lines, 'Array Location')
  const { systemGroup, inverter, stringName } = parsePathLocation(relativePath, arrayLocation)

  return {
    sourcePath: filePath,
    relativePath,
    measurementDate: findCellValue(lines, 'Measurement Date'),
    measurementTime: findCellValue(lines, 'Measurement Time'),
    measurementNote: findCellValue(lines, 'Measurement Note'),
    station: findCellValue(lines, 'Project File'),
    arrayLocation,
    systemGroup,
    inverter,
    stringName,
    irradianceWm2: toNumber(findCellValue(lines, 'Irradiance used in model (W/m^2)')),
    cellTemperatureC: toNumber(findCellValue(lines, 'Cell Temperature used in model (Deg C)')),
    latitude: toNumber(findCellValue(lines, 'Latitude')),
    longitude: toNumber(findCellValue(lines, 'Longitude')),
    timeZone: toNumber(findCellValue(lines, 'Time Zone')),
    moduleManufacturer: findJoinedValue(lines, 'Module Mfr'),
    moduleModel: findCellValue(lines, 'Module Model'),
    modulesInString: toNumber(findCellValue(lines, '# of Modules in String')),
    stringsInParallel: toNumber(findCellValue(lines, '# of Strings in Parallel')),
    wireGaugeMm2: toNumber(findCellValue(lines, 'Wire Gauge (mm^2)')),
    wireLengthM: toNumber(findCellValue(lines, 'Wire Length (m; one way)')),
    measurementSummary: parseMeasurementSummary(lines),
    ivMeasurements: parseIvMeasurements(lines),
  }
}

function findCellValue(lines: string[], key: string): string | null {
  const row = findRow(lines, key)

  if (!row || row.length < 2) {
    return null
  }

  return cleanValue(row[1])
}

function findJoinedValue(lines: string[], key: string): string | null {
  const row = findRow(lines, key)

  if (!row || row.length < 2) {
    return null
  }

  return cleanValue(row.slice(1).join(','))
}

function findRow(lines: string[], key: string): string[] | null {
  for (const line of lines) {
    const row = parseCsvLine(line)

    if (cleanValue(row[0]) === key) {
      return row
    }
  }

  return null
}

function parseMeasurementSummary(lines: string[]): MeasurementSummary {
  const summary: MeasurementSummary = {
    pmaxW: null,
    vmppV: null,
    imppA: null,
    vocV: null,
    iscA: null,
  }
  const headerIndex = lines.findIndex((line) => {
    const row = parseCsvLine(line).map((cell) => cleanValue(cell))
    return row[1] === 'MEASUREMENTS' && row[2] === 'MODEL PREDICTIONS'
  })

  if (headerIndex === -1) {
    return summary
  }

  for (const line of lines.slice(headerIndex + 1)) {
    const row = parseCsvLine(line)
    const key = cleanValue(row[0])

    if (!key) {
      break
    }

    const value = toNumber(cleanValue(row[1]))

    if (key === 'Pmax') {
      summary.pmaxW = value
    } else if (key === 'Vmpp') {
      summary.vmppV = value
    } else if (key === 'Impp') {
      summary.imppA = value
    } else if (key === 'Voc') {
      summary.vocV = value
    } else if (key === 'Isc') {
      summary.iscA = value
    }
  }

  return summary
}

function parseIvMeasurements(lines: string[]): IvPoint[] {
  const markerIndex = lines.findIndex((line) => line.trim() === 'IV Measurements:')

  if (markerIndex === -1) {
    return []
  }

  const headerIndex = lines.findIndex((line, index) => {
    return index > markerIndex && line.trim().toUpperCase() === 'VOLTS,AMPS,WATTS'
  })

  if (headerIndex === -1) {
    return []
  }

  const points: IvPoint[] = []

  for (const line of lines.slice(headerIndex + 1)) {
    const row = parseCsvLine(line)
    const volts = toNumber(row[0])
    const amps = toNumber(row[1])
    const watts = toNumber(row[2])

    if (volts === null || amps === null || watts === null) {
      break
    }

    points.push({ volts, amps, watts })
  }

  return points
}

function parseCsvLine(line: string): string[] {
  const cells: string[] = []
  let cell = ''
  let insideQuotes = false

  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]
    const nextChar = line[index + 1]

    if (char === '"' && insideQuotes && nextChar === '"') {
      cell += char
      index += 1
    } else if (char === '"') {
      insideQuotes = !insideQuotes
    } else if (char === ',' && !insideQuotes) {
      cells.push(cell)
      cell = ''
    } else {
      cell += char
    }
  }

  cells.push(cell)
  return cells
}

function cleanValue(value: string | undefined): string | null {
  const cleaned = value?.trim().replace(/^"|"$/g, '') ?? ''
  const normalizedValue = cleaned.toLowerCase()

  return cleaned.length > 0 && cleaned !== '=NA()' && normalizedValue !== 'none' && normalizedValue !== '(none)'
    ? cleaned
    : null
}

function toNumber(value: string | null | undefined): number | null {
  if (!value) {
    return null
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

function parsePathLocation(relativePath: string, arrayLocation: string | null) {
  const pathParts = relativePath.split(/[\\/]/)
  const systemIndex = pathParts.findIndex((part) => /^X\d+$/i.test(part))
  const inverterIndex = pathParts.findIndex((part) => /^Inverter\d+$/i.test(part))
  const arrayLocationParts = parseArrayLocation(arrayLocation)
  const systemGroup = systemIndex >= 0 ? cleanValue(pathParts[systemIndex]) : null
  const pathInverter =
    systemIndex >= 0 ? cleanValue(pathParts[systemIndex + 1]) : cleanValue(pathParts[inverterIndex])
  const pathString =
    systemIndex >= 0
      ? cleanValue(pathParts[systemIndex + 2])
      : inverterIndex >= 0
        ? cleanValue(pathParts[inverterIndex + 1])
        : null

  return {
    systemGroup,
    inverter: pathInverter ?? arrayLocationParts.inverter,
    stringName: pathString ?? arrayLocationParts.stringName,
  }
}

function parseArrayLocation(arrayLocation: string | null) {
  if (!arrayLocation) {
    return { inverter: null, stringName: null }
  }

  const [inverter, ...stringParts] = arrayLocation.split('-')

  return {
    inverter: cleanValue(inverter),
    stringName: cleanValue(stringParts.join('-')),
  }
}

function createDocxReport({
  folderPath,
  projectInfo,
  record,
  recordImages = {},
  records = [],
  pvModules = [],
  selectedPvModuleId = '',
  toleranceRows,
}: WordExportPayload, reportProgress?: DocxProgressReporter) {
  const normalizedToleranceRows = normalizeToleranceRows(toleranceRows)
  const normalizedProjectInfo = normalizeDocxProjectInfo(projectInfo, folderPath, record)
  const reportRecords = records.length > 0 ? records : [record]
  const selectedPvModule =
    pvModules.find((moduleItem) => moduleItem.id === selectedPvModuleId) ??
    findPvModuleForRecord(record, pvModules)
  const documentTitleProjectName =
    normalizedProjectInfo.projectName || (folderPath ? path.basename(folderPath) : record.station ?? '')
  reportProgress?.({ message: 'Building project information pages...', percent: 12 })
  const projectInformationBlocks = createDocxProjectInformationBlocks(normalizedProjectInfo, selectedPvModule)
  reportProgress?.({ message: 'Building PV information page...', percent: 18 })
  const pvInformationTable = createDocxPvInformationTable(selectedPvModule)
  reportProgress?.({ message: 'Building summary tables...', percent: 24 })
  const summaryRows = buildDocxSummaryRows(reportRecords, pvModules, selectedPvModuleId, normalizedToleranceRows)
  reportProgress?.({ message: 'Building string detail table...', percent: 30 })
  const detailRows = buildDocxDetailRows(reportRecords, pvModules, selectedPvModuleId, normalizedToleranceRows)
  const systemReportBlocks = createDocxSystemReportBlocks(
    reportRecords,
    pvModules,
    selectedPvModuleId,
    normalizedToleranceRows,
    folderPath,
    recordImages,
    reportProgress,
  )

  return new Document({
    title: `IV Report ${documentTitleProjectName}`.trim(),
    creator: 'Build IV Report',
    styles: {
      default: {
        document: {
          run: {
            font: DOCX_REPORT_FONT,
            size: 20,
          },
        },
        title: {
          run: {
            bold: true,
            font: DOCX_REPORT_FONT,
            size: 32,
          },
        },
        heading1: {
          run: {
            bold: true,
            font: DOCX_REPORT_FONT,
            size: 26,
          },
        },
        heading2: {
          run: {
            bold: true,
            font: DOCX_REPORT_FONT,
            size: 22,
          },
        },
        heading3: {
          run: {
            bold: true,
            font: DOCX_REPORT_FONT,
            size: 20,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: DOCX_A4_PAGE_WIDTH_DXA, height: DOCX_A4_PAGE_HEIGHT_DXA },
            margin: {
              top: DOCX_VERTICAL_MARGIN_DXA,
              right: DOCX_HORIZONTAL_MARGIN_DXA,
              bottom: DOCX_VERTICAL_MARGIN_DXA,
              left: DOCX_HORIZONTAL_MARGIN_DXA,
            },
          },
        },
        children: [
          ...projectInformationBlocks,
          createPageBreak(),
          createDocxParagraph('PV MODULE DATASHEET', {
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
          pvInformationTable,
          createPageBreak(),
          createDocxParagraph('TOLERANCE CALCULATION DETAILS', {
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
          createToleranceCalculationDetailsTable(normalizedToleranceRows),
          new Paragraph(''),
          createDocxFormulaReferenceTable(),
        ],
      },
      {
        properties: {
          page: {
            size: { width: DOCX_A4_PAGE_WIDTH_DXA, height: DOCX_A4_PAGE_HEIGHT_DXA },
            margin: {
              top: DOCX_OVERVIEW_VERTICAL_MARGIN_DXA,
              right: DOCX_OVERVIEW_HORIZONTAL_MARGIN_DXA,
              bottom: DOCX_OVERVIEW_VERTICAL_MARGIN_DXA,
              left: DOCX_OVERVIEW_HORIZONTAL_MARGIN_DXA,
            },
          },
        },
        children: [
          createDocxParagraph('OVERVIEW LAYOUT', {
            heading: HeadingLevel.TITLE,
            alignment: AlignmentType.CENTER,
          }),
          createDocxSummaryTable(summaryRows),
          createPageBreak(),
          createDocxDetailTable(detailRows),
          createPageBreak(),
          ...systemReportBlocks,
        ],
      },
    ],
  })
}

function createDocxProjectInformationBlocks(projectInfo: ProjectInfo, pvModule: PvModule | null) {
  const blocks: Array<Paragraph | Table> = [
    createDocxProjectCoverTable(projectInfo),
    createDocxFrontMatterSpacer(360),
    createDocxProjectPartiesTable(projectInfo),
    createDocxFrontMatterSpacer(430),
    createDocxMeasurementTitleTable(projectInfo, pvModule),
    createDocxFrontMatterSpacer(360),
    createDocxApprovalCaption(),
    createDocxProjectApprovalTable(projectInfo),
    createPageBreak(),
    createDocxApplicableStandardsTable(projectInfo),
  ]

  return blocks
}

function createDocxFrontMatterSpacer(after: number) {
  return new Paragraph({
    spacing: { after, before: 0 },
  })
}

function createDocxApprovalCaption() {
  return new Paragraph({
    spacing: { after: 40, before: 0 },
    children: [
      new TextRun({
        text: 'PHÊ DUYỆT BỞI / APPROVAL BY',
        font: DOCX_REPORT_FONT_ATTRIBUTES,
        size: 22,
      }),
    ],
  })
}

function createDocxCoverSignoffCell(dateLabel: string, actionLabel: string, engineerLabel: string, width: number) {
  const labelWidth = 820
  const valueWidth = width - labelWidth - 180

  return new TableCell({
    borders: createDocxCellBorders(),
    margins: { top: 40, bottom: 30, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    children: [
      new Table({
        borders: createDocxNoTableBorders(),
        columnWidths: [labelWidth, valueWidth],
        layout: TableLayoutType.FIXED,
        width: { size: width - 160, type: WidthType.DXA },
        rows: [
          createDocxTableRow(
            [
              createDocxBorderlessTextCell(dateLabel, labelWidth, AlignmentType.LEFT, 20),
              createDocxBorderlessTextCell('', valueWidth, AlignmentType.CENTER, 20),
            ],
            { height: 170 },
          ),
          createDocxTableRow(
            [
              createDocxBorderlessTextCell('', labelWidth, AlignmentType.LEFT, 20),
              createDocxBorderlessTextCell(actionLabel, valueWidth, AlignmentType.CENTER, 20),
            ],
            { height: 240 },
          ),
          createDocxTableRow(
            [
              createDocxBorderlessTextCell('Engineer', labelWidth, AlignmentType.LEFT, 20),
              createDocxBorderlessTextCell(engineerLabel.replace(/^Engineer\s+/i, ''), valueWidth, AlignmentType.CENTER, 20),
            ],
            { height: 170 },
          ),
        ],
      }),
    ],
  })
}

function createDocxBorderlessTextCell(
  text: string,
  width: number,
  alignment: (typeof AlignmentType)[keyof typeof AlignmentType],
  size: number,
) {
  return new TableCell({
    borders: createDocxNoCellBorders(),
    margins: { top: 0, bottom: 0, left: 0, right: 0 },
    verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    children: [
      new Paragraph({
        alignment,
        spacing: { after: 0, before: 0 },
        children: createDocxTextRuns(text, { size }),
      }),
    ],
  })
}

function createDocxCoverTitleCell(projectInfo: ProjectInfo, width: number) {
  return createFrontMatterCell(
    [
      {
        text: projectInfo.companyName,
        alignment: AlignmentType.CENTER,
        bold: true,
        size: 28,
        spacingAfter: 420,
      },
      {
        text: projectInfo.companyAddress,
        alignment: AlignmentType.CENTER,
        size: 18,
        spacingAfter: 220,
      },
      {
        text: projectInfo.reportTitle,
        alignment: AlignmentType.CENTER,
        bold: true,
        size: 30,
      },
    ],
    { rowSpan: 3, width, margins: { top: 120, bottom: 120, left: 160, right: 160 } },
  )
}

function createDocxSignatureCell(label: string, value: string, width: number) {
  const paragraphs: FrontMatterCellParagraph[] = [
    {
      text: label,
      alignment: AlignmentType.LEFT,
      bold: true,
      italics: true,
      underline: true,
      size: 20,
    },
  ]

  if (value.trim()) {
    paragraphs.push({
      text: value.trim(),
      alignment: AlignmentType.CENTER,
      size: 20,
      spacingBefore: 900,
    })
  }

  return createFrontMatterCell(paragraphs, {
    verticalAlign: VerticalAlign.TOP,
    width,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
  })
}

interface FrontMatterCellParagraph {
  alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
  bold?: boolean
  color?: string
  italics?: boolean
  size?: number
  spacingAfter?: number
  spacingBefore?: number
  text: string
  underline?: boolean
}

function createFrontMatterCell(
  content: string | FrontMatterCellParagraph[],
  options: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
    bold?: boolean
    color?: string
    columnSpan?: number
    margins?: { top: number; bottom: number; left: number; right: number }
    rowSpan?: number
    size?: number
    verticalAlign?: Exclude<(typeof VerticalAlign)[keyof typeof VerticalAlign], 'both'>
    width?: number
  } = {},
) {
  const paragraphs =
    typeof content === 'string'
      ? [
          {
            text: content,
            alignment: options.alignment,
            bold: options.bold,
            color: options.color,
            size: options.size,
          },
        ]
      : content

  return new TableCell({
    borders: createDocxCellBorders(),
    columnSpan: options.columnSpan,
    margins: options.margins ?? { top: 80, bottom: 80, left: 100, right: 100 },
    rowSpan: options.rowSpan,
    verticalAlign: options.verticalAlign ?? VerticalAlign.CENTER,
    width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
    children: paragraphs.map(
      (paragraph) =>
        new Paragraph({
          alignment: paragraph.alignment ?? options.alignment ?? AlignmentType.CENTER,
          spacing: {
            after: paragraph.spacingAfter ?? 0,
            before: paragraph.spacingBefore ?? 0,
          },
          children: createDocxTextRuns(paragraph.text, {
            bold: paragraph.bold ?? options.bold,
            color: paragraph.color ?? options.color,
            italics: paragraph.italics,
            size: paragraph.size ?? options.size ?? 20,
            underline: paragraph.underline,
          }),
        }),
    ),
  })
}

function createDocxProjectCoverTable(projectInfo: ProjectInfo) {
  const leftWidth = 2820
  const rightWidth = DOCX_BODY_WIDTH_DXA - leftWidth

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [leftWidth, rightWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createDocxCoverSignoffCell('Date', 'Prepared by', formatEngineerLabel(projectInfo.preparedBy), leftWidth),
          createDocxCoverTitleCell(projectInfo, rightWidth),
        ],
        { height: 700 },
      ),
      createDocxTableRow(
        [createDocxCoverSignoffCell('Date', 'Check by', formatEngineerLabel(projectInfo.checkedBy), leftWidth)],
        { height: 700 },
      ),
      createDocxTableRow(
        [createDocxCoverSignoffCell('Date', 'Approved by', formatEngineerLabel(projectInfo.approvedBy), leftWidth)],
        { height: 700 },
      ),
    ],
  })
}

function createDocxProjectPartiesTable(projectInfo: ProjectInfo) {
  const labelWidth = 2700
  const valueWidth = DOCX_BODY_WIDTH_DXA - labelWidth
  const rows: Array<[string, string]> = [
    ['Nhà đầu tư\nInvestors', `${projectInfo.investorName}\n${projectInfo.investorNameEnglish}`.trim()],
    ['Chủ nhà máy\nFactory Owner', `${projectInfo.factoryOwnerName}\n${projectInfo.factoryOwnerNameEnglish}`.trim()],
  ]

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    layout: TableLayoutType.FIXED,
    rows: rows.map(([label, value]) =>
      createDocxTableRow(
        [
          createFrontMatterCell(label, { width: labelWidth, size: 22 }),
          createFrontMatterCell(value, { width: valueWidth, size: 22 }),
        ],
        { height: 560 },
      ),
    ),
  })
}

function createDocxMeasurementTitleTable(projectInfo: ProjectInfo, pvModule: PvModule | null) {
  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [DOCX_BODY_WIDTH_DXA],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createFrontMatterCell(projectInfo.measurementTitle, {
            alignment: AlignmentType.CENTER,
            bold: true,
            size: 34,
            width: DOCX_BODY_WIDTH_DXA,
          }),
        ],
        { height: 640 },
      ),
      createDocxTableRow(
        [
          createFrontMatterCell(pvModule?.model ?? '', {
            alignment: AlignmentType.CENTER,
            bold: true,
            color: 'FF0000',
            size: 34,
            width: DOCX_BODY_WIDTH_DXA,
          }),
        ],
        { height: 540 },
      ),
    ],
  })
}

function createDocxProjectApprovalTable(projectInfo: ProjectInfo) {
  const halfWidth = Math.floor(DOCX_BODY_WIDTH_DXA / 2)

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [halfWidth, DOCX_BODY_WIDTH_DXA - halfWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createDocxSignatureCell('OWNER:', projectInfo.ownerApproval, halfWidth),
          createDocxSignatureCell('CONSULTANT (Owner Engineer):', projectInfo.consultantApproval, DOCX_BODY_WIDTH_DXA - halfWidth),
        ],
        { height: 3000 },
      ),
      createDocxTableRow(
        [
          createDocxSignatureCell('CONTRACTOR EPC:', projectInfo.contractorEpcApproval, halfWidth),
          createDocxSignatureCell('TESTER:', projectInfo.testerApproval, DOCX_BODY_WIDTH_DXA - halfWidth),
        ],
        { height: 3000 },
      ),
    ],
  })
}

function createDocxApplicableStandardsTable(projectInfo: ProjectInfo) {
  const standards = splitDocxProjectInfoLines(projectInfo.applicableStandards)

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [Math.floor(DOCX_BODY_WIDTH_DXA / 2), Math.ceil(DOCX_BODY_WIDTH_DXA / 2)],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('Tiêu chuẩn áp dụng', {
            alignment: AlignmentType.CENTER,
            bold: true,
            width: Math.floor(DOCX_BODY_WIDTH_DXA / 2),
            size: 22,
          }),
          createReportTableCell('Applicable Standards', {
            alignment: AlignmentType.CENTER,
            bold: true,
            width: Math.ceil(DOCX_BODY_WIDTH_DXA / 2),
            size: 22,
          }),
        ],
        { header: true, height: 480 },
      ),
      ...standards.map((standard) =>
        createDocxTableRow(
          [
            createReportTableCell(standard, {
              alignment: AlignmentType.CENTER,
              columnSpan: 2,
              fill: 'FFFFFF',
              size: 22,
            }),
          ],
          { height: 460 },
        ),
      ),
    ],
  })
}

function createDocxPvInformationTable(pvModule: PvModule | null) {
  const labelWidth = 6800
  const valueWidth = DOCX_BODY_WIDTH_DXA - labelWidth
  const rows: Array<[string, string]> = [
    ['Model', pvModule?.model ?? ''],
    ['Rated Maximum Power(Pmax) [W]', formatDocxPvModuleValue(pvModule?.ratedMaximumPowerW ?? null)],
    ['Open Circuit Voltage(Voc) [V] (Tolerance +/- 3%)', formatDocxPvModuleValue(pvModule?.openCircuitVoltageV ?? null)],
    ['Maximum Power Voltage(Vmp) [V]', formatDocxPvModuleValue(pvModule?.maximumPowerVoltageV ?? null)],
    ['Short Circuit Current(Isc) [A] (Tolerance +/- 3%)', formatDocxPvModuleValue(pvModule?.shortCircuitCurrentA ?? null)],
    ['Maximum Power Current(Imp) [A]', formatDocxPvModuleValue(pvModule?.maximumPowerCurrentA ?? null)],
    ['Module Efficiency [%]', formatDocxPvModuleValue(pvModule?.moduleEfficiencyPercent ?? null)],
    ['Power Tolerance', pvModule?.powerTolerance ?? ''],
    [
      'Temperature Coefficient of Isc(Alpha+/-_Isc) [Deg C]',
      formatDocxPvModuleValue(pvModule?.temperatureCoefficientIscPercentPerC ?? null),
    ],
    [
      'Temperature Coefficient of Voc(Beta_Voc) [Deg C]',
      formatDocxPvModuleValue(pvModule?.temperatureCoefficientVocPercentPerC ?? null),
    ],
    [
      'Temperature Coefficient of Pmax(Gama_Pmp) [Deg C]',
      formatDocxPvModuleValue(pvModule?.temperatureCoefficientPmaxPercentPerC ?? null),
    ],
  ]

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [labelWidth, valueWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('PV INFORMATION', {
            alignment: AlignmentType.CENTER,
            bold: true,
            color: 'FFFFFF',
            columnSpan: 2,
            fill: DOCX_REPORT_COLORS.headerBlue,
            size: 26,
          }),
        ],
        { header: true, height: 520 },
      ),
      ...rows.map(([label, value], index) => {
        const fill = index % 2 === 0 ? DOCX_REPORT_COLORS.headerLight : 'FFFFFF'

        return createDocxTableRow(
          [
            createReportTableCell(label, { alignment: AlignmentType.LEFT, fill, width: labelWidth, size: 22 }),
            createReportTableCell(value, { alignment: AlignmentType.CENTER, fill, width: valueWidth, size: 22 }),
          ],
          { height: 420 },
        )
      }),
    ],
  })
}

function createDocxSummaryTable(rows: DocxSummaryRow[]) {
  const colWidths = [2800, 2400, 2000, 2000, DOCX_OVERVIEW_BODY_WIDTH_DXA - 9200]

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_OVERVIEW_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: colWidths,
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('PV STRINGS - IV CURVE ANALYSIS REPORT SUMMARY', {
            alignment: AlignmentType.CENTER,
            bold: true,
            columnSpan: 5,
            fill: DOCX_REPORT_COLORS.headerBlue,
            color: 'FFFFFF',
            size: 22,
          }),
        ],
        { header: true, height: 560 },
      ),
      createDocxTableRow(
        [
          createReportTableCell('INVERTER', { bold: true, fill: DOCX_REPORT_COLORS.neutral, width: colWidths[0], size: 18 }),
          createReportTableCell('TOTAL OF STRING', { bold: true, fill: DOCX_REPORT_COLORS.neutral, width: colWidths[1], size: 18 }),
          createReportTableCell('PASS', {
            bold: true,
            color: 'FFFFFF',
            fill: DOCX_REPORT_COLORS.pass,
            width: colWidths[2],
            size: 18,
          }),
          createReportTableCell('FAIL', {
            bold: true,
            color: 'FFFFFF',
            fill: DOCX_REPORT_COLORS.fail,
            width: colWidths[3],
            size: 18,
          }),
          createReportTableCell('DEG RATE\n(%)', { bold: true, fill: DOCX_REPORT_COLORS.neutral, width: colWidths[4], size: 18 }),
        ],
        { header: true, height: 600 },
      ),
      ...rows.map(
        (row, index) => {
          const rowFill = index % 2 === 0 ? 'FFFFFF' : DOCX_REPORT_COLORS.neutralBand

          return createDocxTableRow(
            [
              createReportTableCell(row.inverter, { fill: rowFill, width: colWidths[0] }),
              createReportTableCell(row.totalStrings.toString(), { fill: rowFill, width: colWidths[1] }),
              createReportTableCell(row.passCount.toString(), { fill: DOCX_REPORT_COLORS.passLight, width: colWidths[2] }),
              createReportTableCell(row.failCount.toString(), { fill: DOCX_REPORT_COLORS.failLight, width: colWidths[3] }),
              createReportTableCell(formatDocxPercent(row.degradationRatePercent, 0), {
                fill: rowFill,
                width: colWidths[4],
              }),
            ],
            { height: 360 },
          )
        },
      ),
    ],
  })
}

function createDocxDetailTable(rows: DocxDetailRow[]) {
  const colWidths = [1300, 800, 900, 800, 800, 650, 650, 620, 620, 1100, 900, DOCX_OVERVIEW_BODY_WIDTH_DXA - 9140]

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_OVERVIEW_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: colWidths,
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('Inv.', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[0] }),
          createReportTableCell('String', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[1] }),
          createReportTableCell('Module\nNumber.', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[2] }),
          createReportTableCell('Isc (%)', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[3] }),
          createReportTableCell('Voc (%)', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[4] }),
          createReportTableCell('PF', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[5] }),
          createReportTableCell('FF', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[6] }),
          createReportTableCell('Tol+', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[7] }),
          createReportTableCell('Tol-', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[8] }),
          createReportTableCell('Deviation\nfrom\nexpected', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[9] }),
          createReportTableCell('Pass/Fail', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[10] }),
          createReportTableCell('Note', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate, width: colWidths[11] }),
        ],
        { header: true, height: 560 },
      ),
      ...rows.map((row, index) => {
        const rowFill = index % 2 === 0 ? 'FFFFFF' : DOCX_REPORT_COLORS.neutralBand
        const statusFill =
          row.status === 'PASS' ? DOCX_REPORT_COLORS.pass : row.status === 'FAIL' ? DOCX_REPORT_COLORS.fail : undefined
        const deviationFill =
          row.status === 'PASS'
            ? DOCX_REPORT_COLORS.passLight
            : row.status === 'FAIL'
              ? DOCX_REPORT_COLORS.failLight
              : undefined
        const statusTextColor = row.status === 'PASS' || row.status === 'FAIL' ? 'FFFFFF' : undefined

        return createDocxTableRow(
          [
            createReportTableCell(row.inverter, { fill: rowFill, width: colWidths[0] }),
            createReportTableCell(row.stringName, { fill: rowFill, width: colWidths[1] }),
            createReportTableCell(formatDocxInteger(row.moduleNumber), { fill: rowFill, width: colWidths[2] }),
            createReportTableCell(formatDocxPercent(row.iscPercent, 1), { fill: DOCX_REPORT_COLORS.passLight, width: colWidths[3] }),
            createReportTableCell(formatDocxPercent(row.vocPercent, 1), { fill: DOCX_REPORT_COLORS.passLight, width: colWidths[4] }),
            createReportTableCell(formatDocxPercent(row.pfPercent, 1), { fill: DOCX_REPORT_COLORS.passLight, width: colWidths[5] }),
            createReportTableCell(formatDocxPercent(row.ffPercent, 1), { fill: DOCX_REPORT_COLORS.passLight, width: colWidths[6] }),
            createReportTableCell(formatDocxPercent(row.tolPlusPercent, 1), { fill: rowFill, width: colWidths[7] }),
            createReportTableCell(formatDocxPercent(row.tolMinusPercent, 1), { fill: rowFill, width: colWidths[8] }),
            createReportTableCell(formatDocxPercent(row.deviationPercent, 1), { fill: deviationFill, width: colWidths[9] }),
            createReportTableCell(row.status, {
              bold: true,
              color: statusTextColor,
              fill: statusFill ?? rowFill,
              width: colWidths[10],
            }),
            createReportTableCell(row.note, { alignment: AlignmentType.LEFT, fill: rowFill, width: colWidths[11] }),
          ],
          { height: 360 },
        )
      }),
    ],
  })
}

function createDocxSystemReportBlocks(
  records: MeasurementRecord[],
  pvModules: PvModule[],
  selectedPvModuleId: string,
  toleranceRows: ToleranceRow[],
  folderPath: string | null,
  recordImages: RecordImageMap,
  reportProgress?: DocxProgressReporter,
) {
  const selectedModule = pvModules.find((moduleItem) => moduleItem.id === selectedPvModuleId) ?? null
  const sortedRecords = sortDocxRecordsByHierarchy(records)
  const blocks: Array<Paragraph | Table> = []

  sortedRecords.forEach((record, index) => {
    const moduleItem = selectedModule ?? findPvModuleForRecord(record, pvModules)
    const report = buildDocxSystemReportData(record, moduleItem, toleranceRows)
    const progress = 34 + Math.round(((index + 1) / Math.max(sortedRecords.length, 1)) * 46)

    reportProgress?.({
      message: `Building System page ${index + 1}/${sortedRecords.length}: ${report.inverter ?? ''} ${report.stringName ?? ''}`.trim(),
      percent: progress,
    })

    if (index > 0) {
      blocks.push(createPageBreak())
    }

    blocks.push(createDocxSystemTable(report))
    blocks.push(new Paragraph(''))
    blocks.push(createDocxChartPairTable(record, moduleItem))
    blocks.push(new Paragraph(''))
    blocks.push(createDocxRecordImageSlotsTable(getDocxRecordImageSlots(folderPath, record, recordImages)))
  })

  return blocks
}

function createDocxSystemTable(report: DocxSystemReportData) {
  const columnCount = 13
  const baseColumnWidth = Math.floor(DOCX_OVERVIEW_BODY_WIDTH_DXA / columnCount)
  const columnWidths = Array.from({ length: columnCount }, (_value, index) =>
    index === columnCount - 1 ? DOCX_OVERVIEW_BODY_WIDTH_DXA - baseColumnWidth * (columnCount - 1) : baseColumnWidth,
  )
  const statusFill =
    report.status === 'FAIL'
      ? DOCX_REPORT_COLORS.fail
      : report.status === 'N/A'
        ? DOCX_REPORT_COLORS.neutralDark
        : DOCX_REPORT_COLORS.pass
  const statusTextColor = report.status === 'N/A' ? '111827' : 'FFFFFF'
  const createSystemTableCell = (
    text: string,
    options: Parameters<typeof createReportTableCell>[1] = {},
  ) => createReportTableCell(text, { alignment: AlignmentType.CENTER, ...options })

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_OVERVIEW_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths,
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createSystemTableCell('System', {
            bold: true,
            columnSpan: 2,
            fill: DOCX_REPORT_COLORS.neutral,
            rowSpan: 2,
            size: 28,
          }),
          createSystemTableCell(`Model: ${report.model ?? ''}`, {
            bold: true,
            columnSpan: 4,
            fill: DOCX_REPORT_COLORS.neutral,
          }),
          createSystemTableCell('Date', { bold: true, columnSpan: 2, fill: DOCX_REPORT_COLORS.neutral }),
          createSystemTableCell('Time', { bold: true, columnSpan: 2, fill: DOCX_REPORT_COLORS.neutral }),
          createSystemTableCell(report.status, {
            bold: true,
            color: statusTextColor,
            columnSpan: 3,
            fill: statusFill,
            rowSpan: 2,
            size: 30,
          }),
        ],
        { header: true, height: 520 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell('Inv.'),
          createSystemTableCell(report.inverter ?? ''),
          createSystemTableCell('String.'),
          createSystemTableCell(report.stringName ?? ''),
          createSystemTableCell(report.date ?? '', { columnSpan: 2 }),
          createSystemTableCell(report.time ?? '', { columnSpan: 2 }),
        ],
        { header: true, height: 400 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell('Number\nof\nmodules', { rowSpan: 2 }),
          createSystemTableCell('Irr.', { bold: true, fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell('Temp.', { bold: true, fill: DOCX_REPORT_COLORS.amber }),
          createSystemTableCell('Imp', { bold: true, columnSpan: 3, fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell('Vmp', { bold: true, columnSpan: 3, fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell('Isc', { bold: true, columnSpan: 4, fill: DOCX_REPORT_COLORS.headerSlate }),
        ],
        { height: 420 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell('(W/m^2)', { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell('(Deg C)', { fill: DOCX_REPORT_COLORS.amber }),
          createSystemTableCell('Measured\n(Amps)', { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell('Translated\nto STC', { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell('Nominal\nat STC', { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell('Measured\n(Volts)', { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell('Translated\nto STC', { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell('Nominal\nat STC', { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell('Measured\n(Amps)', { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell('Translated\nto STC', { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell('Nominal\nat STC', { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell('%', { fill: DOCX_REPORT_COLORS.headerSlate }),
        ],
        { height: 500 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell(formatDocxInteger(report.moduleCount), { fill: 'FFFFFF' }),
          createSystemTableCell(formatDocxNumberFixed(report.irradianceWm2, 0), { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell(formatDocxNumberFixed(report.temperatureC, 1), { fill: DOCX_REPORT_COLORS.amber }),
          createSystemTableCell(formatDocxNumberFixed(report.impMeasuredA, 2), { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell(formatDocxNumberFixed(report.impTranslatedA, 2), { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell(formatDocxNumberFixed(report.impNominalA, 2), { fill: DOCX_REPORT_COLORS.yellow }),
          createSystemTableCell(formatDocxNumberFixed(report.vmpMeasuredV, 0), { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell(formatDocxNumberFixed(report.vmpTranslatedV, 0), { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell(formatDocxNumberFixed(report.vmpNominalV, 0), { fill: DOCX_REPORT_COLORS.green }),
          createSystemTableCell(formatDocxNumberFixed(report.iscMeasuredA, 2), { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell(formatDocxNumberFixed(report.iscTranslatedA, 2), { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell(formatDocxNumberFixed(report.iscNominalA, 2), { fill: DOCX_REPORT_COLORS.headerSlate }),
          createSystemTableCell(formatDocxPercent(report.iscPercent, 1), { fill: DOCX_REPORT_COLORS.headerSlate }),
        ],
        { height: 420 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell('Voc', { bold: true, columnSpan: 4, fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell('Pmax', { bold: true, columnSpan: 3, fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell('PF', { bold: true, fill: DOCX_REPORT_COLORS.sky, rowSpan: 2 }),
          createSystemTableCell('FF', { bold: true, fill: DOCX_REPORT_COLORS.sky, rowSpan: 2 }),
          createSystemTableCell('Tol +\n( % )', { bold: true, fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell('Tol -\n( % )', { bold: true, columnSpan: 2, fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell('Deviation\nfrom\nexpected', { bold: true, fill: DOCX_REPORT_COLORS.sky }),
        ],
        { height: 420 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell('Measured\n(Volts)', { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell('Translated\nto STC', { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell('Nominal\nat STC', { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell('%', { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell('Measured\n(Watts)', { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell('Translated\nto STC', { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell('Nominal\nat STC', { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell('Equipment', { fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell('Equipment + Degradation', { columnSpan: 2, fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell('(%)', { fill: DOCX_REPORT_COLORS.sky }),
        ],
        { height: 500 },
      ),
      createDocxTableRow(
        [
          createSystemTableCell(formatDocxNumberFixed(report.vocMeasuredV, 0), { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell(formatDocxNumberFixed(report.vocTranslatedV, 0), { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell(formatDocxNumberFixed(report.vocNominalV, 0), { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell(formatDocxPercent(report.vocPercent, 1), { fill: DOCX_REPORT_COLORS.orange }),
          createSystemTableCell(formatDocxNumberFixed(report.pmaxMeasuredW, 0), { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell(formatDocxNumberFixed(report.pmaxTranslatedW, 0), { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell(formatDocxNumberFixed(report.pmaxNominalW, 0), { fill: DOCX_REPORT_COLORS.red }),
          createSystemTableCell(formatDocxPercent(report.pfPercent, 1), { fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell(formatDocxPercent(report.ffPercent, 1), { fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell(formatDocxPercent(report.tolPlusPercent, 1), { fill: DOCX_REPORT_COLORS.sky }),
          createSystemTableCell(formatDocxPercent(report.tolMinusPercent, 1), {
            columnSpan: 2,
            fill: DOCX_REPORT_COLORS.sky,
          }),
          createSystemTableCell(formatDocxPercent(report.deviationPercent, 1), { fill: DOCX_REPORT_COLORS.sky }),
        ],
        { height: 420 },
      ),
    ],
  })
}

function createDocxChartPairTable(record: MeasurementRecord, pvModule: PvModule | null) {
  const columnWidth = Math.floor(DOCX_OVERVIEW_BODY_WIDTH_DXA / 2)

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_OVERVIEW_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [columnWidth, DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('I-V Curve', {
            alignment: AlignmentType.CENTER,
            bold: true,
            fill: DOCX_REPORT_COLORS.neutral,
            width: columnWidth,
          }),
          createReportTableCell('P-V Curve', {
            alignment: AlignmentType.CENTER,
            bold: true,
            fill: DOCX_REPORT_COLORS.neutral,
            width: DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth,
          }),
        ],
        { header: true, height: 360 },
      ),
      createDocxTableRow(
        [
          createDocxChartCell(record, pvModule, 'amps', columnWidth),
          createDocxChartCell(record, pvModule, 'watts', DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth),
        ],
        { height: 2400 },
      ),
    ],
  })
}

function createDocxChartCell(
  record: MeasurementRecord,
  pvModule: PvModule | null,
  yField: 'amps' | 'watts',
  width: number,
) {
  const chartSvg = createCurveChartSvg(record, pvModule, yField)
  const title = yField === 'amps' ? 'I-V Curve' : 'P-V Curve'

  return new TableCell({
    borders: createDocxCellBorders(),
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    children: [
      chartSvg
        ? new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                type: 'svg',
                data: Buffer.from(chartSvg, 'utf8'),
                fallback: {
                  type: 'png',
                  data: TRANSPARENT_PNG,
                },
                transformation: {
                  width: 360,
                  height: 180,
                },
              }),
            ],
          })
        : createDocxParagraph(`No ${title} data: ${getStcConversionProblem(record, pvModule) ?? 'invalid STC data'}`),
    ],
  })
}

function createDocxRecordImageSlotsTable(slots: Array<RecordImageSlot | null>) {
  const columnWidth = Math.floor(DOCX_OVERVIEW_BODY_WIDTH_DXA / 2)

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_OVERVIEW_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [columnWidth, DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createReportTableCell('Thermal', {
            alignment: AlignmentType.CENTER,
            bold: true,
            fill: DOCX_REPORT_COLORS.neutral,
            width: columnWidth,
          }),
          createReportTableCell('Visible', {
            alignment: AlignmentType.CENTER,
            bold: true,
            fill: DOCX_REPORT_COLORS.neutral,
            width: DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth,
          }),
        ],
        { header: true, height: 360 },
      ),
      createDocxTableRow(
        [
          createDocxRecordImageCell(slots[0] ?? null, columnWidth),
          createDocxRecordImageCell(slots[1] ?? null, DOCX_OVERVIEW_BODY_WIDTH_DXA - columnWidth),
        ],
        { height: 2500 },
      ),
    ],
  })
}

function createDocxRecordImageCell(slot: RecordImageSlot | null, width: number) {
  const imageSvg = slot ? createAnnotatedRecordImageSvg(slot) : null

  return new TableCell({
    borders: createDocxCellBorders(),
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    children: [
      imageSvg
        ? new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new ImageRun({
                type: 'svg',
                data: Buffer.from(imageSvg, 'utf8'),
                fallback: {
                  type: 'png',
                  data: TRANSPARENT_PNG,
                },
                transformation: {
                  width: 360,
                  height: 210,
                },
              }),
            ],
          })
        : createDocxParagraph('No image selected', { alignment: AlignmentType.CENTER }),
    ],
  })
}

function createAnnotatedRecordImageSvg(slot: RecordImageSlot) {
  const width = 720
  const height = 420
  const zoom = typeof slot.zoom === 'number' && Number.isFinite(slot.zoom) ? Math.min(4, Math.max(0.5, slot.zoom)) : 1
  const strokeWidth = normalizeDocxRecordImageStrokeWidth(slot.strokeWidth)
  const imageWidth = width * zoom
  const imageHeight = height * zoom
  const imageX = (width - imageWidth) / 2 + slot.offsetX * width
  const imageY = (height - imageHeight) / 2 + slot.offsetY * height
  const shapes = slot.shapes
    .map((shape) => createDocxRecordImageShapeSvg(shape, imageX, imageY, imageWidth, imageHeight, strokeWidth))
    .join('\n  ')

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect width="${width}" height="${height}" fill="#f8fafc"/>
  <image href="${escapeXml(slot.dataUrl)}" x="${imageX.toFixed(1)}" y="${imageY.toFixed(1)}" width="${imageWidth.toFixed(1)}" height="${imageHeight.toFixed(1)}" preserveAspectRatio="xMidYMid meet"/>
  ${shapes}
</svg>`
}

function createDocxRecordImageShapeSvg(
  shape: RecordImageShape,
  imageX: number,
  imageY: number,
  imageWidth: number,
  imageHeight: number,
  strokeWidth: number,
) {
  const color = normalizeDocxRecordImageColor(shape.color)

  if (shape.type === 'rectangle' && shape.points.length >= 2) {
    const [startPoint, endPoint] = shape.points
    const x = imageX + Math.min(startPoint.x, endPoint.x) * imageWidth
    const y = imageY + Math.min(startPoint.y, endPoint.y) * imageHeight
    const width = Math.abs(startPoint.x - endPoint.x) * imageWidth
    const height = Math.abs(startPoint.y - endPoint.y) * imageHeight

    return `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${width.toFixed(1)}" height="${height.toFixed(1)}" fill="none" stroke="${color}" stroke-width="${strokeWidth}"/>`
  }

  if (shape.type !== 'polygon' || shape.points.length === 0) {
    return ''
  }

  const points = shape.points
    .map((point) => `${(imageX + point.x * imageWidth).toFixed(1)},${(imageY + point.y * imageHeight).toFixed(1)}`)
    .join(' ')

  return [
    shape.points.length >= 3 ? `<polygon points="${points}" fill="none" stroke="${color}" stroke-width="${strokeWidth}"/>` : '',
    shape.points.length >= 2 ? `<polyline points="${points}" fill="none" stroke="${color}" stroke-width="${strokeWidth}" stroke-dasharray="8 6"/>` : '',
  ]
    .filter(Boolean)
    .join('\n  ')
}

function getDocxRecordImageSlots(
  folderPath: string | null,
  record: MeasurementRecord,
  recordImages: RecordImageMap,
): Array<RecordImageSlot | null> {
  const key = getDocxRecordImageKey(folderPath, record)
  const slots = recordImages[key]

  if (!Array.isArray(slots)) {
    return [null, null]
  }

  return [normalizeDocxRecordImageSlot(slots[0]), normalizeDocxRecordImageSlot(slots[1])]
}

function getDocxRecordImageKey(folderPath: string | null, record: MeasurementRecord) {
  return `${folderPath ?? ''}::${record.relativePath || record.sourcePath}`
}

function normalizeDocxRecordImageSlot(slot: RecordImageSlot | null | undefined): RecordImageSlot | null {
  if (!slot?.dataUrl?.startsWith('data:image/')) {
    return null
  }

  const legacyPolygon = normalizeDocxRecordImagePolygon(slot.polygon)
  const shapes = normalizeDocxRecordImageShapes(slot.shapes)

  return {
    dataUrl: slot.dataUrl,
    fit: 'contain',
    name: slot.name ?? 'Image',
    offsetX: typeof slot.offsetX === 'number' && Number.isFinite(slot.offsetX) ? Math.min(1, Math.max(-1, slot.offsetX)) : 0,
    offsetY: typeof slot.offsetY === 'number' && Number.isFinite(slot.offsetY) ? Math.min(1, Math.max(-1, slot.offsetY)) : 0,
    polygon: [],
    shapes: shapes.length > 0 ? shapes : legacyPolygon.length > 0 ? [{ color: '#16a34a', points: legacyPolygon, type: 'polygon' }] : [],
    strokeWidth: normalizeDocxRecordImageStrokeWidth(slot.strokeWidth),
    zoom: typeof slot.zoom === 'number' && Number.isFinite(slot.zoom) ? Math.min(4, Math.max(0.5, slot.zoom)) : 1,
  }
}

function normalizeDocxRecordImageShapes(source: unknown): RecordImageShape[] {
  if (!Array.isArray(source)) {
    return []
  }

  return source
    .map((shape) => {
      if (!shape || typeof shape !== 'object' || Array.isArray(shape)) {
        return null
      }

      const entry = shape as Partial<RecordImageShape>
      const type = entry.type === 'rectangle' ? 'rectangle' : entry.type === 'polygon' ? 'polygon' : null
      const points = normalizeDocxRecordImagePolygon(entry.points)

      if (!type || points.length === 0 || (type === 'rectangle' && points.length < 2)) {
        return null
      }

      return {
        color: normalizeDocxRecordImageColor(entry.color),
        points: type === 'rectangle' ? points.slice(0, 2) : points,
        type,
      }
    })
    .filter((shape): shape is RecordImageShape => shape !== null)
}

function normalizeDocxRecordImagePolygon(source: unknown) {
  if (!Array.isArray(source)) {
    return []
  }

  return source
    .map((point) => {
      const x = typeof point?.x === 'number' && Number.isFinite(point.x) ? Math.min(1, Math.max(0, point.x)) : null
      const y = typeof point?.y === 'number' && Number.isFinite(point.y) ? Math.min(1, Math.max(0, point.y)) : null

      return x === null || y === null ? null : { x, y }
    })
    .filter((point): point is { x: number; y: number } => point !== null)
}

function normalizeDocxRecordImageColor(value: unknown) {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value : '#16a34a'
}

function normalizeDocxRecordImageStrokeWidth(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) ? Math.min(8, Math.max(1, value)) : 2
}

function createCurveChartSvg(
  record: MeasurementRecord,
  pvModule: PvModule | null,
  yField: 'amps' | 'watts',
) {
  const measurements = convertIvMeasurementsToStc(record, pvModule)
  const referenceMeasurements = buildPvModuleReferenceMeasurements(pvModule)

  if (measurements.length === 0) {
    return null
  }

  const width = 640
  const height = 320
  const left = 54
  const right = width - 24
  const top = 24
  const bottom = height - 42
  const scaleMeasurements = [...measurements, ...referenceMeasurements]
  const maxVolts = Math.max(...scaleMeasurements.map((point) => point.volts))
  const maxY = Math.max(...scaleMeasurements.map((point) => point[yField]))

  if (maxVolts <= 0 || maxY <= 0) {
    return null
  }

  const tickCount = 5
  const strokeColor = yField === 'amps' ? '#0369a1' : '#d97706'
  const yAxisLabel = yField === 'amps' ? 'A' : 'W'
  const title = yField === 'amps' ? 'I-V Curve' : 'P-V Curve'
  const points = measurements
    .map((point) => {
      const x = left + (point.volts / maxVolts) * (right - left)
      const y = bottom - (point[yField] / maxY) * (bottom - top)

      return `${x.toFixed(2)},${y.toFixed(2)}`
    })
    .join(' ')
  const referencePoints = referenceMeasurements
    .map((point) => {
      const x = left + (point.volts / maxVolts) * (right - left)
      const y = bottom - (point[yField] / maxY) * (bottom - top)

      return `${x.toFixed(2)},${y.toFixed(2)}`
    })
    .join(' ')
  const xGrid = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = (maxVolts / tickCount) * index
    const x = left + (value / maxVolts) * (right - left)
    const label = escapeXml(formatDocxChartTick(value))

    return `<line x1="${x.toFixed(2)}" y1="${top}" x2="${x.toFixed(2)}" y2="${bottom}" stroke="#e2e8f0" stroke-width="1"/><text x="${x.toFixed(2)}" y="${bottom + 20}" text-anchor="middle" fill="#475569" font-size="11" font-family="Arial">${label}</text>`
  }).join('')
  const yGrid = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = (maxY / tickCount) * index
    const y = bottom - (value / maxY) * (bottom - top)
    const label = escapeXml(formatDocxChartTick(value))

    return `<line x1="${left}" y1="${y.toFixed(2)}" x2="${right}" y2="${y.toFixed(2)}" stroke="#e2e8f0" stroke-width="1"/><text x="${left - 8}" y="${(y + 4).toFixed(2)}" text-anchor="end" fill="#475569" font-size="11" font-family="Arial">${label}</text>`
  }).join('')

  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
  <rect x="0" y="0" width="${width}" height="${height}" fill="#f8fafc"/>
  <text x="${left}" y="18" fill="#334155" font-size="14" font-weight="700" font-family="Arial">${escapeXml(title)}</text>
  <line x1="${right - 126}" y1="14" x2="${right - 96}" y2="14" stroke="${strokeColor}" stroke-width="3" stroke-linecap="round"/>
  <text x="${right - 88}" y="18" fill="#475569" font-size="11" font-family="Arial">Measured</text>
  ${referencePoints ? `<line x1="${right - 126}" y1="30" x2="${right - 96}" y2="30" stroke="#166534" stroke-width="2.4" stroke-linecap="round" stroke-dasharray="8 6"/><text x="${right - 88}" y="34" fill="#475569" font-size="11" font-family="Arial">Reference</text>` : ''}
  <rect x="${left}" y="${top}" width="${right - left}" height="${bottom - top}" fill="#ffffff" stroke="#cbd5e1"/>
  ${xGrid}
  ${yGrid}
  <line x1="${left}" y1="${bottom}" x2="${right}" y2="${bottom}" stroke="#64748b" stroke-width="1.4"/>
  <line x1="${left}" y1="${top}" x2="${left}" y2="${bottom}" stroke="#64748b" stroke-width="1.4"/>
  ${referencePoints ? `<polyline fill="none" stroke="#166534" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" stroke-dasharray="8 6" points="${referencePoints}"/>` : ''}
  <polyline fill="none" stroke="${strokeColor}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" points="${points}"/>
  <text x="${right}" y="${height - 8}" fill="#475569" font-size="16" font-weight="700" font-family="Arial">V</text>
  <text x="8" y="${top + 8}" fill="#475569" font-size="16" font-weight="700" font-family="Arial">${yAxisLabel}</text>
</svg>`
}

function createToleranceCalculationDetailsTable(toleranceRows: ToleranceRow[]) {
  const labelWidth = 1400
  const baseValueWidth = Math.floor((DOCX_BODY_WIDTH_DXA - labelWidth) / TOLERANCE_COLUMNS.length)
  const valueWidths = TOLERANCE_COLUMNS.map((_, index) =>
    index === TOLERANCE_COLUMNS.length - 1
      ? DOCX_BODY_WIDTH_DXA - labelWidth - baseValueWidth * (TOLERANCE_COLUMNS.length - 1)
      : baseValueWidth,
  )

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [labelWidth, ...valueWidths],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createToleranceTableCell('', { fill: DOCX_REPORT_COLORS.headerSlateDark, color: 'FFFFFF', bold: true, width: labelWidth }),
          ...TOLERANCE_COLUMNS.map((column, columnIndex) =>
            createToleranceTableCell(column.label, {
              fill: DOCX_REPORT_COLORS.headerSlateDark,
              color: 'FFFFFF',
              bold: true,
              alignment: AlignmentType.LEFT,
              width: valueWidths[columnIndex],
            }),
          ),
        ],
        { header: true, height: 900 },
      ),
      ...toleranceRows.map(
        (row, rowIndex) => {
          const rowFill = rowIndex % 2 === 0 ? 'FFFFFF' : DOCX_REPORT_COLORS.neutralBand

          return createDocxTableRow(
            [
              createToleranceTableCell(row.label, { alignment: AlignmentType.LEFT, fill: rowFill, size: 26, width: labelWidth }),
              ...TOLERANCE_COLUMNS.map((column, columnIndex) => {
                const isTotalColumn = column.key === TOLERANCE_TOTAL_KEY
                const isHighlightedTotal = row.label === 'Tol-' && isTotalColumn

                return createToleranceTableCell(
                  formatTolerancePercent(getToleranceCellValue(row, column.key), column.digits),
                  {
                    alignment: AlignmentType.RIGHT,
                    fill: isHighlightedTotal ? DOCX_REPORT_COLORS.amber : rowFill,
                    size: 24,
                    width: valueWidths[columnIndex],
                  },
                )
              }),
            ],
            { height: 420 },
          )
        },
      ),
    ],
  })
}

function createDocxFormulaReferenceTable() {
  const labelWidth = 2300
  const formulaWidth = DOCX_BODY_WIDTH_DXA - labelWidth

  return new Table({
    alignment: AlignmentType.CENTER,
    borders: createDocxTableBorders(),
    width: { size: DOCX_BODY_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [labelWidth, formulaWidth],
    layout: TableLayoutType.FIXED,
    rows: [
      createDocxTableRow(
        [
          createFormulaLabelCell('The Fill Factor is\ncalculated as this\nformula:', labelWidth),
          createFormulaMathCell([createFillFactorEquation()], formulaWidth),
        ],
        { height: 1320 },
      ),
      createDocxTableRow(
        [
          createFormulaLabelCell('The Performance\nFactor is calculated\nas this formula:', labelWidth),
          createFormulaMathCell([createPerformanceFactorEquation()], formulaWidth),
        ],
        { height: 1320 },
      ),
      createDocxTableRow(
        [
          createFormulaLabelCell('All are calculated\nusing this formula\nfor translation:', labelWidth),
          createFormulaMathCell(
            [
              createIscTranslationEquation(),
              createVocTranslationEquation(),
              createImpTranslationEquation(),
              createVmpTranslationEquation(),
              createPmpTranslationEquation(),
            ],
            formulaWidth,
            { paragraphSpacingAfter: 420 },
          ),
        ],
        { height: 7100 },
      ),
    ],
  })
}

function createFormulaLabelCell(text: string, width: number) {
  return createReportTableCell(text, {
    alignment: AlignmentType.CENTER,
    noProof: true,
    size: 22,
    width,
  })
}

function createFormulaMathCell(
  formulas: DocxMath[],
  width: number,
  options: { paragraphSpacingAfter?: number } = {},
) {
  return new TableCell({
    borders: createDocxCellBorders(),
    margins: { top: 120, bottom: 120, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    width: { size: width, type: WidthType.DXA },
    children: formulas.map(
      (formula, index) =>
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: {
            after: index === formulas.length - 1 ? 0 : options.paragraphSpacingAfter ?? 0,
            before: 0,
          },
          children: [formula],
        }),
    ),
  })
}

function createFillFactorEquation() {
  return createMathEquation([
    mathText('FF = '),
    mathFraction(
      [mathSub('V', 'max,meas'), mathText(' × '), mathSub('I', 'max,meas')],
      [mathSub('V', 'oc,meas'), mathText(' × '), mathSub('I', 'sc,meas')],
    ),
  ])
}

function createPerformanceFactorEquation() {
  return createMathEquation([
    mathText('PF = '),
    mathFraction(
      [mathSub('P', 'max,meas trans to STC')],
      [mathSub('P', 'max,predicted')],
    ),
  ])
}

function createIscTranslationEquation() {
  return createMathEquation([
    mathSub('I', 'sc,trans'),
    mathText(' = '),
    mathFraction(
      [
        mathSub('I', 'sc,meas'),
        mathText(' × '),
        mathFraction([mathSub('E', 'trans')], [mathSub('E', 'meas')]),
      ],
      createTemperatureDenominator('α', 'Isc'),
    ),
  ])
}

function createVocTranslationEquation() {
  return createMathEquation([
    mathSub('V', 'oc,trans'),
    mathText(' = '),
    mathFraction([mathSub('V', 'oc,meas')], createTemperatureDenominator('β', 'Voc')),
  ])
}

function createImpTranslationEquation() {
  return createMathEquation([
    mathSub('I', 'mp,trans'),
    mathText(' = '),
    mathSub('I', 'mp,meas'),
    mathText(' × '),
    mathFraction([mathSub('E', 'trans')], [mathSub('E', 'meas')]),
  ])
}

function createVmpTranslationEquation() {
  return createMathEquation([
    mathSub('V', 'mp,trans'),
    mathText(' = '),
    mathFraction([mathSub('V', 'mp,meas')], createTemperatureDenominator('γ', 'mpp')),
  ])
}

function createPmpTranslationEquation() {
  return createMathEquation([
    mathSub('P', 'mp,trans'),
    mathText(' = '),
    mathFraction(
      [
        mathSub('P', 'mp,meas'),
        mathText(' × '),
        mathFraction([mathSub('E', 'trans')], [mathSub('E', 'meas')]),
      ],
      createTemperatureDenominator('γ', 'mpp'),
    ),
  ])
}

function createTemperatureDenominator(coefficient: string, coefficientSubscript: string): MathComponent[] {
  return [
    mathText('1 + '),
    mathFraction([mathSub(coefficient, coefficientSubscript)], [mathText('100')]),
    mathText(' × '),
    new MathRoundBrackets({
      children: [mathSub('T', 'meas'), mathText(' − '), mathSub('T', 'trans')],
    }),
  ]
}

function createMathEquation(children: MathComponent[]) {
  return new DocxMath({ children })
}

function mathText(text: string) {
  return new MathRun(text)
}

function mathSub(text: string, subScript: string) {
  return new MathSubScript({
    children: [mathText(text)],
    subScript: [mathText(subScript)],
  })
}

function mathFraction(numerator: MathComponent[], denominator: MathComponent[]) {
  return new MathFraction({ numerator, denominator })
}

function createToleranceTableCell(
  text: string,
  options: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
    bold?: boolean
    color?: string
    fill?: string
    size?: number
    width?: number
  } = {},
) {
  return new TableCell({
    borders: createDocxCellBorders(),
    shading: options.fill ? { fill: options.fill } : undefined,
    margins: { top: 70, bottom: 70, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
    children: [
      new Paragraph({
        alignment: options.alignment ?? AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
        children: createDocxTextRuns(text, {
          bold: options.bold,
          color: options.color,
          size: options.size,
        }),
      }),
    ],
  })
}

function createReportTableCell(
  text: string,
  options: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
    bold?: boolean
    color?: string
    columnSpan?: number
    fill?: string
    noProof?: boolean
    rowSpan?: number
    size?: number
    width?: number
  } = {},
) {
  return new TableCell({
    borders: createDocxCellBorders(),
    margins: { top: 80, bottom: 80, left: 80, right: 80 },
    columnSpan: options.columnSpan,
    rowSpan: options.rowSpan,
    shading: options.fill ? { fill: options.fill } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    width: options.width ? { size: options.width, type: WidthType.DXA } : undefined,
    children: [
      new Paragraph({
        alignment: options.alignment ?? AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
        children: createDocxTextRuns(text, {
          bold: options.bold,
          color: options.color,
          noProof: options.noProof,
          size: options.size ?? 18,
        }),
      }),
    ],
  })
}

function createDocxTableRow(
  children: TableCell[],
  options: {
    header?: boolean
    height?: number
  } = {},
) {
  return new TableRow({
    cantSplit: true,
    children,
    height: options.height ? { rule: HeightRule.ATLEAST, value: options.height } : undefined,
    tableHeader: options.header,
  })
}

function createDocxCellBorders() {
  const border = {
    color: DOCX_REPORT_COLORS.border,
    size: 6,
    style: BorderStyle.SINGLE,
  }

  return {
    bottom: border,
    left: border,
    right: border,
    top: border,
  }
}

function createDocxNoCellBorders() {
  const border = {
    color: 'FFFFFF',
    size: 0,
    style: BorderStyle.NONE,
  }

  return {
    bottom: border,
    left: border,
    right: border,
    top: border,
  }
}

function createDocxTableBorders() {
  const border = {
    color: DOCX_REPORT_COLORS.border,
    size: 8,
    style: BorderStyle.SINGLE,
  }

  return {
    bottom: border,
    insideHorizontal: border,
    insideVertical: border,
    left: border,
    right: border,
    top: border,
  }
}

function createDocxNoTableBorders() {
  const border = {
    color: 'FFFFFF',
    size: 0,
    style: BorderStyle.NONE,
  }

  return {
    bottom: border,
    insideHorizontal: border,
    insideVertical: border,
    left: border,
    right: border,
    top: border,
  }
}

function createDocxParagraph(
  text: string,
  options: {
    alignment?: (typeof AlignmentType)[keyof typeof AlignmentType]
    bold?: boolean
    heading?: (typeof HeadingLevel)[keyof typeof HeadingLevel]
    size?: number
  } = {},
) {
  return new Paragraph({
    alignment: options.alignment,
    heading: options.heading,
    children: createDocxTextRuns(text, {
      bold: options.bold,
      size: options.size,
    }),
  })
}

function createPageBreak() {
  return new Paragraph({
    children: [new PageBreak()],
  })
}

function createDocxTextRuns(
  text: string,
  options: {
    bold?: boolean
    color?: string
    italics?: boolean
    noProof?: boolean
    size?: number
    underline?: boolean
  } = {},
) {
  return normalizeDocxText(text).split('\n').map(
    (line, index) =>
      new TextRun({
        text: line,
        bold: options.bold,
        break: index === 0 ? undefined : 1,
        color: options.color,
        font: DOCX_REPORT_FONT_ATTRIBUTES,
        italics: options.italics,
        noProof: options.noProof,
        size: options.size,
        underline: options.underline ? { type: UnderlineType.SINGLE } : undefined,
      }),
  )
}

function buildDocxSummaryRows(
  records: MeasurementRecord[],
  pvModules: PvModule[],
  selectedPvModuleId: string,
  toleranceRows: ToleranceRow[],
): DocxSummaryRow[] {
  const selectedModule = pvModules.find((moduleItem) => moduleItem.id === selectedPvModuleId) ?? null
  const rows = new Map<string, DocxSummaryRow>()

  records.forEach((record) => {
    const inverter = getDocxReportInverterLabel(record)
    const currentRow =
      rows.get(inverter) ??
      ({
        degradationRatePercent: 0,
        failCount: 0,
        inverter,
        passCount: 0,
        totalStrings: 0,
      } satisfies DocxSummaryRow)
    const moduleItem = selectedModule ?? findPvModuleForRecord(record, pvModules)
    const metrics = buildDocxReportMetrics(record, moduleItem, toleranceRows)

    currentRow.totalStrings += 1

    if (metrics.status === 'PASS') {
      currentRow.passCount += 1
    } else {
      currentRow.failCount += 1
    }

    rows.set(inverter, currentRow)
  })

  return [...rows.values()]
    .map((row) => ({
      ...row,
      degradationRatePercent: row.totalStrings > 0 ? (row.failCount / row.totalStrings) * 100 : 0,
    }))
    .sort((left, right) => naturalCollator.compare(left.inverter, right.inverter))
}

function buildDocxDetailRows(
  records: MeasurementRecord[],
  pvModules: PvModule[],
  selectedPvModuleId: string,
  toleranceRows: ToleranceRow[],
): DocxDetailRow[] {
  const selectedModule = pvModules.find((moduleItem) => moduleItem.id === selectedPvModuleId) ?? null

  return sortDocxRecordsByHierarchy(records).map((record) => {
    const moduleItem = selectedModule ?? findPvModuleForRecord(record, pvModules)
    const metrics = buildDocxReportMetrics(record, moduleItem, toleranceRows)

    return {
      ...metrics,
      inverter: getDocxReportInverterLabel(record),
      moduleNumber: record.modulesInString,
      note: blankNoneText(record.measurementNote),
      stringName: getDocxStringKey(record),
    }
  })
}

function buildDocxReportMetrics(
  record: MeasurementRecord,
  pvModule: PvModule | null,
  toleranceRows: ToleranceRow[],
): DocxReportMetrics {
  const summary = record.measurementSummary
  const stcSummary = convertMeasurementSummaryToStc(record, pvModule)
  const moduleCount = record.modulesInString
  const tolPlusPercent = getDocxToleranceTotal(toleranceRows, 'Tol+')
  const tolMinusPercent = getDocxToleranceTotal(toleranceRows, 'Tol-')
  const iscTranslatedA = stcSummary?.iscA ?? null
  const vocTranslatedV = multiplyNullable(stcSummary?.vocV ?? null, moduleCount)
  const pmaxTranslatedW = multiplyNullable(stcSummary?.pmaxW ?? null, moduleCount)
  const iscNominalA = pvModule?.shortCircuitCurrentA ?? null
  const vocNominalV = multiplyNullable(pvModule?.openCircuitVoltageV ?? null, moduleCount)
  const pmaxNominalW = multiplyNullable(pvModule?.ratedMaximumPowerW ?? null, moduleCount)
  const pfPercent = ratioPercent(pmaxTranslatedW, pmaxNominalW)
  const ffPercent = ratioPercent(summary.pmaxW, multiplyNullable(summary.vocV, summary.iscA))

  return {
    deviationPercent: pfPercent === null ? null : Math.abs(100 - pfPercent),
    ffPercent,
    iscPercent: ratioPercent(iscTranslatedA, iscNominalA),
    pfPercent,
    status: getDocxReportStatus(pfPercent, tolPlusPercent, tolMinusPercent),
    tolMinusPercent,
    tolPlusPercent,
    vocPercent: ratioPercent(vocTranslatedV, vocNominalV),
  }
}

function buildDocxSystemReportData(
  record: MeasurementRecord,
  pvModule: PvModule | null,
  toleranceRows: ToleranceRow[],
): DocxSystemReportData {
  const summary = record.measurementSummary
  const stcSummary = convertMeasurementSummaryToStc(record, pvModule)
  const moduleCount = record.modulesInString
  const metrics = buildDocxReportMetrics(record, pvModule, toleranceRows)
  const impTranslatedA = stcSummary?.imppA ?? null
  const vmpTranslatedV = multiplyNullable(stcSummary?.vmppV ?? null, moduleCount)
  const iscTranslatedA = stcSummary?.iscA ?? null
  const vocTranslatedV = multiplyNullable(stcSummary?.vocV ?? null, moduleCount)
  const pmaxTranslatedW = multiplyNullable(stcSummary?.pmaxW ?? null, moduleCount)

  return {
    ...metrics,
    date: record.measurementDate,
    impMeasuredA: summary.imppA,
    impNominalA: pvModule?.maximumPowerCurrentA ?? null,
    impTranslatedA,
    inverter: getDocxReportInverterLabel(record),
    iscMeasuredA: summary.iscA,
    iscNominalA: pvModule?.shortCircuitCurrentA ?? null,
    iscTranslatedA,
    irradianceWm2: record.irradianceWm2,
    model: pvModule?.model ?? record.moduleModel,
    moduleCount,
    pmaxMeasuredW: summary.pmaxW,
    pmaxNominalW: multiplyNullable(pvModule?.ratedMaximumPowerW ?? null, moduleCount),
    pmaxTranslatedW,
    stringName: getDocxStringKey(record),
    temperatureC: record.cellTemperatureC,
    time: record.measurementTime,
    vmpMeasuredV: summary.vmppV,
    vmpNominalV: multiplyNullable(pvModule?.maximumPowerVoltageV ?? null, moduleCount),
    vmpTranslatedV,
    vocMeasuredV: summary.vocV,
    vocNominalV: multiplyNullable(pvModule?.openCircuitVoltageV ?? null, moduleCount),
    vocTranslatedV,
  }
}

function getDocxReportStatus(
  pfPercent: number | null,
  tolPlusPercent: number | null,
  tolMinusPercent: number | null,
): 'PASS' | 'FAIL' | 'N/A' {
  if (pfPercent === null || tolPlusPercent === null || tolMinusPercent === null) {
    return 'N/A'
  }

  return pfPercent <= 100 + tolPlusPercent && pfPercent >= 100 - tolMinusPercent ? 'PASS' : 'FAIL'
}

function getDocxToleranceTotal(rows: ToleranceRow[], label: string) {
  const row = rows.find((item) => item.label === label)

  return row ? getToleranceCellValue(row, TOLERANCE_TOTAL_KEY) : null
}

function sortDocxRecordsByHierarchy(records: MeasurementRecord[]) {
  return [...records].sort((left, right) => {
    return (
      naturalCollator.compare(left.systemGroup ?? 'Unknown X', right.systemGroup ?? 'Unknown X') ||
      naturalCollator.compare(left.inverter ?? 'Unknown inverter', right.inverter ?? 'Unknown inverter') ||
      naturalCollator.compare(getDocxStringKey(left), getDocxStringKey(right)) ||
      naturalCollator.compare(left.relativePath, right.relativePath)
    )
  })
}

function getDocxReportInverterLabel(record: MeasurementRecord) {
  const systemGroup = record.systemGroup?.trim()
  const inverter = record.inverter ?? 'Unknown inverter'

  return systemGroup ? `${systemGroup}.${inverter}` : inverter
}

function getDocxStringKey(record: MeasurementRecord) {
  return record.stringName ?? record.arrayLocation ?? record.relativePath
}

function blankNoneText(value: string | null | undefined) {
  const text = value?.trim() ?? ''
  const normalizedText = text.toLowerCase()

  return normalizedText === 'none' || normalizedText === '(none)' ? '' : text
}

function multiplyNullable(left: number | null, right: number | null) {
  return left === null || right === null ? null : left * right
}

function ratioPercent(value: number | null, base: number | null) {
  if (value === null || base === null || base === 0) {
    return null
  }

  return (value / base) * 100
}

function formatDocxPvModuleValue(value: number | null) {
  if (value === null || !Number.isFinite(value)) {
    return ''
  }

  return Number.isInteger(value) ? value.toFixed(0) : value.toString()
}

function formatDocxPercent(value: number | null, digits: number) {
  return value === null || !Number.isFinite(value) ? '' : `${value.toFixed(digits)}%`
}

function formatDocxNumberFixed(value: number | null, digits: number) {
  return value === null || !Number.isFinite(value) ? '' : value.toFixed(digits)
}

function formatDocxInteger(value: number | null) {
  return value === null || !Number.isFinite(value) ? '' : value.toFixed(0)
}

function formatDocxChartTick(value: number) {
  if (Math.abs(value) >= 100) {
    return value.toFixed(0)
  }

  if (Math.abs(value) >= 10) {
    return value.toFixed(1)
  }

  return value.toFixed(2)
}

function normalizeDocxProjectInfo(
  source: ProjectInfo | null | undefined,
  folderPath: string | null,
  record: MeasurementRecord,
): ProjectInfo {
  const fallbackProjectName = folderPath ? path.basename(folderPath) : record.station ?? ''

  return {
    projectName: getDocxProjectInfoText(source?.projectName, fallbackProjectName),
    investorName: getDocxProjectInfoText(source?.investorName, DEFAULT_PROJECT_INFO.investorName),
    investorNameEnglish: getDocxProjectInfoText(
      source?.investorNameEnglish,
      DEFAULT_PROJECT_INFO.investorNameEnglish,
    ),
    factoryOwnerName: getDocxProjectInfoText(source?.factoryOwnerName, DEFAULT_PROJECT_INFO.factoryOwnerName),
    factoryOwnerNameEnglish: getDocxProjectInfoText(
      source?.factoryOwnerNameEnglish,
      DEFAULT_PROJECT_INFO.factoryOwnerNameEnglish,
    ),
    reportTitle: getDocxProjectInfoText(source?.reportTitle, DEFAULT_PROJECT_INFO.reportTitle),
    measurementTitle: getDocxProjectInfoText(source?.measurementTitle, DEFAULT_PROJECT_INFO.measurementTitle),
    companyName: getDocxProjectInfoText(source?.companyName, DEFAULT_PROJECT_INFO.companyName),
    companyAddress: getDocxProjectInfoText(source?.companyAddress, DEFAULT_PROJECT_INFO.companyAddress),
    preparedBy: getDocxProjectInfoText(source?.preparedBy, DEFAULT_PROJECT_INFO.preparedBy),
    checkedBy: getDocxProjectInfoText(source?.checkedBy, DEFAULT_PROJECT_INFO.checkedBy),
    approvedBy: getDocxProjectInfoText(source?.approvedBy, DEFAULT_PROJECT_INFO.approvedBy),
    ownerApproval: getDocxProjectInfoText(source?.ownerApproval, DEFAULT_PROJECT_INFO.ownerApproval),
    consultantApproval: getDocxProjectInfoText(
      source?.consultantApproval,
      DEFAULT_PROJECT_INFO.consultantApproval,
    ),
    contractorEpcApproval: getDocxProjectInfoText(
      source?.contractorEpcApproval,
      DEFAULT_PROJECT_INFO.contractorEpcApproval,
    ),
    testerApproval: getDocxProjectInfoText(source?.testerApproval, DEFAULT_PROJECT_INFO.testerApproval),
    applicableStandards: getDocxProjectInfoText(
      source?.applicableStandards,
      DEFAULT_PROJECT_INFO.applicableStandards,
    ),
  }
}

function getDocxProjectInfoText(value: string | null | undefined, fallback: string) {
  const text = typeof value === 'string' ? value.trim() : ''

  return normalizeDocxText(text || fallback)
}

function formatEngineerLabel(name: string) {
  return normalizeDocxText(['Engineer', name.trim()].filter(Boolean).join(' '))
}

function splitDocxProjectInfoLines(value: string) {
  const lines = normalizeDocxText(value)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)

  return lines.length > 0 ? lines : ['']
}

function normalizeDocxText(value: string) {
  return repairVietnameseTextSpacing(value.normalize('NFC'))
}

function repairVietnameseTextSpacing(value: string) {
  const separator = /[\s\u00A0\u2000-\u200D\uFEFF]+/g

  return value
    .replace(new RegExp(`PHẦ${separator.source}N`, 'g'), 'PHẦN')
    .replace(new RegExp(`Phầ${separator.source}n`, 'g'), 'Phần')
    .replace(new RegExp(`phầ${separator.source}n`, 'g'), 'phần')
    .replace(new RegExp(`ĐẦ${separator.source}U`, 'g'), 'ĐẦU')
    .replace(new RegExp(`Đầ${separator.source}u`, 'g'), 'Đầu')
    .replace(new RegExp(`đầ${separator.source}u`, 'g'), 'đầu')
    .replace(new RegExp(`MÁ${separator.source}Y`, 'g'), 'MÁY')
    .replace(new RegExp(`Má${separator.source}y`, 'g'), 'Máy')
    .replace(new RegExp(`má${separator.source}y`, 'g'), 'máy')
}

function escapeXml(value: string) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

const TRANSPARENT_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=',
  'base64',
)

function buildReportFileName(folderPath: string | null, record: MeasurementRecord) {
  const projectName = folderPath ? path.basename(folderPath) : record.station

  return sanitizeReportFileName(`IV Report - ${projectName || 'Project'}`)
}

async function writeReportFileWithFallback(filePath: string, buffer: Buffer) {
  try {
    await writeFile(filePath, buffer)

    return { filePath, renamed: false }
  } catch (error) {
    if (!isLockedFileError(error)) {
      throw error
    }

    const fallbackFilePath = await writeReportFileToNextAvailableName(filePath, buffer)

    return { filePath: fallbackFilePath, renamed: true }
  }
}

async function writeReportFileToNextAvailableName(filePath: string, buffer: Buffer) {
  const parsedPath = path.parse(filePath)
  const extension = parsedPath.ext || '.docx'
  const baseName = parsedPath.name || 'IV Report'

  for (let index = 2; index <= 999; index += 1) {
    const candidatePath = path.join(parsedPath.dir, `${baseName} (${index})${extension}`)

    try {
      await writeFile(candidatePath, buffer, { flag: 'wx' })

      return candidatePath
    } catch (error) {
      if (isFileExistsError(error) || isBusyFileError(error)) {
        continue
      }

      throw error
    }
  }

  throw new Error(`Unable to find an available file name near ${filePath}`)
}

function isLockedFileError(error: unknown) {
  const code = getFileSystemErrorCode(error)

  return code === 'EBUSY' || code === 'EPERM' || code === 'EACCES'
}

function isBusyFileError(error: unknown) {
  return getFileSystemErrorCode(error) === 'EBUSY'
}

function isFileExistsError(error: unknown) {
  return getFileSystemErrorCode(error) === 'EEXIST'
}

function getFileSystemErrorCode(error: unknown) {
  return error && typeof error === 'object' && 'code' in error
    ? String((error as { code?: unknown }).code)
    : ''
}

function sanitizeReportFileName(fileName: string) {
  return fileName.replace(/[<>:"/\\|?*]+/g, '_').replace(/\s+/g, ' ').trim()
}

function sanitizeFileName(fileName: string) {
  return fileName.replace(/[<>:"/\\|?*]+/g, '_').replace(/\s+/g, '_')
}

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
    win = null
  }
})

app.on('activate', () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

app.whenReady().then(createWindow)
