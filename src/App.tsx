import { useEffect, useMemo, useRef, useState } from 'react'
import type { FormEvent, MouseEvent, PointerEvent, ReactNode, WheelEvent } from 'react'
import type {
  DataImportResult,
  IvPoint,
  MeasurementRecord,
  ProjectInfo,
  ProjectInfoLibraryItem,
  PvModule,
  RecordImageMap,
  RecordImagePoint,
  RecordImageShape,
  RecordImageSlot,
  WordExportPayload,
  WordExportProgress,
  WordExportResult,
} from './types/data'
import type {
  ChatAppContext,
  ChatCommand,
  ChatCompletionRequest,
  ChatCompletionResult,
  ChatHistoryItem,
  ChatImageAttachment,
  ChatPvModulePayload,
  CodexSetupActionResult,
  CodexSetupStatus,
} from './types/chat'
import {
  buildPvModuleReferenceMeasurements,
  convertMeasurementSummaryToStc,
  convertIvMeasurementsToStc,
  findPvModuleForRecord,
  getStcConversionProblem,
} from './lib/stc'
import {
  cloneToleranceRows,
  getToleranceCellValue,
  normalizeToleranceRows,
  TOLERANCE_COLUMNS,
  TOLERANCE_TOTAL_KEY,
  type ToleranceRow,
} from './lib/tolerance'

type AppView = 'home' | 'project-info' | 'pv-module'

interface ProjectModuleCacheEntry {
  moduleId: string
  moduleModel: string
}

type ProjectModuleCache = Record<string, ProjectModuleCacheEntry>
type RecordNoteOverrides = Record<string, string>
type RecordImageOverrides = RecordImageMap

interface AppChatMessage extends ChatHistoryItem {
  id: string
}

interface AppSessionState {
  version: 1
  activeView: AppView
  folderPath: string | null
  isChatOpen: boolean
  selectedInverter: string
  selectedPvModuleId: string
  selectedRecordPath: string
  selectedRecordRelativePath: string
  selectedSystemGroup: string
}

const chatImageMaxCount = 4
const chatImageMaxBytes = 10 * 1024 * 1024
type CodexSetupActionChannel =
  | 'codex:install'
  | 'codex:login'
  | 'codex:reset-thread'
  | 'codex:select-workspace'
  | 'codex:set-workspace'
  | 'codex:open-workspace'
  | 'codex:reset-workspace'

const projectModuleCacheKey = 'eb-iv.project-module-cache'
const projectInfoStorageKey = 'eb-iv.project-info'
const projectInfoLibraryStorageKey = 'eb-iv.project-info-library'
const recordNoteOverridesStorageKey = 'eb-iv.record-note-overrides'
const recordImageOverridesStorageKey = 'eb-iv.record-image-overrides'
const appSessionStorageKey = 'eb-iv.app-session'
const toleranceRowsStorageKey = 'eb-iv.tolerance-rows'
const defaultFirstYearDegradationPercent = 2
const defaultAnnualDegradationPercent = 0.55
const restoreImportTimeoutMs = 45_000
const recordImageMinZoom = 0.5
const recordImageMaxZoom = 4
const recordImageMinStrokeWidth = 1
const recordImageMaxStrokeWidth = 8
const defaultRecordImageShapeColor = '#16a34a'

const defaultProjectInfo: ProjectInfo = {
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

const referenceProjectInfo: ProjectInfo = {
  ...defaultProjectInfo,
  projectName: 'I-V report sample',
  investorName: 'CÔNG TY CỔ PHẦN ĐẦU TƯ CME SOLAR',
  investorNameEnglish: 'CME SOLAR INVESTMENT JOINT STOCK COMPANY',
  factoryOwnerName: 'CÔNG TY TNHH VIET-SCREW',
  factoryOwnerNameEnglish: 'VIET-SCREW CO., LTD.',
}

const seededProjectInfoLibrary: ProjectInfoLibraryItem[] = [
  {
    ...referenceProjectInfo,
    id: 'reference-iv-report',
    label: referenceProjectInfo.factoryOwnerName,
    updatedAt: '2026-05-28T00:00:00.000Z',
  },
]

function App() {
  const [savedSession] = useState<AppSessionState | null>(() => readAppSession())
  const sessionRestoreAttemptedRef = useRef(false)
  const [importResult, setImportResult] = useState<DataImportResult | null>(null)
  const [isSelecting, setIsSelecting] = useState(false)
  const [isExporting, setIsExporting] = useState(false)
  const [exportProgress, setExportProgress] = useState<WordExportProgress | null>(null)
  const [activeView, setActiveView] = useState<AppView>(savedSession?.activeView ?? 'home')
  const [isModuleLibraryLoading, setIsModuleLibraryLoading] = useState(false)
  const [exportedFilePath, setExportedFilePath] = useState<string | null>(null)
  const [exportStatus, setExportStatus] = useState<string | null>(null)
  const [moduleLibraryStatus, setModuleLibraryStatus] = useState<string | null>(null)
  const [pvModules, setPvModules] = useState<PvModule[]>([])
  const [projectInfo, setProjectInfo] = useState<ProjectInfo>(readProjectInfo)
  const [projectInfoLibrary, setProjectInfoLibrary] = useState<ProjectInfoLibraryItem[]>(readProjectInfoLibrary)
  const [projectInfoStatus, setProjectInfoStatus] = useState<string | null>(null)
  const [recordNoteOverrides, setRecordNoteOverrides] = useState<RecordNoteOverrides>(readRecordNoteOverrides)
  const [recordImageOverrides, setRecordImageOverrides] = useState<RecordImageOverrides>(readRecordImageOverrides)
  const [selectedSystemGroup, setSelectedSystemGroup] = useState(savedSession?.selectedSystemGroup ?? '')
  const [selectedInverter, setSelectedInverter] = useState(savedSession?.selectedInverter ?? '')
  const [selectedRecordPath, setSelectedRecordPath] = useState(savedSession?.selectedRecordPath ?? '')
  const [selectedPvModuleId, setSelectedPvModuleId] = useState(savedSession?.selectedPvModuleId ?? '')
  const [toleranceRows, setToleranceRows] = useState<ToleranceRow[]>(readToleranceRows)
  const [isChatOpen, setIsChatOpen] = useState(savedSession?.isChatOpen ?? true)
  const [isChatBusy, setIsChatBusy] = useState(false)
  const [isChatSettingsOpen, setIsChatSettingsOpen] = useState(false)
  const [isCodexActionRunning, setIsCodexActionRunning] = useState(false)
  const [codexStatus, setCodexStatus] = useState<CodexSetupStatus | null>(null)
  const [codexActionOutput, setCodexActionOutput] = useState<string | null>(null)
  const [aiWorkspaceInput, setAiWorkspaceInput] = useState('')
  const [chatInput, setChatInput] = useState('')
  const [chatImages, setChatImages] = useState<ChatImageAttachment[]>([])
  const [chatMessages, setChatMessages] = useState<AppChatMessage[]>([
    {
      id: createChatMessageId(),
      role: 'assistant',
      content: 'Tôi có thể điều khiển app bằng lệnh chat. Ví dụ: mở PV Module, chọn X3 inverter 4 string 4.9.2, next string, export DOCX.',
    },
  ])

  const selector = useMemo(
    () =>
      buildModuleSelector(
        importResult?.records ?? [],
        selectedSystemGroup,
        selectedInverter,
        selectedRecordPath,
      ),
    [importResult, selectedInverter, selectedRecordPath, selectedSystemGroup],
  )
  const selectedRecord = selector.selectedRecord
  const matchedPvModule = useMemo(
    () => findPvModuleForRecord(selectedRecord, pvModules),
    [pvModules, selectedRecord],
  )
  const selectedPvModule = useMemo(
    () => pvModules.find((pvModule) => pvModule.id === selectedPvModuleId) ?? matchedPvModule,
    [matchedPvModule, pvModules, selectedPvModuleId],
  )
  const reportSummaryRows = useMemo(
    () =>
      buildIvCurveReportSummaryRows(
        importResult?.records ?? [],
        pvModules,
        selectedPvModuleId,
        toleranceRows,
      ),
    [importResult, pvModules, selectedPvModuleId, toleranceRows],
  )
  const reportDetailRows = useMemo(
    () =>
      buildIvCurveReportDetailRows(
        importResult?.records ?? [],
        pvModules,
        selectedPvModuleId,
        toleranceRows,
      ),
    [importResult, pvModules, selectedPvModuleId, toleranceRows],
  )
  const orderedRecords = useMemo(
    () => sortRecordsByHierarchy(importResult?.records ?? []),
    [importResult],
  )
  const selectedRecordIndex = selectedRecord
    ? orderedRecords.findIndex((record) => record.sourcePath === selectedRecord.sourcePath)
    : -1
  const canMovePrevious = selectedRecordIndex > 0
  const canMoveNext = selectedRecordIndex >= 0 && selectedRecordIndex < orderedRecords.length - 1

  useEffect(() => {
    let isActive = true

    const loadPvModules = async () => {
      setIsModuleLibraryLoading(true)

      try {
        const modules = (await window.ipcRenderer.invoke('pv-modules:list')) as PvModule[]

        if (isActive) {
          setPvModules(modules)
        }
      } finally {
        if (isActive) {
          setIsModuleLibraryLoading(false)
        }
      }
    }

    loadPvModules()

    return () => {
      isActive = false
    }
  }, [])

  useEffect(() => {
    if (sessionRestoreAttemptedRef.current || !savedSession?.folderPath) {
      return
    }

    sessionRestoreAttemptedRef.current = true
    let isActive = true

    const restoreSession = async () => {
      setExportStatus('Restoring previous data folder...')

      try {
        const selectedData = await invokeWithTimeout<DataImportResult | null>(
          'data:import-folder',
          restoreImportTimeoutMs,
          savedSession.folderPath,
        )

        if (!isActive || !selectedData) {
          return
        }

        const nextImportResult = applyRecordNoteOverrides(selectedData as DataImportResult, recordNoteOverrides)
        const nextSelection = getRestoredModuleSelection(nextImportResult.records, savedSession)

        setImportResult(nextImportResult)
        setSelectedSystemGroup(nextSelection.systemGroup)
        setSelectedInverter(nextSelection.inverter)
        setSelectedRecordPath(nextSelection.recordPath)
        setSelectedPvModuleId(savedSession.selectedPvModuleId)
        setActiveView(savedSession.activeView)
        setExportStatus(null)
      } catch (error) {
        if (isActive) {
          setExportStatus(`Unable to restore previous data folder: ${getExportErrorMessage(error)}`)
        }
      }
    }

    restoreSession()

    return () => {
      isActive = false
    }
  }, [recordNoteOverrides, savedSession])

  useEffect(() => {
    if (!importResult) {
      setSelectedPvModuleId('')
      return
    }

    if (pvModules.length === 0) {
      return
    }

    const cachedModule = getCachedProjectModule(importResult.folderPath, pvModules)
    setSelectedPvModuleId((currentModuleId) => {
      if (currentModuleId && pvModules.some((pvModule) => pvModule.id === currentModuleId)) {
        return currentModuleId
      }

      return cachedModule?.id ?? ''
    })
  }, [importResult, pvModules])

  useEffect(() => {
    writeToleranceRows(toleranceRows)
  }, [toleranceRows])

  useEffect(() => {
    writeProjectInfo(projectInfo)
  }, [projectInfo])

  useEffect(() => {
    writeProjectInfoLibrary(projectInfoLibrary)
  }, [projectInfoLibrary])

  useEffect(() => {
    writeRecordNoteOverrides(recordNoteOverrides)
  }, [recordNoteOverrides])

  useEffect(() => {
    writeRecordImageOverrides(recordImageOverrides)
  }, [recordImageOverrides])

  useEffect(() => {
    if (!importResult) {
      return
    }

    writeAppSession({
      version: 1,
      activeView,
      folderPath: importResult.folderPath,
      isChatOpen,
      selectedInverter,
      selectedPvModuleId,
      selectedRecordPath,
      selectedRecordRelativePath: selectedRecord?.relativePath ?? '',
      selectedSystemGroup,
    })
  }, [
    activeView,
    importResult,
    isChatOpen,
    selectedInverter,
    selectedPvModuleId,
    selectedRecord,
    selectedRecordPath,
    selectedSystemGroup,
  ])

  const importDataFromChannel = async (channel: 'data:select-folder' | 'data:import-ai-workspace') => {
    setIsSelecting(true)
    setExportedFilePath(null)
    setExportStatus(null)

    try {
      const selectedData = await window.ipcRenderer.invoke(channel)

      if (selectedData) {
        const nextImportResult = applyRecordNoteOverrides(selectedData as DataImportResult, recordNoteOverrides)
        const nextSelection = getInitialModuleSelection(nextImportResult.records)

        setImportResult(nextImportResult)
        setSelectedSystemGroup(nextSelection.systemGroup)
        setSelectedInverter(nextSelection.inverter)
        setSelectedRecordPath(nextSelection.recordPath)
        setSelectedPvModuleId('')
        setActiveView('home')
        return true
      }

      return false
    } catch (error) {
      setExportStatus(`Unable to import data folder: ${getExportErrorMessage(error)}`)
      return false
    } finally {
      setIsSelecting(false)
    }
  }

  const handleAddData = async () => {
    return importDataFromChannel('data:select-folder')
  }

  const handleAddAiWorkspaceData = async () => {
    return importDataFromChannel('data:import-ai-workspace')
  }

  const handlePvModuleChange = (moduleId: string) => {
    setSelectedPvModuleId(moduleId)

    if (!importResult?.folderPath) {
      return
    }

    const pvModule = pvModules.find((moduleItem) => moduleItem.id === moduleId) ?? null

    if (pvModule) {
      cacheProjectModule(importResult.folderPath, pvModule)
    } else {
      clearCachedProjectModule(importResult.folderPath)
    }
  }

  const resolvePvModuleForRecord = (record: MeasurementRecord | null, moduleId = selectedPvModuleId) => {
    return pvModules.find((pvModule) => pvModule.id === moduleId) ?? findPvModuleForRecord(record, pvModules)
  }

  const exportDocxForRecord = async (
    recordToExport: MeasurementRecord | null,
    pvModuleToExport: PvModule | null,
    moduleIdToExport: string,
  ) => {
    if (!recordToExport) {
      return false
    }

    setIsExporting(true)
    setExportedFilePath(null)
    setExportStatus(null)
    setExportProgress({
      elapsedMs: 0,
      message: 'Opening save dialog...',
      percent: 3,
    })

    const handleExportProgress = (_event: unknown, progress: WordExportProgress) => {
      setExportProgress(progress)
    }

    window.ipcRenderer.on('data:export-docx-progress', handleExportProgress)

    try {
      const payload: WordExportPayload = {
        folderPath: importResult?.folderPath ?? null,
        projectInfo,
        pvModule: pvModuleToExport,
        record: recordToExport,
        recordImages: recordImageOverrides,
        records: importResult?.records ?? [],
        pvModules,
        selectedPvModuleId: moduleIdToExport,
        toleranceRows,
      }
      const result = (await window.ipcRenderer.invoke('data:export-docx', payload)) as WordExportResult

      if (!result.canceled && result.filePath) {
        setExportedFilePath(result.filePath)
        setExportStatus(`${result.warning ? `${result.warning} ` : ''}Saved: ${result.filePath}`)
        setExportProgress({
          completed: true,
          elapsedMs: result.elapsedMs,
          message: `Completed in ${formatDuration(result.elapsedMs ?? 0)}`,
          percent: 100,
        })
        return true
      } else if (result.canceled) {
        setExportProgress({
          completed: true,
          message: 'Export canceled.',
          percent: 100,
        })
      }
    } catch (error) {
      const message = getExportErrorMessage(error)

      setExportStatus(`Export failed: ${message}`)
      setExportProgress({
        completed: true,
        message: 'Export failed.',
        percent: 100,
      })
    } finally {
      window.ipcRenderer.off('data:export-docx-progress', handleExportProgress)
      setIsExporting(false)
    }

    return false
  }

  const handleExportDocx = async () => {
    return exportDocxForRecord(selectedRecord, selectedPvModule, selectedPvModuleId)
  }

  const handleOpenExportedFile = async () => {
    if (!exportedFilePath) {
      return false
    }

    const errorMessage = (await window.ipcRenderer.invoke('file:open', exportedFilePath)) as string | null

    if (errorMessage) {
      setExportStatus(`Unable to open file: ${errorMessage}`)
      return false
    }

    return true
  }

  const savePvModuleLibrary = async (nextModules: PvModule[]) => {
    setModuleLibraryStatus(null)
    const savedModules = (await window.ipcRenderer.invoke('pv-modules:save', nextModules)) as PvModule[]

    setPvModules(savedModules)
    setModuleLibraryStatus('PV module library saved')
  }

  const selectRecord = (record: MeasurementRecord) => {
    setSelectedSystemGroup(getSystemGroupKey(record))
    setSelectedInverter(getInverterKey(record))
    setSelectedRecordPath(record.sourcePath)
  }

  const moveRecord = (direction: -1 | 1) => {
    const nextRecord = orderedRecords[selectedRecordIndex + direction]

    if (nextRecord) {
      selectRecord(nextRecord)
    }
  }

  const updateToleranceValue = (rowLabel: string, columnKey: string, value: number) => {
    if (columnKey === TOLERANCE_TOTAL_KEY) {
      return
    }

    setToleranceRows((currentRows) =>
      currentRows.map((row) =>
        row.label === rowLabel
          ? {
              ...row,
              values: {
                ...row.values,
                [columnKey]: value,
              },
            }
          : row,
      ),
    )
  }

  const updateProjectInfoValue = (key: keyof ProjectInfo, value: string) => {
    setProjectInfoStatus(null)
    setProjectInfo((currentInfo) => ({ ...currentInfo, [key]: normalizeDisplayText(value) }))
  }

  const applyProjectInfoFromLibrary = (itemId: string) => {
    const item = projectInfoLibrary.find((libraryItem) => libraryItem.id === itemId)

    if (!item) {
      return
    }

    setProjectInfo(normalizeProjectInfoForDisplay(createProjectInfoFromLibraryItem(item)))
    setProjectInfoStatus(`Loaded project info: ${getProjectInfoLibraryLabel(item)}`)
  }

  const saveProjectInfoToLibrary = () => {
    const now = new Date().toISOString()
    const normalizedProjectInfo = normalizeProjectInfoForDisplay(projectInfo)
    const label = getProjectInfoLibraryLabel(normalizedProjectInfo)
    const nextItem: ProjectInfoLibraryItem = {
      ...normalizedProjectInfo,
      id: createProjectInfoLibraryId(),
      label,
      updatedAt: now,
    }

    setProjectInfoLibrary((currentItems) => sortProjectInfoLibrary([nextItem, ...currentItems]))
    setProjectInfoStatus(`Saved project info: ${label}`)
  }

  const deleteProjectInfoFromLibrary = (itemId: string) => {
    const item = projectInfoLibrary.find((libraryItem) => libraryItem.id === itemId)

    const label = item ? getProjectInfoLibraryLabel(item) : ''

    if (!item || !window.confirm(`Delete project info "${label}"?`)) {
      return
    }

    setProjectInfoLibrary((currentItems) => currentItems.filter((libraryItem) => libraryItem.id !== itemId))
    setProjectInfoStatus(`Deleted project info: ${label}`)
  }

  const updateRecordNote = (sourcePath: string, value: string) => {
    const normalizedValue = value.trim() ? value : ''

    setImportResult((currentImportResult) => {
      if (!currentImportResult) {
        return currentImportResult
      }

      const targetRecord = currentImportResult.records.find((record) => record.sourcePath === sourcePath)

      if (!targetRecord) {
        return currentImportResult
      }

      const noteKey = getRecordNoteOverrideKey(currentImportResult.folderPath, targetRecord)

      setRecordNoteOverrides((currentOverrides) => ({
        ...currentOverrides,
        [noteKey]: normalizedValue,
      }))

      return {
        ...currentImportResult,
        records: currentImportResult.records.map((record) =>
          record.sourcePath === sourcePath
            ? {
                ...record,
                measurementNote: normalizedValue || null,
              }
            : record,
        ),
      }
    })
  }

  const updateRecordImageSlot = (sourcePath: string, slotIndex: number, slot: RecordImageSlot | null) => {
    setImportResult((currentImportResult) => {
      if (!currentImportResult) {
        return currentImportResult
      }

      const targetRecord = currentImportResult.records.find((record) => record.sourcePath === sourcePath)

      if (!targetRecord) {
        return currentImportResult
      }

      const imageKey = getRecordImageOverrideKey(currentImportResult.folderPath, targetRecord)

      setRecordImageOverrides((currentOverrides) => {
        const nextSlots = normalizeRecordImageSlots(currentOverrides[imageKey])

        nextSlots[slotIndex] = slot

        const nextOverrides = { ...currentOverrides }

        if (nextSlots.some(Boolean)) {
          nextOverrides[imageKey] = nextSlots
        } else {
          delete nextOverrides[imageKey]
        }

        return nextOverrides
      })

      return currentImportResult
    })
  }

  const executeChatCommands = async (commands: ChatCommand[]) => {
    const results: string[] = []
    let workingRecord = selectedRecord
    let workingModuleId = selectedPvModuleId
    let workingPvModule = workingRecord ? resolvePvModuleForRecord(workingRecord, workingModuleId) : selectedPvModule
    let workingPvModules = pvModules

    for (const command of commands) {
      if (command.action === 'set_view') {
        if (command.view) {
          setActiveView(command.view)
          results.push(`- Đã mở ${getAppViewLabel(command.view)}.`)
        }

        continue
      }

      if (command.action === 'select_data_folder') {
        const imported = await handleAddAiWorkspaceData()
        results.push(imported ? '- Đã import folder dữ liệu.' : '- Bạn đã hủy chọn folder dữ liệu.')
        continue
      }

      if (command.action === 'select_record') {
        const record = findRecordForChatCommand(importResult?.records ?? [], command)

        if (record) {
          selectRecord(record)
          workingRecord = record
          workingPvModule = resolvePvModuleForRecord(record, workingModuleId)
          setActiveView('home')
          results.push(`- Đã chọn ${getReportInverterLabel(record)} / ${getStringKey(record)}.`)
        } else {
          results.push('- Không tìm thấy string phù hợp trong dữ liệu đã import.')
        }

        continue
      }

      if (command.action === 'move_record') {
        const nextRecord = getChatMovedRecord(orderedRecords, workingRecord, command.direction)

        if (nextRecord) {
          selectRecord(nextRecord)
          workingRecord = nextRecord
          workingPvModule = resolvePvModuleForRecord(nextRecord, workingModuleId)
          setActiveView('home')
          results.push(`- Đã chọn ${getReportInverterLabel(nextRecord)} / ${getStringKey(nextRecord)}.`)
        } else {
          results.push('- Không còn string để chuyển theo hướng đó.')
        }

        continue
      }

      if (command.action === 'select_module') {
        const pvModule = findPvModuleForChatCommand(workingPvModules, command)

        if (pvModule) {
          handlePvModuleChange(pvModule.id)
          workingModuleId = pvModule.id
          workingPvModule = pvModule
          setActiveView('pv-module')
          results.push(`- Đã chọn module ${pvModule.model}.`)
        } else {
          results.push('- Không tìm thấy PV module phù hợp trong thư viện.')
        }

        continue
      }

      if (command.action === 'upsert_pv_module') {
        const pvModulePayload = command.pvModule

        if (!pvModulePayload?.model.trim()) {
          results.push('- Lệnh tạo PV module thiếu model.')
          continue
        }

        const existingModule =
          workingPvModules.find((pvModule) => chatTextMatches(pvModule.model, pvModulePayload.model)) ?? null
        const nextModule = createPvModuleFromChatPayload(pvModulePayload, existingModule)
        const nextModules = sortPvModules(
          existingModule
            ? workingPvModules.map((pvModule) => (pvModule.id === existingModule.id ? nextModule : pvModule))
            : [...workingPvModules, nextModule],
        )

        await savePvModuleLibrary(nextModules)
        workingPvModules = nextModules
        workingModuleId = nextModule.id
        workingPvModule = nextModule
        setSelectedPvModuleId(nextModule.id)
        setActiveView('pv-module')

        if (importResult?.folderPath) {
          cacheProjectModule(importResult.folderPath, nextModule)
        }

        results.push(`${existingModule ? '- Đã cập nhật' : '- Đã thêm'} PV module ${nextModule.model}.`)
        continue
      }

      if (command.action === 'set_tolerance') {
        if (
          command.rowLabel &&
          command.columnKey &&
          command.numericValue !== null &&
          TOLERANCE_COLUMNS.some((column) => column.key === command.columnKey)
        ) {
          updateToleranceValue(command.rowLabel, command.columnKey, command.numericValue)
          results.push(`- Đã set ${command.rowLabel} ${command.columnKey} = ${command.numericValue}%.`)
        } else {
          results.push('- Lệnh tolerance thiếu row, cột hoặc giá trị số.')
        }

        continue
      }

      if (command.action === 'export_docx') {
        const exported = await exportDocxForRecord(workingRecord, workingPvModule, workingModuleId)
        results.push(exported ? '- Đã export DOCX.' : '- Chưa export được DOCX.')
        continue
      }

      if (command.action === 'open_exported_file') {
        const opened = await handleOpenExportedFile()
        results.push(opened ? '- Đã mở file export gần nhất.' : '- Chưa có file export để mở.')
      }
    }

    return results
  }

  const handleChatSubmit = async (message: string) => {
    const trimmedMessage = message.trim()

    if ((!trimmedMessage && chatImages.length === 0) || isChatBusy) {
      return
    }

    const userMessage: AppChatMessage = {
      id: createChatMessageId(),
      role: 'user',
      content: trimmedMessage || 'Analyze attached image.',
      images: chatImages,
    }
    const nextMessages = [...chatMessages, userMessage]
    const history = nextMessages.slice(-10).map(({ role, content, images }) => ({ role, content, images }))
    const request: ChatCompletionRequest = {
      message: userMessage.content,
      history,
      context: buildChatAppContext({
        activeView,
        canMoveNext,
        canMovePrevious,
        importResult,
        orderedRecords,
        pvModules,
        reportDetailRows,
        reportSummaryRows,
        selectedPvModule,
        selectedRecord,
        selectedRecordIndex,
        selector,
        toleranceRows,
      }),
      images: chatImages,
    }

    setChatInput('')
    setChatImages([])
    setChatMessages((currentMessages) => [...currentMessages, userMessage])
    setIsChatBusy(true)

    try {
      const response = (await window.ipcRenderer.invoke('chat:complete', request)) as ChatCompletionResult
      const commandResults = await executeChatCommands(response.commands)
      const sourceText =
        response.source === 'codex'
          ? `\n\nCodex SDK${response.model ? `: ${response.model}` : ''}`
          : response.source === 'openai' && response.model
          ? `\n\nModel: ${response.model}`
          : response.source === 'local'
            ? '\n\nLocal parser'
            : ''
      const assistantText = [response.reply, ...commandResults].filter(Boolean).join('\n')

      setChatMessages((currentMessages) => [
        ...currentMessages,
        {
          id: createChatMessageId(),
          role: 'assistant',
          content: `${assistantText}${sourceText}`,
        },
      ])
    } catch (error) {
      setChatMessages((currentMessages) => [
        ...currentMessages,
        {
          id: createChatMessageId(),
          role: 'assistant',
          content: `Chatbot lỗi: ${getExportErrorMessage(error)}`,
        },
      ])
    } finally {
      setIsChatBusy(false)
    }
  }

  const handleChatImagesAdd = async (files: FileList | null) => {
    if (!files || files.length === 0) {
      return
    }

    const imageSlots = Math.max(0, chatImageMaxCount - chatImages.length)
    const imageFiles = Array.from(files)
      .filter((file) => file.type.startsWith('image/') && file.size <= chatImageMaxBytes)
      .slice(0, imageSlots)
    const nextImages = await Promise.all(imageFiles.map(readChatImageAttachment))

    setChatImages((currentImages) => [...currentImages, ...nextImages])
  }

  const handleChatImageRemove = (imageId: string) => {
    setChatImages((currentImages) => currentImages.filter((image) => image.id !== imageId))
  }

  const refreshCodexStatus = async () => {
    const status = (await window.ipcRenderer.invoke('codex:status')) as CodexSetupStatus

    setCodexStatus(status)
    setAiWorkspaceInput(status.workspacePath)
    return status
  }

  const runCodexSetupAction = async (channel: CodexSetupActionChannel, ...args: unknown[]) => {
    setIsCodexActionRunning(true)
    setCodexActionOutput(null)

    try {
      const result = (await window.ipcRenderer.invoke(channel, ...args)) as CodexSetupActionResult

      setCodexActionOutput([result.message, result.output].filter(Boolean).join('\n\n'))
      await refreshCodexStatus()
    } catch (error) {
      setCodexActionOutput(`Codex setup failed: ${getExportErrorMessage(error)}`)
    } finally {
      setIsCodexActionRunning(false)
    }
  }

  return (
    <>
      <main className="min-h-screen bg-blue-50 text-slate-950">
        <div className="mx-auto flex w-full max-w-[1600px] flex-col gap-5 p-4 sm:p-6">
        <div className="flex min-h-14 flex-wrap items-center justify-between gap-3 rounded-md border border-blue-100 bg-white px-4 py-3 shadow-sm">
          <div
            className="flex h-11 rounded-md border border-blue-200 bg-blue-50 p-1"
            role="tablist"
            aria-label="Primary views"
          >
            <button
              className={viewTabClass(activeView === 'home')}
              type="button"
              role="tab"
              aria-selected={activeView === 'home'}
              onClick={() => setActiveView('home')}
            >
              Home
            </button>

            <button
              className={viewTabClass(activeView === 'project-info')}
              type="button"
              role="tab"
              aria-selected={activeView === 'project-info'}
              onClick={() => setActiveView('project-info')}
            >
              Information
            </button>

            <button
              className={viewTabClass(activeView === 'pv-module')}
              type="button"
              role="tab"
              aria-selected={activeView === 'pv-module'}
              onClick={() => setActiveView('pv-module')}
            >
              PV Module
            </button>
          </div>

          <div className="flex flex-wrap items-center justify-end gap-2">
            <button
              className="h-11 rounded-md bg-blue-950 px-5 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-900 focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-blue-400 disabled:cursor-default disabled:opacity-60"
              onClick={handleAddData}
              disabled={isSelecting}
            >
              {isSelecting ? 'Selecting...' : '+ Add Data'}
            </button>

            <button
              className="h-11 rounded-md border border-blue-200 bg-white px-5 text-sm font-semibold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50 focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-blue-400 disabled:cursor-default disabled:text-slate-300 disabled:hover:border-slate-300 disabled:hover:bg-white"
              onClick={handleExportDocx}
              disabled={!selectedRecord || isExporting}
            >
              {isExporting ? 'Exporting...' : 'Export DOCX'}
            </button>
          </div>

          {importResult && (
            <div className="flex items-baseline gap-1 text-sm text-blue-700">
              <strong className="text-lg font-semibold text-blue-950">
                {importResult.records.length}
              </strong>
              <span>/{importResult.totalFiles} CSV loaded</span>
            </div>
          )}
        </div>

        {activeView === 'pv-module' && (
          <PvModuleLibrary
            isLoading={isModuleLibraryLoading}
            modules={pvModules}
            onModuleChange={handlePvModuleChange}
            onSave={savePvModuleLibrary}
            selectedModuleId={selectedPvModuleId}
            selectedModuleModel={selectedPvModule?.model ?? selectedRecord?.moduleModel ?? null}
            status={moduleLibraryStatus}
          />
        )}

        {activeView === 'project-info' && (
          <ProjectInformationPanel
            info={projectInfo}
            library={projectInfoLibrary}
            selectedModuleModel={selectedPvModule?.model ?? selectedRecord?.moduleModel ?? null}
            status={projectInfoStatus}
            onApplyLibraryItem={applyProjectInfoFromLibrary}
            onDeleteLibraryItem={deleteProjectInfoFromLibrary}
            onLoadReference={() => {
              setProjectInfo(normalizeProjectInfoForDisplay(referenceProjectInfo))
              setProjectInfoStatus('Loaded I-V report sample')
            }}
            onSaveToLibrary={saveProjectInfoToLibrary}
            onUpdate={updateProjectInfoValue}
          />
        )}

        {activeView === 'home' && (
          <section className="flex flex-col gap-4" aria-label="Imported data overview">
            {importResult ? (
              <>
                <p className="break-anywhere m-0 text-sm text-slate-500">{importResult.folderPath}</p>
                {exportStatus && (
                  <div className="flex flex-wrap items-center gap-2">
                    <p className="break-anywhere m-0 text-sm font-medium text-emerald-700">{exportStatus}</p>
                    {exportedFilePath && (
                      <button
                        className="h-8 rounded-md border border-emerald-300 bg-white px-3 text-sm font-semibold text-emerald-800 shadow-sm transition hover:border-emerald-400 hover:bg-emerald-50"
                        type="button"
                        onClick={handleOpenExportedFile}
                      >
                        Open
                      </button>
                    )}
                  </div>
                )}
                {exportProgress && <ExportProgressBar progress={exportProgress} />}

                <ProjectInformationSummary
                  info={projectInfo}
                  selectedModuleModel={selectedPvModule?.model ?? selectedRecord?.moduleModel ?? null}
                  onEdit={() => setActiveView('project-info')}
                />

                {importResult.records.length > 0 && (
                  <ToleranceCalculationDetails rows={toleranceRows} onValueChange={updateToleranceValue} />
                )}

                {importResult.records.length > 0 && (
                  <IvCurveReportSummaryTable rows={reportSummaryRows} selectedInverter={selectedRecord ? getReportInverterLabel(selectedRecord) : null} />
                )}

                {importResult.records.length > 0 && (
                  <IvCurveReportDetailTable
                    onNoteChange={updateRecordNote}
                    rows={reportDetailRows}
                    selectedRowKey={
                      selectedRecord ? `${getReportInverterLabel(selectedRecord)}-${getStringKey(selectedRecord)}` : null
                    }
                  />
                )}

                {importResult.records.length > 0 && (
                  <ModuleSelector
                    selector={selector}
                    onSystemChange={(systemGroup) => {
                      const nextSelection = getInitialModuleSelection(importResult.records, systemGroup)

                      setSelectedSystemGroup(nextSelection.systemGroup)
                      setSelectedInverter(nextSelection.inverter)
                      setSelectedRecordPath(nextSelection.recordPath)
                    }}
                    onInverterChange={(inverter) => {
                      const nextSelection = getInitialModuleSelection(
                        importResult.records,
                        selector.selectedSystemGroup,
                        inverter,
                      )

                      setSelectedSystemGroup(nextSelection.systemGroup)
                      setSelectedInverter(nextSelection.inverter)
                      setSelectedRecordPath(nextSelection.recordPath)
                    }}
                    onRecordChange={setSelectedRecordPath}
                    onPrevious={() => moveRecord(-1)}
                    onNext={() => moveRecord(1)}
                    canMovePrevious={canMovePrevious}
                    canMoveNext={canMoveNext}
                    selectedIndex={selectedRecordIndex}
                    totalRecords={orderedRecords.length}
                  />
                )}

                {selectedRecord ? (
                  <MeasurementReport
                    imageSlots={
                      importResult
                        ? getRecordImageSlots(recordImageOverrides, importResult.folderPath, selectedRecord)
                        : createEmptyRecordImageSlots()
                    }
                    onImageSlotChange={(slotIndex, slot) => updateRecordImageSlot(selectedRecord.sourcePath, slotIndex, slot)}
                    record={selectedRecord}
                    pvModule={selectedPvModule}
                    toleranceRows={toleranceRows}
                  />
                ) : (
                  <p className="m-0 text-sm text-slate-500">No valid CSV files found.</p>
                )}

                {importResult.errors.length > 0 && (
                  <p className="m-0 text-sm font-medium text-red-700">
                    {importResult.errors.length} file(s) could not be parsed.
                  </p>
                )}
              </>
            ) : (
              <>
                <section className="rounded-md border border-dashed border-blue-200 bg-white p-4 text-sm text-slate-600 shadow-sm">
                  Add a data folder to preview IV records and export the DOCX report.
                </section>
                <ProjectInformationSummary
                  info={projectInfo}
                  selectedModuleModel={selectedPvModule?.model ?? selectedRecord?.moduleModel ?? null}
                  onEdit={() => setActiveView('project-info')}
                />
              </>
            )}
          </section>
        )}
        </div>
      </main>
      <ChatControlPanel
        aiWorkspaceInput={aiWorkspaceInput}
        input={chatInput}
        images={chatImages}
        isCodexActionRunning={isCodexActionRunning}
        isBusy={isChatBusy}
        isOpen={isChatOpen}
        isSettingsOpen={isChatSettingsOpen}
        messages={chatMessages}
        codexActionOutput={codexActionOutput}
        codexStatus={codexStatus}
        onImagesAdd={handleChatImagesAdd}
        onImageRemove={handleChatImageRemove}
        onInputChange={setChatInput}
        onCodexInstall={() => runCodexSetupAction('codex:install')}
        onCodexLogin={() => runCodexSetupAction('codex:login')}
        onCodexOpenWorkspace={() => runCodexSetupAction('codex:open-workspace')}
        onCodexRefresh={() => {
          void refreshCodexStatus()
        }}
        onCodexResetWorkspace={() => runCodexSetupAction('codex:reset-workspace')}
        onCodexResetThread={() => runCodexSetupAction('codex:reset-thread')}
        onCodexSaveWorkspace={() => runCodexSetupAction('codex:set-workspace', aiWorkspaceInput)}
        onCodexSelectWorkspace={() => runCodexSetupAction('codex:select-workspace')}
        onAiWorkspaceInputChange={setAiWorkspaceInput}
        onSettingsToggle={() => {
          setIsChatSettingsOpen((currentValue) => {
            const nextValue = !currentValue

            if (nextValue) {
              void refreshCodexStatus()
            }

            return nextValue
          })
        }}
        onSubmit={handleChatSubmit}
        onToggle={() => setIsChatOpen((currentValue) => !currentValue)}
      />
    </>
  )
}

interface PvModuleLibraryProps {
  isLoading: boolean
  modules: PvModule[]
  onModuleChange: (moduleId: string) => void
  onSave: (modules: PvModule[]) => Promise<void>
  selectedModuleId: string
  selectedModuleModel: string | null
  status: string | null
}

interface ChatControlPanelProps {
  aiWorkspaceInput: string
  codexActionOutput: string | null
  codexStatus: CodexSetupStatus | null
  input: string
  images: ChatImageAttachment[]
  isCodexActionRunning: boolean
  isBusy: boolean
  isOpen: boolean
  isSettingsOpen: boolean
  messages: AppChatMessage[]
  onAiWorkspaceInputChange: (value: string) => void
  onCodexInstall: () => void | Promise<void>
  onCodexLogin: () => void | Promise<void>
  onCodexOpenWorkspace: () => void | Promise<void>
  onCodexRefresh: () => void | Promise<void>
  onCodexResetWorkspace: () => void | Promise<void>
  onCodexResetThread: () => void | Promise<void>
  onCodexSaveWorkspace: () => void | Promise<void>
  onCodexSelectWorkspace: () => void | Promise<void>
  onImagesAdd: (files: FileList | null) => void | Promise<void>
  onImageRemove: (imageId: string) => void
  onInputChange: (value: string) => void
  onSettingsToggle: () => void
  onSubmit: (message: string) => void | Promise<void>
  onToggle: () => void
}

function ChatControlPanel({
  aiWorkspaceInput,
  codexActionOutput,
  codexStatus,
  input,
  images,
  isCodexActionRunning,
  isBusy,
  isOpen,
  isSettingsOpen,
  messages,
  onAiWorkspaceInputChange,
  onCodexInstall,
  onCodexLogin,
  onCodexOpenWorkspace,
  onCodexRefresh,
  onCodexResetWorkspace,
  onCodexResetThread,
  onCodexSaveWorkspace,
  onCodexSelectWorkspace,
  onImagesAdd,
  onImageRemove,
  onInputChange,
  onSettingsToggle,
  onSubmit,
  onToggle,
}: ChatControlPanelProps) {
  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    onSubmit(input)
  }

  return (
    <>
      <aside className="fixed bottom-4 right-4 z-50 flex w-[min(460px,calc(100vw-2rem))] flex-col items-end gap-2">
      {isOpen && (
        <section className="flex h-[min(680px,calc(100vh-7rem))] w-full overflow-hidden rounded-lg border border-blue-200 bg-white shadow-2xl ring-1 ring-blue-950/10">
          <div className="flex h-full min-w-0 flex-1 flex-col">
          <div className="border-b border-blue-900 bg-blue-950 px-3 py-3 text-white">
            <div className="flex items-center justify-between gap-3">
              <div className="min-w-0">
                <div className="flex items-center gap-2">
                  <span className="flex h-8 w-8 items-center justify-center rounded-md bg-white text-sm font-black text-blue-950">
                    AI
                  </span>
                  <div className="min-w-0">
                    <h2 className="m-0 text-sm font-bold leading-tight">AI Control</h2>
                    <p className="m-0 text-xs text-slate-300">
                      {codexStatus?.installed ? 'Codex ready' : 'Codex setup'}
                    </p>
                  </div>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <span
                  className={`h-2.5 w-2.5 rounded-full ${
                    codexStatus?.installed ? 'bg-emerald-400' : 'bg-amber-400'
                  }`}
                  title={codexStatus?.installed ? 'Codex ready' : 'Codex setup needed'}
                />
                <button className={chatHeaderButtonClass} type="button" onClick={onSettingsToggle}>
                  Settings
                </button>
                <button className={chatHeaderButtonClass} type="button" onClick={onToggle}>
                  Hide
                </button>
              </div>
            </div>
          </div>

          <div className="flex-1 space-y-4 overflow-y-auto bg-blue-50 p-3">
            {messages.map((message) => (
              <div className={`flex gap-2 ${message.role === 'user' ? 'justify-end' : 'justify-start'}`} key={message.id}>
                {message.role === 'assistant' && <ChatAvatar label="AI" tone="assistant" />}
                <div
                  className={`max-w-[82%] rounded-lg px-3 py-2 text-sm leading-relaxed shadow-sm ${
                    message.role === 'user'
                      ? 'bg-blue-950 text-white'
                      : 'border border-slate-200 bg-white text-slate-900'
                  }`}
                >
                  <p className="m-0 whitespace-pre-wrap">{message.content}</p>
                  {message.images && message.images.length > 0 && (
                    <div className="mt-2 grid grid-cols-2 gap-2">
                      {message.images.map((image) => (
                        <img
                          alt={image.name}
                          className="h-24 w-full rounded-md border border-slate-200 object-cover"
                          key={image.id}
                          src={image.dataUrl}
                        />
                      ))}
                    </div>
                  )}
                </div>
                {message.role === 'user' && <ChatAvatar label="You" tone="user" />}
              </div>
            ))}
            {isBusy && (
              <div className="flex items-center gap-2">
                <ChatAvatar label="AI" tone="assistant" />
                <div className="rounded-lg border border-slate-200 bg-white px-3 py-2 text-sm text-slate-600 shadow-sm">
                  Thinking...
                </div>
              </div>
            )}
          </div>

          <form className="border-t border-slate-200 bg-white p-3" onSubmit={handleSubmit}>
            {images.length > 0 && (
              <div className="mb-2 grid grid-cols-4 gap-2">
                {images.map((image) => (
                  <div className="relative" key={image.id}>
                    <img
                      alt={image.name}
                      className="h-20 w-full rounded-md border border-slate-200 object-cover"
                      src={image.dataUrl}
                    />
                    <button
                      aria-label={`Remove ${image.name}`}
                      className="absolute right-1 top-1 h-6 w-6 rounded-md bg-blue-950 text-xs font-bold text-white shadow"
                      disabled={isBusy}
                      onClick={() => onImageRemove(image.id)}
                      type="button"
                    >
                      x
                    </button>
                  </div>
                ))}
              </div>
            )}

            <div className="flex gap-2">
              <label
                className={`flex h-11 w-11 cursor-pointer items-center justify-center rounded-md border border-blue-200 bg-white text-lg font-bold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50 ${
                  isBusy || images.length >= chatImageMaxCount ? 'pointer-events-none opacity-50' : ''
                }`}
                title="Attach image"
              >
                +
                <input
                  accept="image/png,image/jpeg,image/webp,image/gif"
                  className="sr-only"
                  disabled={isBusy || images.length >= chatImageMaxCount}
                  multiple
                  onChange={(event) => {
                    onImagesAdd(event.target.files)
                    event.target.value = ''
                  }}
                  type="file"
                />
              </label>

              <textarea
                className="min-h-11 flex-1 resize-none rounded-md border border-blue-200 bg-white px-3 py-2 text-sm text-blue-950 outline-none transition placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 disabled:bg-slate-100"
                disabled={isBusy}
                onChange={(event) => onInputChange(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' && !event.shiftKey) {
                    event.preventDefault()
                    onSubmit(input)
                  }
                }}
                placeholder="Nhap lenh hoac gui anh..."
                rows={2}
                value={input}
              />
              <button
                className="h-11 rounded-md bg-blue-950 px-4 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-900 disabled:cursor-default disabled:opacity-60"
                disabled={isBusy || (!input.trim() && images.length === 0)}
                type="submit"
              >
                Send
              </button>
            </div>
          </form>
          </div>

        </section>
      )}

      {!isOpen && (
        <button
          className="h-11 rounded-md bg-blue-950 px-4 text-sm font-semibold text-white shadow-lg ring-1 ring-white/10 transition hover:bg-blue-900"
          type="button"
          onClick={onToggle}
        >
          AI Control
        </button>
      )}
      </aside>

      {isOpen && isSettingsOpen && (
        <div
          aria-modal="true"
          className="fixed inset-0 z-[60] flex items-center justify-center bg-blue-950/35 p-4 backdrop-blur-sm"
          role="dialog"
        >
          <section className="w-[min(520px,calc(100vw-2rem))] overflow-hidden rounded-lg border border-blue-200 bg-white shadow-2xl ring-1 ring-blue-950/10">
            <div className="flex items-center justify-between gap-3 border-b border-blue-100 bg-blue-950 px-4 py-3 text-white">
              <div>
                <h3 className="m-0 text-sm font-bold">Codex Settings</h3>
                <p className="m-0 text-xs text-blue-100">Cai dat AI Control va dang nhap Codex</p>
              </div>
              <button className={chatHeaderButtonClass} type="button" onClick={onSettingsToggle}>
                Close
              </button>
            </div>

            <div className="space-y-4 p-4">
              <div className="grid gap-2 rounded-md border border-blue-100 bg-blue-50 p-3 text-xs text-slate-700 sm:grid-cols-2">
                <CodexStatusRow label="Bundled" value={codexStatus?.bundledAvailable ? 'Ready' : 'Missing'} />
                <CodexStatusRow label="Global" value={codexStatus?.globalAvailable ? 'Ready' : 'Missing'} />
                <CodexStatusRow label="Login" value={codexStatus?.hasApiKey || codexStatus?.hasStoredLogin ? 'Ready' : 'Needed'} />
                <CodexStatusRow label="Version" value={codexStatus?.version ?? '-'} />
                <CodexStatusRow label="AI Folder" value={codexStatus?.workspaceReady ? 'Ready' : 'Missing'} />
              </div>

              <div className="space-y-2">
                <label className="block text-xs font-semibold uppercase text-blue-950" htmlFor="ai-workspace-path">
                  AI Folder
                </label>
                <div className="flex flex-col gap-2 sm:flex-row">
                  <input
                    className="h-10 min-w-0 flex-1 rounded-md border border-blue-200 bg-white px-3 text-sm text-blue-950 outline-none transition placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-200 disabled:bg-slate-100"
                    disabled={isCodexActionRunning}
                    id="ai-workspace-path"
                    onChange={(event) => onAiWorkspaceInputChange(event.target.value)}
                    placeholder={codexStatus?.workspaceDefaultPath ?? 'AI folder path'}
                    type="text"
                    value={aiWorkspaceInput}
                  />
                  <button
                    className={chatSettingsButtonClass}
                    disabled={isCodexActionRunning || !aiWorkspaceInput.trim()}
                    type="button"
                    onClick={onCodexSaveWorkspace}
                  >
                    Save Folder
                  </button>
                </div>
                {codexStatus?.workspaceError && (
                  <p className="m-0 break-anywhere text-xs font-medium text-red-700">{codexStatus.workspaceError}</p>
                )}
              </div>

              <div className="grid gap-2 sm:grid-cols-2">
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexRefresh}
                >
                  Refresh
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexInstall}
                >
                  Install Codex
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexLogin}
                >
                  Login Codex
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexSelectWorkspace}
                >
                  Choose Folder
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexOpenWorkspace}
                >
                  Open Folder
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexResetWorkspace}
                >
                  Reset Folder
                </button>
                <button
                  className={chatSettingsButtonClass}
                  disabled={isCodexActionRunning}
                  type="button"
                  onClick={onCodexResetThread}
                >
                  Reset Thread
                </button>
              </div>

              {codexActionOutput && (
                <pre className="max-h-48 overflow-auto whitespace-pre-wrap rounded-md bg-blue-950 p-3 text-[11px] leading-relaxed text-blue-50">
                  {codexActionOutput}
                </pre>
              )}
            </div>
          </section>
        </div>
      )}
    </>
  )
}

function ChatAvatar({ label, tone }: { label: string; tone: 'assistant' | 'user' }) {
  return (
    <span
      className={`mt-1 flex h-7 w-7 shrink-0 items-center justify-center rounded-md text-[10px] font-black ${
        tone === 'assistant'
          ? 'border border-blue-200 bg-white text-blue-950 shadow-sm'
          : 'bg-blue-100 text-blue-900'
      }`}
    >
      {label}
    </span>
  )
}

function CodexStatusRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between gap-2">
      <span className="text-slate-500">{label}</span>
      <span className="truncate font-semibold text-blue-950">{value}</span>
    </div>
  )
}

const chatHeaderButtonClass =
  'h-8 rounded-md border border-white/15 bg-white/10 px-3 text-xs font-semibold text-white transition hover:bg-white/20 disabled:cursor-default disabled:opacity-50'

const chatSettingsButtonClass =
  'h-9 rounded-md border border-blue-200 bg-white px-3 text-sm font-semibold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50 disabled:cursor-default disabled:opacity-50'

function ExportProgressBar({ progress }: { progress: WordExportProgress }) {
  const safePercent = Math.min(100, Math.max(0, progress.percent))

  return (
    <section className="rounded-md border border-slate-200 bg-white p-3 shadow-sm" aria-label="Export progress">
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2 text-sm">
        <span className="font-semibold text-slate-950">{progress.message}</span>
        <span className="text-slate-600">
          {safePercent.toFixed(0)}% · {formatDuration(progress.elapsedMs ?? 0)}
        </span>
      </div>
      <div className="h-3 overflow-hidden rounded-full bg-slate-200">
        <div
          className={`h-full rounded-full transition-all duration-300 ${
            progress.completed ? 'bg-emerald-600' : 'bg-sky-600'
          }`}
          style={{ width: `${safePercent}%` }}
        />
      </div>
    </section>
  )
}

interface ProjectInformationSummaryProps {
  info: ProjectInfo
  selectedModuleModel: string | null
  onEdit: () => void
}

function ProjectInformationSummary({ info, selectedModuleModel, onEdit }: ProjectInformationSummaryProps) {
  const displayInfo = normalizeProjectInfoForDisplay(info)
  const standards = splitProjectInfoLines(displayInfo.applicableStandards)
  const moduleModel = normalizeDisplayText(selectedModuleModel ?? '')

  return (
    <section className="rounded-md border border-slate-200 bg-white p-4 shadow-sm" aria-label="Project information preview">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-2">
        <div>
          <h2 className="m-0 text-lg font-semibold text-slate-950">Information</h2>
          <p className="m-0 text-sm text-slate-500">Preview of the same front matter used in DOCX export.</p>
        </div>
        <button className={smallButtonClass} type="button" onClick={onEdit}>
          Edit Information
        </button>
      </div>

      <div className="overflow-x-auto">
        <div className="min-w-[820px] space-y-5 text-black" style={{ fontFamily: '"Times New Roman", Arial, serif' }}>
          <table className={docxPreviewTableClass}>
            <colgroup>
              <col className="w-[31%]" />
              <col className="w-[69%]" />
            </colgroup>
            <tbody>
              <tr>
                <CoverSignoffPreviewCell action="Prepared by" name={displayInfo.preparedBy} />
                <DocxPreviewCell className="px-5 text-center" rowSpan={3}>
                  <div className="text-xl font-bold uppercase">{displayInfo.companyName}</div>
                  <div className="mt-8 text-sm">{displayInfo.companyAddress}</div>
                  <div className="mt-5 text-2xl font-bold uppercase leading-tight">{displayInfo.reportTitle}</div>
                </DocxPreviewCell>
              </tr>
              <tr>
                <CoverSignoffPreviewCell action="Check by" name={displayInfo.checkedBy} />
              </tr>
              <tr>
                <CoverSignoffPreviewCell action="Approved by" name={displayInfo.approvedBy} />
              </tr>
            </tbody>
          </table>

          <table className={docxPreviewTableClass}>
            <colgroup>
              <col className="w-[30%]" />
              <col className="w-[70%]" />
            </colgroup>
            <tbody>
              <tr>
                <DocxPreviewCell className="whitespace-pre-line text-center text-lg">{'Nhà đầu tư\nInvestors'}</DocxPreviewCell>
                <DocxPreviewCell className="whitespace-pre-line text-center text-lg">
                  {[displayInfo.investorName, displayInfo.investorNameEnglish].filter(Boolean).join('\n')}
                </DocxPreviewCell>
              </tr>
              <tr>
                <DocxPreviewCell className="whitespace-pre-line text-center text-lg">{'Chủ nhà máy\nFactory Owner'}</DocxPreviewCell>
                <DocxPreviewCell className="whitespace-pre-line text-center text-lg">
                  {[displayInfo.factoryOwnerName, displayInfo.factoryOwnerNameEnglish].filter(Boolean).join('\n')}
                </DocxPreviewCell>
              </tr>
            </tbody>
          </table>

          <table className={docxPreviewTableClass}>
            <tbody>
              <tr>
                <DocxPreviewCell className="h-16 text-center text-3xl font-bold uppercase leading-tight">
                  {displayInfo.measurementTitle}
                </DocxPreviewCell>
              </tr>
              <tr>
                <DocxPreviewCell className="h-14 text-center text-3xl font-bold leading-tight text-red-600">
                  {moduleModel}
                </DocxPreviewCell>
              </tr>
            </tbody>
          </table>

          <p className="m-0 pl-2 text-base">PHÊ DUYỆT BỞI / APPROVAL BY</p>
          <table className={docxPreviewTableClass}>
            <tbody>
              <tr>
                <SignaturePreviewCell label="OWNER:" value={displayInfo.ownerApproval} />
                <SignaturePreviewCell label="CONSULTANT (Owner Engineer):" value={displayInfo.consultantApproval} />
              </tr>
              <tr>
                <SignaturePreviewCell label="CONTRACTOR EPC:" value={displayInfo.contractorEpcApproval} />
                <SignaturePreviewCell label="TESTER:" value={displayInfo.testerApproval} />
              </tr>
            </tbody>
          </table>

          <table className={docxPreviewTableClass}>
            <tbody>
              <tr>
                <DocxPreviewCell className="text-center font-bold">Tiêu chuẩn áp dụng</DocxPreviewCell>
                <DocxPreviewCell className="text-center font-bold">Applicable Standards</DocxPreviewCell>
              </tr>
              {standards.map((line, index) => (
                <tr key={`${line}-${index}`}>
                  <DocxPreviewCell className="text-center" colSpan={2}>
                    {line}
                  </DocxPreviewCell>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </section>
  )
}

function CoverSignoffPreviewCell({ action, name }: { action: string; name: string }) {
  return (
    <DocxPreviewCell className="h-20 px-3 py-1">
      <div className="grid h-full grid-cols-[5rem_1fr] grid-rows-3 items-center gap-x-2">
        <div className="self-start text-left">Date</div>
        <div />
        <div />
        <div className="text-center">{action}</div>
        <div className="self-end text-left">Engineer</div>
        <div className="self-end text-center">{name}</div>
      </div>
    </DocxPreviewCell>
  )
}

function SignaturePreviewCell({ label, value }: { label: string; value: string }) {
  const normalizedValue = normalizeDisplayText(value).trim()

  return (
    <DocxPreviewCell className="h-48 align-top">
      <div className="text-left font-bold italic underline">{label}</div>
      {normalizedValue && <div className="mt-16 text-center text-base font-normal not-italic no-underline">{normalizedValue}</div>}
    </DocxPreviewCell>
  )
}

interface DocxPreviewCellProps {
  children?: ReactNode
  className?: string
  colSpan?: number
  rowSpan?: number
}

function DocxPreviewCell({ children, className = '', colSpan, rowSpan }: DocxPreviewCellProps) {
  return (
    <td
      className={`border border-black px-2 py-1 align-middle text-base leading-tight ${className}`}
      colSpan={colSpan}
      rowSpan={rowSpan}
    >
      {children || <span className="invisible">.</span>}
    </td>
  )
}

const docxPreviewTableClass = 'w-full table-fixed border-collapse border border-black bg-white'

function splitProjectInfoLines(value: string) {
  const lines = normalizeDisplayText(value)
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)

  return lines.length > 0 ? lines : ['']
}

function normalizeProjectInfoForDisplay(info: ProjectInfo): ProjectInfo {
  return {
    projectName: normalizeDisplayText(info.projectName),
    investorName: normalizeDisplayText(info.investorName),
    investorNameEnglish: normalizeDisplayText(info.investorNameEnglish),
    factoryOwnerName: normalizeDisplayText(info.factoryOwnerName),
    factoryOwnerNameEnglish: normalizeDisplayText(info.factoryOwnerNameEnglish),
    reportTitle: normalizeDisplayText(info.reportTitle),
    measurementTitle: normalizeDisplayText(info.measurementTitle),
    companyName: normalizeDisplayText(info.companyName),
    companyAddress: normalizeDisplayText(info.companyAddress),
    preparedBy: normalizeDisplayText(info.preparedBy),
    checkedBy: normalizeDisplayText(info.checkedBy),
    approvedBy: normalizeDisplayText(info.approvedBy),
    ownerApproval: normalizeDisplayText(info.ownerApproval),
    consultantApproval: normalizeDisplayText(info.consultantApproval),
    contractorEpcApproval: normalizeDisplayText(info.contractorEpcApproval),
    testerApproval: normalizeDisplayText(info.testerApproval),
    applicableStandards: normalizeDisplayText(info.applicableStandards),
  }
}

function normalizeDisplayText(value: string) {
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

interface ProjectInformationPanelProps {
  info: ProjectInfo
  library: ProjectInfoLibraryItem[]
  selectedModuleModel: string | null
  status: string | null
  onApplyLibraryItem: (itemId: string) => void
  onDeleteLibraryItem: (itemId: string) => void
  onLoadReference: () => void
  onSaveToLibrary: () => void
  onUpdate: (key: keyof ProjectInfo, value: string) => void
}

function ProjectInformationPanel({
  info,
  library,
  selectedModuleModel,
  status,
  onApplyLibraryItem,
  onDeleteLibraryItem,
  onLoadReference,
  onSaveToLibrary,
  onUpdate,
}: ProjectInformationPanelProps) {
  const sortedLibrary = sortProjectInfoLibrary(library)

  return (
    <section className="grid gap-4 rounded-md border border-slate-200 bg-white p-4 shadow-sm xl:grid-cols-[minmax(0,1fr)_minmax(320px,420px)]">
      <div className="min-w-0">
        <div className="mb-4 flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="m-0 text-lg font-semibold text-slate-950">Project Information</h2>
            <p className="m-0 text-sm text-slate-500">Front matter used at the beginning of the exported Word report.</p>
          </div>
          <div className="flex flex-wrap justify-end gap-2">
            <button className={smallButtonClass} type="button" onClick={onLoadReference}>
              Load Sample
            </button>
            <button className={smallButtonClass} type="button" onClick={onSaveToLibrary}>
              Save to Library
            </button>
          </div>
        </div>

        {status && <p className="m-0 mb-3 text-sm font-medium text-emerald-700">{status}</p>}

        <div className="grid gap-3 md:grid-cols-2">
          <ProjectInfoField
            label="Project Name"
            onChange={(value) => onUpdate('projectName', value)}
            placeholder="Rooftop solar project"
            value={info.projectName}
          />
          <ProjectInfoField
            label="Report Title"
            onChange={(value) => onUpdate('reportTitle', value)}
            value={info.reportTitle}
          />
          <ProjectInfoField
            label="Measurement Title"
            onChange={(value) => onUpdate('measurementTitle', value)}
            value={info.measurementTitle}
          />
          <ProjectInfoField
            label="Company"
            onChange={(value) => onUpdate('companyName', value)}
            value={info.companyName}
          />
          <ProjectInfoField
            label="Prepared By"
            onChange={(value) => onUpdate('preparedBy', value)}
            value={info.preparedBy}
          />
          <ProjectInfoField
            label="Checked By"
            onChange={(value) => onUpdate('checkedBy', value)}
            value={info.checkedBy}
          />
          <ProjectInfoField
            label="Approved By"
            onChange={(value) => onUpdate('approvedBy', value)}
            value={info.approvedBy}
          />
          <ProjectInfoField
            label="Current PV Model"
            readOnly
            value={selectedModuleModel ?? ''}
          />
          <ProjectInfoField
            label="Investor"
            onChange={(value) => onUpdate('investorName', value)}
            value={info.investorName}
          />
          <ProjectInfoField
            label="Investor English"
            onChange={(value) => onUpdate('investorNameEnglish', value)}
            value={info.investorNameEnglish}
          />
          <ProjectInfoField
            label="Factory Owner"
            onChange={(value) => onUpdate('factoryOwnerName', value)}
            value={info.factoryOwnerName}
          />
          <ProjectInfoField
            label="Factory Owner English"
            onChange={(value) => onUpdate('factoryOwnerNameEnglish', value)}
            value={info.factoryOwnerNameEnglish}
          />
          <ProjectInfoField
            label="Owner Approval"
            onChange={(value) => onUpdate('ownerApproval', value)}
            value={info.ownerApproval}
          />
          <ProjectInfoField
            label="Consultant Approval"
            onChange={(value) => onUpdate('consultantApproval', value)}
            value={info.consultantApproval}
          />
          <ProjectInfoField
            label="Contractor EPC Approval"
            onChange={(value) => onUpdate('contractorEpcApproval', value)}
            value={info.contractorEpcApproval}
          />
          <ProjectInfoField
            label="Tester Approval"
            onChange={(value) => onUpdate('testerApproval', value)}
            value={info.testerApproval}
          />
          <ProjectInfoTextareaField
            label="Company Address"
            onChange={(value) => onUpdate('companyAddress', value)}
            value={info.companyAddress}
          />
          <ProjectInfoTextareaField
            label="Applicable Standards"
            onChange={(value) => onUpdate('applicableStandards', value)}
            value={info.applicableStandards}
          />
        </div>
      </div>

      <aside className="min-w-0 rounded-md border border-blue-100 bg-blue-50 p-3">
        <div className="mb-3 flex items-center justify-between gap-2">
          <div>
            <h3 className="m-0 text-base font-semibold text-blue-950">Project Info Library</h3>
            <p className="m-0 text-sm text-slate-500">{sortedLibrary.length} saved template(s)</p>
          </div>
        </div>

        {sortedLibrary.length > 0 ? (
          <div className="flex flex-col gap-2">
            <select
              className={selectClass}
              onChange={(event) => {
                if (event.target.value) {
                  onApplyLibraryItem(event.target.value)
                  event.target.value = ''
                }
              }}
              value=""
            >
              <option value="">Load from library...</option>
              {sortedLibrary.map((item) => (
                <option key={item.id} value={item.id}>
                  {getProjectInfoLibraryLabel(item)}
                </option>
              ))}
            </select>

            <div className="max-h-72 space-y-2 overflow-auto pr-1">
              {sortedLibrary.map((item) => (
                <div className="rounded-md border border-blue-100 bg-white p-3 shadow-sm" key={item.id}>
                  <div className="mb-2 min-w-0">
                    <p className="m-0 truncate text-sm font-semibold text-blue-950">
                      {getProjectInfoLibraryLabel(item)}
                    </p>
                    <p className="m-0 truncate text-xs text-slate-500">
                      {item.investorName || item.projectName || 'No investor'}
                    </p>
                  </div>
                  <div className="flex justify-end gap-2">
                    <button className={smallButtonClass} type="button" onClick={() => onApplyLibraryItem(item.id)}>
                      Apply
                    </button>
                    <button className={dangerButtonClass} type="button" onClick={() => onDeleteLibraryItem(item.id)}>
                      Delete
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        ) : (
          <div className="rounded-md border border-dashed border-blue-200 bg-white px-3 py-6 text-center text-sm text-slate-500">
            No project information saved yet.
          </div>
        )}
      </aside>
    </section>
  )
}

interface ProjectInfoFieldProps {
  label: string
  onChange?: (value: string) => void
  placeholder?: string
  readOnly?: boolean
  value: string
}

function ProjectInfoField({ label, onChange, placeholder, readOnly, value }: ProjectInfoFieldProps) {
  return (
    <label className="flex min-w-0 flex-col gap-1.5">
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <input
        className={fieldInputClass}
        onChange={(event) => onChange?.(event.target.value)}
        placeholder={placeholder}
        readOnly={readOnly}
        value={value}
      />
    </label>
  )
}

interface ProjectInfoTextareaFieldProps {
  label: string
  onChange: (value: string) => void
  value: string
}

function ProjectInfoTextareaField({ label, onChange, value }: ProjectInfoTextareaFieldProps) {
  return (
    <label className="flex min-w-0 flex-col gap-1.5 md:col-span-2">
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <textarea
        className="min-h-20 w-full resize-y rounded-md border border-blue-200 bg-white px-3 py-2 text-sm text-blue-950 shadow-sm outline-none transition placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-200"
        onChange={(event) => onChange(event.target.value)}
        value={value}
      />
    </label>
  )
}

interface PvModuleFormState {
  model: string
  ratedMaximumPowerW: string
  openCircuitVoltageV: string
  maximumPowerVoltageV: string
  shortCircuitCurrentA: string
  maximumPowerCurrentA: string
  moduleEfficiencyPercent: string
  powerTolerance: string
  firstYearDegradationPercent: string
  annualDegradationPercent: string
  temperatureCoefficientIscPercentPerC: string
  temperatureCoefficientVocPercentPerC: string
  temperatureCoefficientPmaxPercentPerC: string
}

const emptyPvModuleForm: PvModuleFormState = {
  model: '',
  ratedMaximumPowerW: '',
  openCircuitVoltageV: '',
  maximumPowerVoltageV: '',
  shortCircuitCurrentA: '',
  maximumPowerCurrentA: '',
  moduleEfficiencyPercent: '',
  powerTolerance: '',
  firstYearDegradationPercent: defaultFirstYearDegradationPercent.toString(),
  annualDegradationPercent: defaultAnnualDegradationPercent.toString(),
  temperatureCoefficientIscPercentPerC: '',
  temperatureCoefficientVocPercentPerC: '',
  temperatureCoefficientPmaxPercentPerC: '',
}

function PvModuleLibrary({
  isLoading,
  modules,
  onModuleChange,
  onSave,
  selectedModuleId,
  selectedModuleModel,
  status,
}: PvModuleLibraryProps) {
  const [editingId, setEditingId] = useState<string | null>(null)
  const [displayModuleId, setDisplayModuleId] = useState<string | null>(null)
  const [form, setForm] = useState<PvModuleFormState>(emptyPvModuleForm)
  const [isSaving, setIsSaving] = useState(false)
  const editingModule = modules.find((pvModule) => pvModule.id === editingId) ?? null
  const selectedDisplayModule =
    editingModule ??
    modules.find((pvModule) => pvModule.id === displayModuleId) ??
    modules.find((pvModule) => pvModule.model === selectedModuleModel) ??
    modules[0] ??
    null

  const updateForm = (key: keyof PvModuleFormState, value: string) => {
    setForm((currentForm) => ({ ...currentForm, [key]: value }))
  }

  const resetForm = () => {
    setEditingId(null)
    setForm(emptyPvModuleForm)
  }

  const handleEdit = (pvModule: PvModule) => {
    setDisplayModuleId(pvModule.id)
    setEditingId(pvModule.id)
    setForm(createPvModuleForm(pvModule))
  }

  const handleSelectedModuleChange = (moduleId: string) => {
    onModuleChange(moduleId)

    if (moduleId) {
      setDisplayModuleId(moduleId)
    }
  }

  const handleDelete = async (pvModule: PvModule) => {
    if (!window.confirm(`Delete PV module "${pvModule.model}"?`)) {
      return
    }

    setIsSaving(true)

    try {
      await onSave(modules.filter((moduleItem) => moduleItem.id !== pvModule.id))

      if (editingId === pvModule.id) {
        resetForm()
      }

      if (displayModuleId === pvModule.id) {
        setDisplayModuleId(null)
      }
    } finally {
      setIsSaving(false)
    }
  }

  const handleSubmit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()

    if (!form.model.trim()) {
      return
    }

    setIsSaving(true)

    try {
      const nextModule = createPvModuleFromForm(form, editingModule)
      const nextModules = editingModule
        ? modules.map((moduleItem) => (moduleItem.id === editingModule.id ? nextModule : moduleItem))
        : [...modules, nextModule]

      await onSave(sortPvModules(nextModules))
      setDisplayModuleId(nextModule.id)
      resetForm()
    } finally {
      setIsSaving(false)
    }
  }

  return (
    <section className="grid gap-4 rounded-md border border-slate-200 bg-white p-4 shadow-sm xl:grid-cols-[minmax(360px,1fr)_minmax(420px,520px)]">
      <div className="min-w-0">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="m-0 text-lg font-semibold text-slate-950">PV Module Library</h2>
            <p className="m-0 text-sm text-slate-500">
              {isLoading ? 'Loading module library...' : `${modules.length} module type(s)`}
            </p>
          </div>
          {status && <span className="text-sm font-medium text-emerald-700">{status}</span>}
        </div>

        <div className="mb-3 flex flex-wrap gap-2">
          {modules.map((pvModule) => {
            const isSelected = selectedDisplayModule?.id === pvModule.id
            const isCurrentModule = selectedModuleModel === pvModule.model

            return (
              <button
                className={[
                  'rounded-md border px-3 py-2 text-left text-sm font-semibold transition',
                  isSelected ? 'border-blue-950 bg-blue-950 text-white' : 'border-blue-200 bg-white text-blue-950 hover:bg-blue-50',
                ].join(' ')}
                key={pvModule.id}
                onClick={() => setDisplayModuleId(pvModule.id)}
                type="button"
              >
                {pvModule.model}
                {isCurrentModule && <span className="ml-2 text-xs font-normal">(current)</span>}
              </button>
            )
          })}
        </div>

        {selectedDisplayModule ? (
          <>
            <PvModuleInformationTable pvModule={selectedDisplayModule} />
            <div className="mt-3 flex justify-end gap-2">
              <button
                className="h-9 rounded-md bg-blue-950 px-3 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-900 disabled:cursor-default disabled:bg-emerald-700"
                disabled={selectedModuleId === selectedDisplayModule.id}
                type="button"
                onClick={() => handleSelectedModuleChange(selectedDisplayModule.id)}
              >
                {selectedModuleId === selectedDisplayModule.id ? 'Used for report' : 'Use for report'}
              </button>
              <button className={smallButtonClass} type="button" onClick={() => handleEdit(selectedDisplayModule)}>
                Edit
              </button>
              <button className={dangerButtonClass} type="button" onClick={() => handleDelete(selectedDisplayModule)}>
                Delete
              </button>
            </div>
          </>
        ) : (
          <div className="rounded-md border border-slate-200 px-3 py-6 text-center text-sm text-slate-500">
            No PV modules in library.
          </div>
        )}
      </div>

      <form className="rounded-md border border-slate-200 bg-slate-50 p-4" onSubmit={handleSubmit}>
        <div className="mb-4 flex items-center justify-between gap-3">
          <div>
            <h3 className="m-0 text-base font-semibold text-slate-950">
              {editingModule ? 'Edit PV Module' : 'Add PV Module'}
            </h3>
            <p className="m-0 text-sm text-slate-500">Electrical ratings used for PV module reference.</p>
          </div>
          {editingModule && (
            <button className={smallButtonClass} type="button" onClick={resetForm}>
              New
            </button>
          )}
        </div>

        <div className="grid gap-3 sm:grid-cols-2">
          <label className="flex flex-col gap-1.5 sm:col-span-2">
            <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Model</span>
            <input
              className={fieldInputClass}
              onChange={(event) => updateForm('model', event.target.value)}
              placeholder="JAM72S30-540/MR/1500V"
              required
              value={form.model}
            />
          </label>

          <PvModuleNumberField
            label="Rated Maximum Power(Pmax) [W]"
            onChange={(value) => updateForm('ratedMaximumPowerW', value)}
            value={form.ratedMaximumPowerW}
          />
          <PvModuleNumberField
            label="Open Circuit Voltage(Voc) [V]"
            onChange={(value) => updateForm('openCircuitVoltageV', value)}
            value={form.openCircuitVoltageV}
          />
          <PvModuleNumberField
            label="Maximum Power Voltage(Vmp) [V]"
            onChange={(value) => updateForm('maximumPowerVoltageV', value)}
            value={form.maximumPowerVoltageV}
          />
          <PvModuleNumberField
            label="Short Circuit Current(Isc) [A]"
            onChange={(value) => updateForm('shortCircuitCurrentA', value)}
            value={form.shortCircuitCurrentA}
          />
          <PvModuleNumberField
            label="Maximum Power Current(Imp) [A]"
            onChange={(value) => updateForm('maximumPowerCurrentA', value)}
            value={form.maximumPowerCurrentA}
          />
          <PvModuleNumberField
            label="Module Efficiency [%]"
            onChange={(value) => updateForm('moduleEfficiencyPercent', value)}
            value={form.moduleEfficiencyPercent}
          />

          <label className="flex flex-col gap-1.5 sm:col-span-2">
            <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">Power Tolerance</span>
            <input
              className={fieldInputClass}
              onChange={(event) => updateForm('powerTolerance', event.target.value)}
              placeholder="0~+5W"
              value={form.powerTolerance}
            />
          </label>

          <PvModuleNumberField
            label="1st Degradation [%]"
            onChange={(value) => updateForm('firstYearDegradationPercent', value)}
            value={form.firstYearDegradationPercent}
          />
          <PvModuleNumberField
            label="Degradation Per Year [%]"
            onChange={(value) => updateForm('annualDegradationPercent', value)}
            value={form.annualDegradationPercent}
          />

          <PvModuleNumberField
            label="Temperature Coefficient Isc alpha [%/C]"
            onChange={(value) => updateForm('temperatureCoefficientIscPercentPerC', value)}
            value={form.temperatureCoefficientIscPercentPerC}
          />
          <PvModuleNumberField
            label="Temperature Coefficient Voc beta [%/C]"
            onChange={(value) => updateForm('temperatureCoefficientVocPercentPerC', value)}
            value={form.temperatureCoefficientVocPercentPerC}
          />
          <PvModuleNumberField
            label="Temperature Coefficient Pmax gamma [%/C]"
            onChange={(value) => updateForm('temperatureCoefficientPmaxPercentPerC', value)}
            value={form.temperatureCoefficientPmaxPercentPerC}
          />
        </div>

        <div className="mt-4 flex justify-end gap-2">
          <button className={smallButtonClass} type="button" onClick={resetForm}>
            Clear
          </button>
          <button
            className="h-10 rounded-md bg-blue-950 px-4 text-sm font-semibold text-white shadow-sm transition hover:bg-blue-900 disabled:cursor-default disabled:opacity-60"
            disabled={isSaving || !form.model.trim()}
            type="submit"
          >
            {isSaving ? 'Saving...' : editingModule ? 'Save Changes' : 'Add Module'}
          </button>
        </div>
      </form>
    </section>
  )
}

function PvModuleInformationTable({ pvModule }: { pvModule: PvModule }) {
  const rows: Array<{
    label: ReactNode
    value: string
  }> = [
    { label: 'Model', value: pvModule.model },
    { label: 'Rated Maximum Power(Pmax) [W]', value: formatPvModuleInfoValue(pvModule.ratedMaximumPowerW) },
    {
      label: (
        <>
          Open Circuit Voltage(Voc) [V] <strong>(Tolerance ± 3%)</strong>
        </>
      ),
      value: formatPvModuleInfoValue(pvModule.openCircuitVoltageV),
    },
    { label: 'Maximum Power Voltage(Vmp) [V]', value: formatPvModuleInfoValue(pvModule.maximumPowerVoltageV) },
    {
      label: (
        <>
          Short Circuit Current(Isc) [A] <strong>(Tolerance ± 3%)</strong>
        </>
      ),
      value: formatPvModuleInfoValue(pvModule.shortCircuitCurrentA),
    },
    { label: 'Maximum Power Current(Imp) [A]', value: formatPvModuleInfoValue(pvModule.maximumPowerCurrentA) },
    { label: 'Module Efficiency [%]', value: formatPvModuleInfoValue(pvModule.moduleEfficiencyPercent) },
    { label: 'Power Tolerance', value: pvModule.powerTolerance || '-' },
    { label: '1st Degradation [%]', value: formatPvModuleInfoValue(pvModule.firstYearDegradationPercent) },
    { label: 'Degradation Per Year [%]', value: formatPvModuleInfoValue(pvModule.annualDegradationPercent) },
    {
      label: 'Temperature Coefficient of Isc(Alpha±_Isc) [°C]',
      value: formatPvModuleInfoValue(pvModule.temperatureCoefficientIscPercentPerC),
    },
    {
      label: 'Temperature Coefficient of Voc(Beta_Voc) [°C]',
      value: formatPvModuleInfoValue(pvModule.temperatureCoefficientVocPercentPerC),
    },
    {
      label: 'Temperature Coefficient of Pmax(Gama_Pmp) [°C]',
      value: formatPvModuleInfoValue(pvModule.temperatureCoefficientPmaxPercentPerC),
    },
  ]

  return (
    <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white p-1 shadow-sm">
      <table className="w-full min-w-[720px] border-collapse table-fixed font-sans text-slate-950">
        <thead>
          <tr>
            <th className="border border-slate-500 bg-[#1f4e79] px-3 py-2 text-center text-2xl font-bold text-white" colSpan={2}>
              PV INFORMATION
            </th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row, index) => {
            const fillClass = index % 2 === 0 ? 'bg-[#eaf2f8]' : 'bg-white'

            return (
              <tr className="group transition-colors" key={index}>
                <td className={`w-[64%] border border-slate-500 px-3 py-2 text-lg leading-tight transition-[background-color,filter] group-hover:brightness-[0.98] ${fillClass}`}>
                  {row.label}
                </td>
                <td className={`border border-slate-500 px-3 py-2 text-center text-lg leading-tight transition-[background-color,filter] group-hover:brightness-[0.98] ${fillClass}`}>
                  {row.value}
                </td>
              </tr>
            )
          })}
        </tbody>
      </table>
    </div>
  )
}

interface PvModuleNumberFieldProps {
  label: string
  onChange: (value: string) => void
  value: string
}

function PvModuleNumberField({ label, onChange, value }: PvModuleNumberFieldProps) {
  return (
    <label className="flex flex-col gap-1.5">
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <input
        className={fieldInputClass}
        inputMode="decimal"
        onChange={(event) => onChange(event.target.value)}
        type="number"
        step="any"
        value={value}
      />
    </label>
  )
}

interface ModuleSelectorProps {
  selector: ModuleSelectorState
  onSystemChange: (systemGroup: string) => void
  onInverterChange: (inverter: string) => void
  onRecordChange: (recordPath: string) => void
  onPrevious: () => void
  onNext: () => void
  canMovePrevious: boolean
  canMoveNext: boolean
  selectedIndex: number
  totalRecords: number
}

function ModuleSelector({
  selector,
  onSystemChange,
  onInverterChange,
  onRecordChange,
  onPrevious,
  onNext,
  canMovePrevious,
  canMoveNext,
  selectedIndex,
  totalRecords,
}: ModuleSelectorProps) {
  return (
    <div
      className="grid items-end gap-2 rounded-md border border-slate-200 bg-white p-3 shadow-sm lg:grid-cols-[44px_minmax(120px,160px)_minmax(150px,220px)_minmax(220px,360px)_44px_auto]"
      aria-label="Module selector"
    >
      <button
        className={navButtonClass}
        type="button"
        aria-label="Previous record"
        disabled={!canMovePrevious}
        onClick={onPrevious}
      >
        &lt;
      </button>

      <LevelSelector
        label="X"
        value={selector.selectedSystemGroup}
        options={selector.systemGroups.map((systemGroup) => ({
          label: systemGroup,
          value: systemGroup,
        }))}
        onChange={onSystemChange}
      />

      <LevelSelector
        label="Inverter"
        value={selector.selectedInverter}
        options={selector.inverters.map((inverter) => ({
          label: inverter,
          value: inverter,
        }))}
        onChange={onInverterChange}
      />

      <LevelSelector
        label="String"
        value={selector.selectedRecord?.sourcePath ?? ''}
        options={selector.records.map((record) => ({
          label: formatStringOption(record),
          value: record.sourcePath,
        }))}
        onChange={onRecordChange}
      />

      <button
        className={navButtonClass}
        type="button"
        aria-label="Next record"
        disabled={!canMoveNext}
        onClick={onNext}
      >
        &gt;
      </button>

      <span className="flex min-h-11 items-center text-sm font-medium text-slate-500">
        {selectedIndex + 1}/{totalRecords}
      </span>
    </div>
  )
}

interface LevelSelectorProps {
  label: string
  value: string
  options: Array<{
    label: string
    value: string
  }>
  onChange: (value: string) => void
}

const selectClass =
  'h-11 w-full rounded-md border border-blue-200 bg-white px-3 text-sm font-medium text-blue-950 shadow-sm outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-200'

const fieldInputClass =
  'h-10 w-full rounded-md border border-blue-200 bg-white px-3 text-sm text-blue-950 shadow-sm outline-none transition placeholder:text-slate-400 focus:border-blue-500 focus:ring-2 focus:ring-blue-200'

const navButtonClass =
  'h-11 w-11 rounded-md border border-blue-200 bg-white text-base font-bold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50 focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-blue-400 disabled:cursor-default disabled:text-slate-300 disabled:hover:border-slate-300 disabled:hover:bg-white'

const viewTabClass = (isActive: boolean) =>
  [
    'h-9 rounded px-4 text-sm font-semibold transition focus-visible:outline-3 focus-visible:outline-offset-2 focus-visible:outline-blue-400',
    isActive ? 'bg-blue-950 text-white shadow-sm' : 'text-blue-800 hover:text-blue-950',
  ].join(' ')

function getAppViewLabel(view: AppView) {
  if (view === 'project-info') {
    return 'Information'
  }

  return view === 'pv-module' ? 'PV Module' : 'Home'
}

const smallButtonClass =
  'h-9 rounded-md border border-blue-200 bg-white px-3 text-sm font-semibold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50'

const compactToolButtonClass = (isActive: boolean) =>
  [
    'grid h-7 w-7 place-items-center rounded border shadow-sm transition disabled:cursor-not-allowed disabled:opacity-40',
    isActive ? 'border-blue-500 bg-blue-50 text-blue-950' : 'border-slate-200 bg-white text-slate-700 hover:border-blue-300 hover:bg-blue-50',
  ].join(' ')

const dangerButtonClass =
  'h-9 rounded-md border border-red-200 bg-white px-3 text-sm font-semibold text-red-700 shadow-sm transition hover:border-red-300 hover:bg-red-50'

type ToolIconName = 'clear' | 'move' | 'polygon' | 'rectangle' | 'trash' | 'undo' | 'upload' | 'width'

function ToolIcon({ name }: { name: ToolIconName }) {
  const commonProps = {
    'aria-hidden': true,
    className: 'h-4 w-4',
    fill: 'none',
    stroke: 'currentColor',
    strokeLinecap: 'round' as const,
    strokeLinejoin: 'round' as const,
    strokeWidth: 2,
    viewBox: '0 0 24 24',
  }

  if (name === 'upload') {
    return (
      <svg {...commonProps}>
        <path d="M12 15V3" />
        <path d="m7 8 5-5 5 5" />
        <path d="M5 21h14" />
      </svg>
    )
  }

  if (name === 'clear') {
    return (
      <svg {...commonProps}>
        <path d="M18 6 6 18" />
        <path d="m6 6 12 12" />
      </svg>
    )
  }

  if (name === 'move') {
    return (
      <svg {...commonProps}>
        <path d="M12 3v18" />
        <path d="M3 12h18" />
        <path d="m8 7 4-4 4 4" />
        <path d="m8 17 4 4 4-4" />
        <path d="m7 8-4 4 4 4" />
        <path d="m17 8 4 4-4 4" />
      </svg>
    )
  }

  if (name === 'polygon') {
    return (
      <svg {...commonProps}>
        <path d="M5 17 9 5l10 4-3 10Z" />
        <path d="M5 17 16 19" />
      </svg>
    )
  }

  if (name === 'rectangle') {
    return (
      <svg {...commonProps}>
        <rect height="12" rx="1.5" width="16" x="4" y="6" />
      </svg>
    )
  }

  if (name === 'undo') {
    return (
      <svg {...commonProps}>
        <path d="m9 7-5 5 5 5" />
        <path d="M20 18v-2a4 4 0 0 0-4-4H4" />
      </svg>
    )
  }

  if (name === 'trash') {
    return (
      <svg {...commonProps}>
        <path d="M4 7h16" />
        <path d="M10 11v6" />
        <path d="M14 11v6" />
        <path d="M6 7l1 14h10l1-14" />
        <path d="M9 7V4h6v3" />
      </svg>
    )
  }

  return (
    <svg {...commonProps}>
      <path d="M5 7h14" />
      <path d="M5 12h14" strokeWidth={3} />
      <path d="M5 18h14" strokeWidth={4} />
    </svg>
  )
}

function LevelSelector({ label, value, options, onChange }: LevelSelectorProps) {
  return (
    <label className="flex min-w-0 flex-col gap-1.5">
      <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      <select className={selectClass} value={value} onChange={(event) => onChange(event.target.value)}>
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
  )
}

interface ModuleSelectorState {
  systemGroups: string[]
  selectedSystemGroup: string
  inverters: string[]
  selectedInverter: string
  records: MeasurementRecord[]
  selectedRecord: MeasurementRecord | null
}

interface ModuleSelection {
  systemGroup: string
  inverter: string
  recordPath: string
}

const naturalCollator = new Intl.Collator(undefined, {
  numeric: true,
  sensitivity: 'base',
})

function createPvModuleForm(pvModule: PvModule): PvModuleFormState {
  return {
    model: pvModule.model,
    ratedMaximumPowerW: formatPvModuleInputValue(pvModule.ratedMaximumPowerW),
    openCircuitVoltageV: formatPvModuleInputValue(pvModule.openCircuitVoltageV),
    maximumPowerVoltageV: formatPvModuleInputValue(pvModule.maximumPowerVoltageV),
    shortCircuitCurrentA: formatPvModuleInputValue(pvModule.shortCircuitCurrentA),
    maximumPowerCurrentA: formatPvModuleInputValue(pvModule.maximumPowerCurrentA),
    moduleEfficiencyPercent: formatPvModuleInputValue(pvModule.moduleEfficiencyPercent),
    powerTolerance: pvModule.powerTolerance,
    firstYearDegradationPercent: formatPvModuleInputValue(pvModule.firstYearDegradationPercent),
    annualDegradationPercent: formatPvModuleInputValue(pvModule.annualDegradationPercent),
    temperatureCoefficientIscPercentPerC: formatPvModuleInputValue(
      pvModule.temperatureCoefficientIscPercentPerC,
    ),
    temperatureCoefficientVocPercentPerC: formatPvModuleInputValue(
      pvModule.temperatureCoefficientVocPercentPerC,
    ),
    temperatureCoefficientPmaxPercentPerC: formatPvModuleInputValue(
      pvModule.temperatureCoefficientPmaxPercentPerC,
    ),
  }
}

function createPvModuleFromForm(form: PvModuleFormState, existingModule: PvModule | null): PvModule {
  const now = new Date().toISOString()

  return {
    id: existingModule?.id ?? createPvModuleId(),
    model: form.model.trim(),
    ratedMaximumPowerW: parsePvModuleNumber(form.ratedMaximumPowerW),
    openCircuitVoltageV: parsePvModuleNumber(form.openCircuitVoltageV),
    maximumPowerVoltageV: parsePvModuleNumber(form.maximumPowerVoltageV),
    shortCircuitCurrentA: parsePvModuleNumber(form.shortCircuitCurrentA),
    maximumPowerCurrentA: parsePvModuleNumber(form.maximumPowerCurrentA),
    moduleEfficiencyPercent: parsePvModuleNumber(form.moduleEfficiencyPercent),
    powerTolerance: form.powerTolerance.trim(),
    firstYearDegradationPercent:
      parsePvModuleNumber(form.firstYearDegradationPercent) ?? defaultFirstYearDegradationPercent,
    annualDegradationPercent:
      parsePvModuleNumber(form.annualDegradationPercent) ?? defaultAnnualDegradationPercent,
    temperatureCoefficientIscPercentPerC: parsePvModuleNumber(form.temperatureCoefficientIscPercentPerC),
    temperatureCoefficientVocPercentPerC: parsePvModuleNumber(form.temperatureCoefficientVocPercentPerC),
    temperatureCoefficientPmaxPercentPerC: parsePvModuleNumber(form.temperatureCoefficientPmaxPercentPerC),
    createdAt: existingModule?.createdAt ?? now,
    updatedAt: now,
  }
}

function createPvModuleFromChatPayload(payload: ChatPvModulePayload, existingModule: PvModule | null): PvModule {
  const now = new Date().toISOString()

  return {
    id: existingModule?.id ?? createPvModuleId(),
    model: payload.model.trim(),
    ratedMaximumPowerW: normalizeChatPvModuleNumber(payload.ratedMaximumPowerW),
    openCircuitVoltageV: normalizeChatPvModuleNumber(payload.openCircuitVoltageV),
    maximumPowerVoltageV: normalizeChatPvModuleNumber(payload.maximumPowerVoltageV),
    shortCircuitCurrentA: normalizeChatPvModuleNumber(payload.shortCircuitCurrentA),
    maximumPowerCurrentA: normalizeChatPvModuleNumber(payload.maximumPowerCurrentA),
    moduleEfficiencyPercent: normalizeChatPvModuleNumber(payload.moduleEfficiencyPercent),
    powerTolerance: payload.powerTolerance?.trim() ?? existingModule?.powerTolerance ?? '',
    firstYearDegradationPercent:
      normalizeChatPvModuleNumber(payload.firstYearDegradationPercent) ??
      existingModule?.firstYearDegradationPercent ??
      defaultFirstYearDegradationPercent,
    annualDegradationPercent:
      normalizeChatPvModuleNumber(payload.annualDegradationPercent) ??
      existingModule?.annualDegradationPercent ??
      defaultAnnualDegradationPercent,
    temperatureCoefficientIscPercentPerC: normalizeChatPvModuleNumber(payload.temperatureCoefficientIscPercentPerC),
    temperatureCoefficientVocPercentPerC: normalizeChatPvModuleNumber(payload.temperatureCoefficientVocPercentPerC),
    temperatureCoefficientPmaxPercentPerC: normalizeChatPvModuleNumber(payload.temperatureCoefficientPmaxPercentPerC),
    createdAt: existingModule?.createdAt ?? now,
    updatedAt: now,
  }
}

function normalizeChatPvModuleNumber(value: number | null) {
  return typeof value === 'number' && Number.isFinite(value) ? value : null
}

function createPvModuleId() {
  return globalThis.crypto?.randomUUID?.() ?? `pv-module-${Date.now()}`
}

function parsePvModuleNumber(value: string) {
  if (!value.trim()) {
    return null
  }

  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

function formatPvModuleInputValue(value: number | null) {
  return value === null ? '' : value.toString()
}

function sortPvModules(modules: PvModule[]) {
  return [...modules].sort((left, right) => naturalCollator.compare(left.model, right.model))
}

function formatPvModuleInfoValue(value: number | null) {
  if (value === null) {
    return '-'
  }

  return Number.isInteger(value) ? value.toString() : value.toString()
}

function buildModuleSelector(
  records: MeasurementRecord[],
  systemGroup: string,
  inverter: string,
  recordPath: string,
): ModuleSelectorState {
  const systemGroups = uniqueSorted(records.map(getSystemGroupKey))
  const selectedSystemGroup = systemGroups.includes(systemGroup) ? systemGroup : systemGroups[0] ?? ''
  const systemRecords = records.filter((record) => getSystemGroupKey(record) === selectedSystemGroup)
  const inverters = uniqueSorted(systemRecords.map(getInverterKey))
  const selectedInverter = inverters.includes(inverter) ? inverter : inverters[0] ?? ''
  const inverterRecords = systemRecords.filter((record) => getInverterKey(record) === selectedInverter)
  const sortedRecords = [...inverterRecords].sort((left, right) => {
    return naturalCollator.compare(getStringKey(left), getStringKey(right))
  })
  const selectedRecord =
    sortedRecords.find((record) => record.sourcePath === recordPath) ?? sortedRecords[0] ?? null

  return {
    systemGroups,
    selectedSystemGroup,
    inverters,
    selectedInverter,
    records: sortedRecords,
    selectedRecord,
  }
}

function sortRecordsByHierarchy(records: MeasurementRecord[]) {
  return [...records].sort((left, right) => {
    return (
      naturalCollator.compare(getSystemGroupKey(left), getSystemGroupKey(right)) ||
      naturalCollator.compare(getInverterKey(left), getInverterKey(right)) ||
      naturalCollator.compare(getStringKey(left), getStringKey(right)) ||
      naturalCollator.compare(left.relativePath, right.relativePath)
    )
  })
}

function getInitialModuleSelection(
  records: MeasurementRecord[],
  systemGroup = '',
  inverter = '',
): ModuleSelection {
  const selector = buildModuleSelector(records, systemGroup, inverter, '')

  return {
    systemGroup: selector.selectedSystemGroup,
    inverter: selector.selectedInverter,
    recordPath: selector.selectedRecord?.sourcePath ?? '',
  }
}

function getRestoredModuleSelection(records: MeasurementRecord[], session: AppSessionState): ModuleSelection {
  const restoredRecord =
    records.find((record) => record.sourcePath === session.selectedRecordPath) ??
    records.find((record) => record.relativePath === session.selectedRecordRelativePath) ??
    null

  if (!restoredRecord) {
    return getInitialModuleSelection(records, session.selectedSystemGroup, session.selectedInverter)
  }

  return {
    systemGroup: getSystemGroupKey(restoredRecord),
    inverter: getInverterKey(restoredRecord),
    recordPath: restoredRecord.sourcePath,
  }
}

function uniqueSorted(values: string[]) {
  return [...new Set(values)].sort(naturalCollator.compare)
}

function getSystemGroupKey(record: MeasurementRecord) {
  return record.systemGroup ?? 'Unknown X'
}

function getInverterKey(record: MeasurementRecord) {
  return record.inverter ?? 'Unknown inverter'
}

function getReportInverterLabel(record: MeasurementRecord) {
  const systemGroup = record.systemGroup?.trim()
  const inverter = getInverterKey(record)

  return systemGroup ? `${systemGroup}.${inverter}` : inverter
}

function getStringKey(record: MeasurementRecord) {
  return record.stringName ?? record.arrayLocation ?? record.relativePath
}

interface ToleranceCalculationDetailsProps {
  rows: ToleranceRow[]
  onValueChange: (rowLabel: string, columnKey: string, value: number) => void
}

function ToleranceCalculationDetails({ rows, onValueChange }: ToleranceCalculationDetailsProps) {
  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm" aria-label="Tolerance calculation details">
      <h2 className="m-0 mb-3 text-center font-sans text-2xl font-bold uppercase leading-tight text-slate-900">
        Tolerance Calculation Details
      </h2>

      <div className="overflow-x-auto">
        <table className="w-full min-w-[920px] table-fixed border-collapse font-sans text-slate-950">
          <thead>
            <tr>
              <th className="sticky top-0 z-10 w-[82px] border border-slate-500 bg-slate-600 px-2 py-2" aria-label="Tolerance row" />
              {TOLERANCE_COLUMNS.map((column) => (
                <th
                  className="sticky top-0 z-10 break-words border border-slate-500 bg-slate-600 px-2 py-2 text-left align-top text-sm font-bold leading-tight text-white"
                  key={column.key}
                  scope="col"
                >
                  {column.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => (
              <tr className="group transition-colors" key={row.label}>
                <th
                  className="border border-slate-500 bg-slate-50 px-2 py-2 text-left align-middle text-xl font-semibold text-slate-950 transition-[background-color,filter] group-hover:brightness-[0.98]"
                  scope="row"
                >
                  {row.label}
                </th>
                {TOLERANCE_COLUMNS.map((column) => {
                  const isTotalColumn = column.key === TOLERANCE_TOTAL_KEY
                  const cellValue = getToleranceCellValue(row, column.key)
                  const isHighlightedTotal = row.label === 'Tol-' && isTotalColumn

                  return (
                    <td
                      className={`border border-slate-500 px-2 py-2 text-right align-middle text-lg text-slate-950 transition-[background-color,filter] group-hover:brightness-[0.98] ${isHighlightedTotal ? 'bg-amber-100' : 'bg-white'}`}
                      key={column.key}
                    >
                      {isTotalColumn ? (
                        <span>{formatToleranceDisplayValue(cellValue, column.digits)}%</span>
                      ) : (
                        <label className="flex items-center justify-end gap-1">
                          <span className="sr-only">{`${row.label} ${column.label}`}</span>
                          <input
                            aria-label={`${row.label} ${column.label}`}
                            className="w-16 bg-transparent text-right font-sans text-lg text-slate-950 outline-none focus:bg-sky-50 focus:ring-2 focus:ring-sky-200"
                            inputMode="decimal"
                            onChange={(event) =>
                              onValueChange(row.label, column.key, parseToleranceInputValue(event.target.value))
                            }
                            step={column.digits === 2 ? '0.01' : '0.1'}
                            type="number"
                            value={formatToleranceInputValue(cellValue)}
                          />
                          <span>%</span>
                        </label>
                      )}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

interface IvCurveReportSummaryRow {
  inverter: string
  totalStrings: number
  passCount: number
  failCount: number
  degradationRatePercent: number
}

interface IvCurveReportDetailRow {
  inverter: string
  stringName: string
  sourcePath: string
  moduleNumber: number | null
  iscPercent: number | null
  vocPercent: number | null
  pfPercent: number | null
  ffPercent: number | null
  tolPlusPercent: number | null
  tolMinusPercent: number | null
  deviationPercent: number | null
  status: 'PASS' | 'FAIL' | 'N/A'
  note: string
}

interface BuildChatAppContextArgs {
  activeView: AppView
  canMoveNext: boolean
  canMovePrevious: boolean
  importResult: DataImportResult | null
  orderedRecords: MeasurementRecord[]
  pvModules: PvModule[]
  reportDetailRows: IvCurveReportDetailRow[]
  reportSummaryRows: IvCurveReportSummaryRow[]
  selectedPvModule: PvModule | null
  selectedRecord: MeasurementRecord | null
  selectedRecordIndex: number
  selector: ModuleSelectorState
  toleranceRows: ToleranceRow[]
}

function buildChatAppContext({
  activeView,
  canMoveNext,
  canMovePrevious,
  importResult,
  orderedRecords,
  pvModules,
  reportDetailRows,
  reportSummaryRows,
  selectedPvModule,
  selectedRecord,
  selectedRecordIndex,
  selector,
  toleranceRows,
}: BuildChatAppContextArgs): ChatAppContext {
  const recordLimit = 500
  const moduleLimit = 250
  const detailRowsByKey = new Map(
    reportDetailRows.map((row) => [`${row.inverter}::${row.stringName}`, row]),
  )
  const invertersBySystem = orderedRecords.reduce<Record<string, string[]>>((groups, record) => {
    const systemGroup = getSystemGroupKey(record)
    const inverter = getInverterKey(record)
    const currentInverters = groups[systemGroup] ?? []

    if (!currentInverters.includes(inverter)) {
      groups[systemGroup] = uniqueSorted([...currentInverters, inverter])
    }

    return groups
  }, {})
  const records = orderedRecords.slice(0, recordLimit).map((record) => {
    const inverter = getReportInverterLabel(record)
    const stringName = getStringKey(record)
    const detailRow = detailRowsByKey.get(`${inverter}::${stringName}`)

    return {
      systemGroup: getSystemGroupKey(record),
      inverter: getInverterKey(record),
      stringName,
      relativePath: record.relativePath,
      status: detailRow?.status ?? 'N/A',
      pfPercent: detailRow?.pfPercent ?? null,
      deviationPercent: detailRow?.deviationPercent ?? null,
    }
  })

  return {
    activeView,
    hasImportedData: Boolean(importResult),
    folderPath: importResult?.folderPath ?? null,
    selected: {
      systemGroup: selectedRecord ? getSystemGroupKey(selectedRecord) : selector.selectedSystemGroup || null,
      inverter: selectedRecord ? getInverterKey(selectedRecord) : selector.selectedInverter || null,
      stringName: selectedRecord ? getStringKey(selectedRecord) : null,
      moduleModel: selectedPvModule?.model ?? null,
      recordIndex: selectedRecordIndex >= 0 ? selectedRecordIndex + 1 : 0,
      totalRecords: orderedRecords.length,
    },
    navigation: {
      canMoveNext,
      canMovePrevious,
    },
    systems: uniqueSorted(orderedRecords.map(getSystemGroupKey)),
    invertersBySystem,
    records,
    recordsTruncated: orderedRecords.length > recordLimit,
    modules: pvModules.slice(0, moduleLimit).map((pvModule) => ({
      id: pvModule.id,
      model: pvModule.model,
    })),
    modulesTruncated: pvModules.length > moduleLimit,
    toleranceRows,
    summaryRows: reportSummaryRows,
  }
}

function findRecordForChatCommand(records: MeasurementRecord[], command: ChatCommand) {
  const filteredRecords = records.filter((record) => {
    return (
      (!command.systemGroup || chatTextMatches(getSystemGroupKey(record), command.systemGroup)) &&
      (!command.inverter || chatTextMatches(getInverterKey(record), command.inverter))
    )
  })
  const candidates = filteredRecords.length > 0 ? filteredRecords : records
  const query = command.stringName ?? command.recordQuery

  if (!query) {
    return filteredRecords[0] ?? null
  }

  return (
    candidates.find((record) =>
      [getStringKey(record), record.relativePath, record.sourcePath].some((value) => chatTextMatches(value, query)),
    ) ??
    records.find((record) =>
      [getStringKey(record), record.relativePath, record.sourcePath].some((value) => chatTextMatches(value, query)),
    ) ??
    null
  )
}

function getChatMovedRecord(
  orderedRecords: MeasurementRecord[],
  currentRecord: MeasurementRecord | null,
  direction: 'next' | 'previous' | null,
) {
  if (!currentRecord || !direction) {
    return null
  }

  const currentIndex = orderedRecords.findIndex((record) => record.sourcePath === currentRecord.sourcePath)
  const nextIndex = direction === 'next' ? currentIndex + 1 : currentIndex - 1

  return nextIndex >= 0 && nextIndex < orderedRecords.length ? orderedRecords[nextIndex] : null
}

function findPvModuleForChatCommand(pvModules: PvModule[], command: ChatCommand) {
  if (command.moduleId) {
    const moduleById = pvModules.find((pvModule) => pvModule.id === command.moduleId)

    if (moduleById) {
      return moduleById
    }
  }

  const query = command.moduleModel ?? command.recordQuery

  if (!query) {
    return null
  }

  return pvModules.find((pvModule) => chatTextMatches(pvModule.model, query)) ?? null
}

function chatTextMatches(value: string | null | undefined, query: string | null | undefined) {
  const normalizedValue = normalizeChatSearchText(value ?? '')
  const normalizedQuery = normalizeChatSearchText(query ?? '')
  const compactValue = compactChatSearchText(normalizedValue)
  const compactQuery = compactChatSearchText(normalizedQuery)

  if (!normalizedValue || !normalizedQuery) {
    return false
  }

  return (
    normalizedValue === normalizedQuery ||
    normalizedValue.includes(normalizedQuery) ||
    normalizedQuery.includes(normalizedValue) ||
    compactValue === compactQuery ||
    compactValue.includes(compactQuery) ||
    compactQuery.includes(compactValue)
  )
}

function normalizeChatSearchText(value: string) {
  return value
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim()
}

function compactChatSearchText(value: string) {
  return value.replace(/[^a-z0-9]+/g, '')
}

function createChatMessageId() {
  return `chat-${Date.now()}-${Math.random().toString(36).slice(2)}`
}

function readChatImageAttachment(file: File): Promise<ChatImageAttachment> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onerror = () => reject(new Error(`Unable to read ${file.name}`))
    reader.onload = () => {
      resolve({
        id: createChatMessageId(),
        name: file.name,
        mimeType: file.type,
        dataUrl: String(reader.result),
        size: file.size,
      })
    }
    reader.readAsDataURL(file)
  })
}

function IvCurveReportSummaryTable({
  rows,
  selectedInverter,
}: {
  rows: IvCurveReportSummaryRow[]
  selectedInverter: string | null
}) {
  if (rows.length === 0) {
    return null
  }

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm" aria-label="IV curve report summary">
      <h2 className="m-0 mb-6 text-center text-2xl font-bold uppercase tracking-wide text-slate-900">
        PV Strings - IV Curve Analysis Report Summary
      </h2>

      <div className="overflow-x-auto">
        <table className="mx-auto w-full max-w-[1180px] min-w-[820px] border-collapse border border-slate-500 bg-white text-slate-950">
          <thead>
            <tr>
              <th className="sticky top-0 z-10 border border-slate-500 bg-slate-100 px-3 py-6 text-center text-lg font-bold uppercase">
                Inverter
              </th>
              <th className="sticky top-0 z-10 border border-slate-500 bg-slate-100 px-3 py-6 text-center text-lg font-bold uppercase">
                Total Of String
              </th>
              <th className="sticky top-0 z-10 border border-slate-500 bg-emerald-700 px-3 py-6 text-center text-lg font-bold uppercase text-white">
                Pass
              </th>
              <th className="sticky top-0 z-10 border border-slate-500 bg-red-600 px-3 py-6 text-center text-lg font-bold uppercase text-white">
                Fail
              </th>
              <th className="sticky top-0 z-10 border border-slate-500 bg-slate-100 px-3 py-6 text-center text-lg font-bold uppercase">
                Deg Rate
                <br />
                (%)
              </th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const isSelected = row.inverter === selectedInverter
              const rowClass = isSelected ? 'bg-sky-50 outline outline-2 outline-sky-300' : 'hover:bg-slate-50'

              return (
                <tr className={`transition-colors ${rowClass}`} key={row.inverter}>
                  <td className="border border-slate-500 px-2 py-2 text-center text-lg">{row.inverter}</td>
                  <td className="border border-slate-500 px-2 py-2 text-center text-lg">{row.totalStrings}</td>
                  <td className="border border-slate-500 bg-emerald-50 px-2 py-2 text-center text-lg">{row.passCount}</td>
                  <td className="border border-slate-500 bg-red-50 px-2 py-2 text-center text-lg">{row.failCount}</td>
                  <td className="border border-slate-500 px-2 py-2 text-center text-lg">
                    {formatCompactPercent(row.degradationRatePercent)}
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function IvCurveReportDetailTable({
  onNoteChange,
  rows,
  selectedRowKey,
}: {
  onNoteChange: (sourcePath: string, value: string) => void
  rows: IvCurveReportDetailRow[]
  selectedRowKey: string | null
}) {
  if (rows.length === 0) {
    return null
  }

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm" aria-label="IV curve string detail">
      <div className="max-h-[70vh] overflow-auto rounded-md border border-slate-200">
        <table className="w-full min-w-[1240px] table-fixed border-collapse border border-slate-500 text-slate-950">
          <thead>
            <tr>
              <DetailHeader className="w-[170px]" label="Inv." />
              <DetailHeader className="w-[90px]" label="String" />
              <DetailHeader className="w-[95px]" label="Module Number." />
              <DetailHeader className="w-[86px]" label="Isc (%)" />
              <DetailHeader className="w-[86px]" label="Voc (%)" />
              <DetailHeader className="w-[76px]" label="PF" />
              <DetailHeader className="w-[76px]" label="FF" />
              <DetailHeader className="w-[70px]" label="Tol+" />
              <DetailHeader className="w-[70px]" label="Tol-" />
              <DetailHeader className="w-[130px]" label="Deviation from expected" />
              <DetailHeader className="w-[120px]" label="Pass/Fail" />
              <DetailHeader className="w-[250px]" label="Note" />
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const rowKey = `${row.inverter}-${row.stringName}`
              const isSelected = rowKey === selectedRowKey
              const deviationColor =
                row.status === 'PASS' ? 'bg-emerald-100' : row.status === 'FAIL' ? 'bg-red-100 text-red-950' : 'bg-white'
              const rowClass = isSelected ? 'bg-sky-50 outline outline-2 outline-sky-300' : 'hover:bg-slate-50'

              return (
                <tr className={`group transition-colors ${rowClass}`} key={rowKey}>
                  <DetailCell align="left" value={row.inverter} />
                  <DetailCell align="left" value={row.stringName} />
                  <DetailCell align="left" value={formatNullableInteger(row.moduleNumber)} />
                  <DetailCell color="bg-emerald-100" value={formatReportPercent(row.iscPercent, 1)} />
                  <DetailCell color="bg-emerald-100" value={formatReportPercent(row.vocPercent, 1)} />
                  <DetailCell color="bg-emerald-100" value={formatReportPercent(row.pfPercent, 1)} />
                  <DetailCell color="bg-emerald-100" value={formatReportPercent(row.ffPercent, 1)} />
                  <DetailCell value={formatReportPercent(row.tolPlusPercent, 1)} />
                  <DetailCell value={formatReportPercent(row.tolMinusPercent, 1)} />
                  <DetailCell color={deviationColor} value={formatReportPercent(row.deviationPercent, 1)} />
                  <DetailStatusCell status={row.status} />
                  <DetailNoteCell value={row.note} onChange={(value) => onNoteChange(row.sourcePath, value)} />
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function DetailHeader({ className = '', label }: { className?: string; label: string }) {
  return (
    <th className={`sticky top-0 z-10 border border-slate-500 bg-slate-200 px-2 py-3 text-left text-lg font-bold ${className}`}>
      {label}
    </th>
  )
}

function DetailCell({
  align = 'right',
  color = 'bg-white',
  value,
}: {
  align?: 'left' | 'right'
  color?: string
  value: string
}) {
  const alignClass = align === 'left' ? 'text-left' : 'text-right'

  return (
    <td className={`border border-slate-500 px-2 py-2 text-lg transition-[background-color,filter] group-hover:brightness-[0.98] ${alignClass} ${color}`}>
      {value}
    </td>
  )
}

function DetailNoteCell({ onChange, value }: { onChange: (value: string) => void; value: string }) {
  return (
    <td className="border border-slate-500 bg-white p-1 transition-colors group-hover:bg-slate-50">
      <textarea
        aria-label="Edit note"
        className="min-h-16 w-full resize-y rounded-md border border-blue-200 bg-white px-2 py-1 text-sm leading-snug text-slate-950 outline-none transition focus:border-blue-500 focus:ring-2 focus:ring-blue-200"
        onChange={(event) => onChange(event.target.value)}
        rows={2}
        value={value}
      />
    </td>
  )
}

function DetailStatusCell({ status }: { status: 'PASS' | 'FAIL' | 'N/A' }) {
  const badgeClass =
    status === 'PASS'
      ? 'bg-emerald-700 text-white'
      : status === 'FAIL'
        ? 'bg-red-600 text-white'
        : 'bg-slate-200 text-slate-700'

  return (
    <td className="border border-slate-500 bg-white px-2 py-2 text-left text-lg transition-colors group-hover:bg-slate-50">
      <span className={`inline-flex min-w-20 justify-center rounded-full px-3 py-1 text-sm font-bold tracking-wide ${badgeClass}`}>
        {status}
      </span>
    </td>
  )
}

function parseToleranceInputValue(value: string) {
  const parsed = Number(value)

  return Number.isFinite(parsed) ? parsed : 0
}

function formatToleranceInputValue(value: number) {
  return Number.isFinite(value) ? value.toString() : ''
}

function formatToleranceDisplayValue(value: number, digits: number) {
  return Number.isFinite(value) ? value.toFixed(digits) : (0).toFixed(digits)
}

function formatStringOption(record: MeasurementRecord) {
  return getStringKey(record)
}

interface MeasurementReportProps {
  imageSlots: Array<RecordImageSlot | null>
  onImageSlotChange: (slotIndex: number, slot: RecordImageSlot | null) => void
  record: MeasurementRecord
  pvModule: PvModule | null
  toleranceRows: ToleranceRow[]
}

function MeasurementReport({ imageSlots, onImageSlotChange, record, pvModule, toleranceRows }: MeasurementReportProps) {
  const stcProblem = useMemo(() => getStcConversionProblem(record, pvModule), [record, pvModule])
  const stcMeasurements = useMemo(() => convertIvMeasurementsToStc(record, pvModule), [record, pvModule])
  const referenceMeasurements = useMemo(() => buildPvModuleReferenceMeasurements(pvModule), [pvModule])
  const report = useMemo(
    () => buildMeasurementReportData(record, pvModule, toleranceRows),
    [pvModule, record, toleranceRows],
  )
  const ivChart = useMemo(
    () => buildCurveChart(stcMeasurements, referenceMeasurements, 'amps'),
    [referenceMeasurements, stcMeasurements],
  )
  const powerChart = useMemo(
    () => buildCurveChart(stcMeasurements, referenceMeasurements, 'watts'),
    [referenceMeasurements, stcMeasurements],
  )

  return (
    <section className="rounded-lg border border-slate-200 bg-white p-4 shadow-sm" aria-label="Measurement report">
      <MeasurementReportTable report={report} />

      {!ivChart || !powerChart ? (
        <div className="mt-4 grid min-h-[220px] place-items-center border border-slate-200 bg-white text-sm text-slate-500">
          {stcProblem ? `No STC curve data: ${stcProblem}` : 'No STC curve data'}
        </div>
      ) : (
        <div className="mt-4 grid gap-4 xl:grid-cols-2">
          <CurveChart
            ariaLabel="IV curve"
            chart={ivChart}
            compact
            strokeClassName="stroke-sky-700"
            title="I-V Curve"
            yAxisLabel="A"
          />
          <CurveChart
            ariaLabel="PV curve"
            chart={powerChart}
            compact
            strokeClassName="stroke-amber-600"
            title="P-V Curve"
            yAxisLabel="W"
          />
        </div>
      )}

      <RecordImageSlotsPanel imageSlots={imageSlots} onImageSlotChange={onImageSlotChange} />
    </section>
  )
}

interface RecordImageSlotsPanelProps {
  imageSlots: Array<RecordImageSlot | null>
  onImageSlotChange: (slotIndex: number, slot: RecordImageSlot | null) => void
}

function RecordImageSlotsPanel({ imageSlots, onImageSlotChange }: RecordImageSlotsPanelProps) {
  const labels = ['Thermal', 'Visible']

  return (
    <section className="mt-4 grid gap-4 xl:grid-cols-2" aria-label="IV image annotations">
      {[0, 1].map((slotIndex) => (
        <RecordImageSlotEditor
          key={slotIndex}
          label={labels[slotIndex]}
          slot={imageSlots[slotIndex] ?? null}
          onChange={(slot) => onImageSlotChange(slotIndex, slot)}
        />
      ))}
    </section>
  )
}

interface RecordImageSlotEditorProps {
  label: string
  onChange: (slot: RecordImageSlot | null) => void
  slot: RecordImageSlot | null
}

function RecordImageSlotEditor({ label, onChange, slot }: RecordImageSlotEditorProps) {
  const fileInputRef = useRef<HTMLInputElement | null>(null)
  const dragStartRef = useRef<{ clientX: number; clientY: number; offsetX: number; offsetY: number } | null>(null)
  const rectangleStartRef = useRef<RecordImagePoint | null>(null)
  const didPanRef = useRef(false)
  const [draftShape, setDraftShape] = useState<RecordImageShape | null>(null)
  const [drawMode, setDrawMode] = useState<'polygon' | 'rectangle'>('polygon')
  const [isMoveEnabled, setIsMoveEnabled] = useState(false)
  const [shapeColor, setShapeColor] = useState(defaultRecordImageShapeColor)

  const handleFileChange = async (files: FileList | null) => {
    const file = files?.[0]

    if (!file) {
      return
    }

    const nextSlot = await createRecordImageSlotFromFile(file)

    onChange(nextSlot)
  }

  const handleOverlayClick = (event: MouseEvent<SVGSVGElement>) => {
    if (didPanRef.current) {
      didPanRef.current = false
      return
    }

    if (!slot || drawMode !== 'polygon' || dragStartRef.current || rectangleStartRef.current) {
      return
    }

    event.preventDefault()
    const point = getRecordImagePointFromPointer(event, slot)

    if (!point) {
      return
    }

    onChange({
      ...slot,
      polygon: [],
      shapes: addRecordImagePolygonPoint(slot.shapes, point, shapeColor),
    })
  }

  const startImagePan = (event: PointerEvent<SVGSVGElement>) => {
    if (!slot) {
      return
    }

    if (event.altKey && isMoveEnabled) {
      event.preventDefault()
      didPanRef.current = false
      dragStartRef.current = {
        clientX: event.clientX,
        clientY: event.clientY,
        offsetX: slot.offsetX,
        offsetY: slot.offsetY,
      }
      event.currentTarget.setPointerCapture?.(event.pointerId)
      return
    }

    if (drawMode === 'rectangle') {
      const point = getRecordImagePointFromPointer(event, slot)

      if (!point) {
        return
      }

      event.preventDefault()
      didPanRef.current = false
      rectangleStartRef.current = point
      setDraftShape({ color: shapeColor, points: [point, point], type: 'rectangle' })
      event.currentTarget.setPointerCapture?.(event.pointerId)
    }
  }

  const moveImagePan = (event: PointerEvent<SVGSVGElement>) => {
    const dragStart = dragStartRef.current

    if (!slot) {
      return
    }

    if (dragStart) {
      const rect = event.currentTarget.getBoundingClientRect()
      const nextOffsetX = clampNumber(dragStart.offsetX + (event.clientX - dragStart.clientX) / rect.width, -1, 1)
      const nextOffsetY = clampNumber(dragStart.offsetY + (event.clientY - dragStart.clientY) / rect.height, -1, 1)

      if (Math.abs(event.clientX - dragStart.clientX) > 2 || Math.abs(event.clientY - dragStart.clientY) > 2) {
        didPanRef.current = true
      }

      onChange({
        ...slot,
        offsetX: nextOffsetX,
        offsetY: nextOffsetY,
      })
      return
    }

    if (rectangleStartRef.current) {
      const point = getRecordImagePointFromPointer(event, slot, true)

      if (!point) {
        return
      }

      didPanRef.current = true
      setDraftShape({ color: shapeColor, points: [rectangleStartRef.current, point], type: 'rectangle' })
    }
  }

  const endImagePan = () => {
    if (slot && rectangleStartRef.current && draftShape) {
      const [startPoint, endPoint] = draftShape.points
      const hasSize = Math.abs(startPoint.x - endPoint.x) > 0.01 && Math.abs(startPoint.y - endPoint.y) > 0.01

      if (hasSize) {
        onChange({
          ...slot,
          polygon: [],
          shapes: [...slot.shapes, draftShape],
        })
      }
    }

    dragStartRef.current = null
    rectangleStartRef.current = null
    setDraftShape(null)
  }

  const handleImageWheel = (event: WheelEvent<SVGSVGElement>) => {
    if (!slot || !event.ctrlKey) {
      return
    }

    event.preventDefault()
    const nextZoom = clampNumber(
      getRecordImageZoom(slot.zoom) * (event.deltaY < 0 ? 1.1 : 0.9),
      recordImageMinZoom,
      recordImageMaxZoom,
    )

    onChange({ ...slot, zoom: Number(nextZoom.toFixed(3)) })
  }

  return (
    <section className="rounded-md border border-slate-200 bg-white p-3 shadow-sm">
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
        <div>
          <h3 className="m-0 text-sm font-semibold text-slate-950">{label}</h3>
          <p className="m-0 text-xs text-slate-500">{slot?.name ?? 'No image selected'}</p>
        </div>
        <div className="flex flex-wrap gap-1">
          <button
            className={compactToolButtonClass(false)}
            type="button"
            title="Upload image"
            aria-label="Upload image"
            onClick={() => fileInputRef.current?.click()}
          >
            <ToolIcon name="upload" />
          </button>
          <button
            className={compactToolButtonClass(false)}
            type="button"
            title="Clear image"
            aria-label="Clear image"
            disabled={!slot}
            onClick={() => onChange(null)}
          >
            <ToolIcon name="clear" />
          </button>
        </div>
      </div>

      <input
        ref={fileInputRef}
        className="hidden"
        type="file"
        accept="image/png,image/jpeg,image/webp"
        onChange={(event) => {
          void handleFileChange(event.currentTarget.files)
          event.currentTarget.value = ''
        }}
      />

      <div className="relative grid aspect-[4/3] place-items-center overflow-hidden rounded-md border border-dashed border-slate-300 bg-slate-50">
        {slot ? (
          <>
            <img
              alt=""
              className="absolute max-w-none select-none"
              draggable={false}
              src={slot.dataUrl}
              style={{
                height: `${getRecordImageZoom(slot.zoom) * 100}%`,
                left: `${50 + slot.offsetX * 100}%`,
                objectFit: 'contain',
                top: `${50 + slot.offsetY * 100}%`,
                transform: 'translate(-50%, -50%)',
                width: `${getRecordImageZoom(slot.zoom) * 100}%`,
              }}
            />
            <svg
              className={`absolute inset-0 h-full w-full ${isMoveEnabled ? 'cursor-move' : 'cursor-crosshair'}`}
              role="img"
              aria-label={`${label} polygon editor`}
              viewBox="0 0 1 1"
              preserveAspectRatio="none"
              onClick={handleOverlayClick}
              onPointerDown={startImagePan}
              onPointerMove={moveImagePan}
              onPointerUp={endImagePan}
              onPointerCancel={endImagePan}
              onWheel={handleImageWheel}
            >
              <g transform={getRecordImageSvgTransform(slot)}>
                {[...slot.shapes, ...(draftShape ? [draftShape] : [])].map((shape, index) => (
                  <RecordImageShapeOverlay
                    key={`${shape.type}-${index}`}
                    shape={shape}
                    strokeWidth={getRecordImageSvgStrokeWidth(slot.strokeWidth)}
                  />
                ))}
              </g>
            </svg>
          </>
        ) : (
          <button className="h-10 rounded-md border border-blue-200 bg-white px-4 text-sm font-semibold text-blue-950 shadow-sm transition hover:border-blue-400 hover:bg-blue-50" type="button" onClick={() => fileInputRef.current?.click()}>
            Add image
          </button>
        )}
      </div>

      <div className="mt-2 flex flex-wrap items-center justify-between gap-2">
        <div className="flex flex-wrap gap-1">
          <button
            className={compactToolButtonClass(isMoveEnabled)}
            type="button"
            title="Move image: hold Alt and drag"
            aria-label="Move image"
            disabled={!slot}
            onClick={() => setIsMoveEnabled((value) => !value)}
          >
            <ToolIcon name="move" />
          </button>
          <button
            className={compactToolButtonClass(drawMode === 'polygon')}
            type="button"
            title="Draw polygon"
            aria-label="Draw polygon"
            disabled={!slot}
            onClick={() => setDrawMode('polygon')}
          >
            <ToolIcon name="polygon" />
          </button>
          <button
            className={compactToolButtonClass(drawMode === 'rectangle')}
            type="button"
            title="Draw rectangle"
            aria-label="Draw rectangle"
            disabled={!slot}
            onClick={() => setDrawMode('rectangle')}
          >
            <ToolIcon name="rectangle" />
          </button>
          <label
            className="grid h-7 w-9 place-items-center rounded border border-slate-200 bg-white shadow-sm"
            title="Shape color"
            aria-label="Shape color"
          >
            <input
              className="h-5 w-7 cursor-pointer border-0 bg-transparent p-0"
              type="color"
              value={shapeColor}
              onChange={(event) => setShapeColor(normalizeRecordImageColor(event.currentTarget.value))}
            />
          </label>
          <label
            className="flex h-7 items-center gap-1 rounded border border-slate-200 bg-white px-1.5 text-slate-700 shadow-sm"
            title={slot ? `Outline width: ${normalizeRecordImageStrokeWidth(slot.strokeWidth)}` : 'Outline width'}
            aria-label="Outline width"
          >
            <ToolIcon name="width" />
            <input
              className="h-1.5 w-20 accent-blue-600 disabled:cursor-not-allowed disabled:opacity-40"
              type="range"
              min={recordImageMinStrokeWidth}
              max={recordImageMaxStrokeWidth}
              step={0.5}
              disabled={!slot}
              value={slot ? normalizeRecordImageStrokeWidth(slot.strokeWidth) : 2}
              onChange={(event) =>
                slot &&
                onChange({
                  ...slot,
                  strokeWidth: normalizeRecordImageStrokeWidth(Number(event.currentTarget.value)),
                })
              }
            />
          </label>
        </div>
        <div className="flex flex-wrap gap-2">
          <button
            className={compactToolButtonClass(false)}
            type="button"
            title="Undo point"
            aria-label="Undo point"
            disabled={!slot || slot.shapes.length === 0}
            onClick={() => slot && onChange({ ...slot, polygon: [], shapes: undoRecordImageShape(slot.shapes) })}
          >
            <ToolIcon name="undo" />
          </button>
          <button
            className={compactToolButtonClass(false)}
            type="button"
            title="Clear shapes"
            aria-label="Clear shapes"
            disabled={!slot || slot.shapes.length === 0}
            onClick={() => slot && onChange({ ...slot, polygon: [], shapes: [] })}
          >
            <ToolIcon name="trash" />
          </button>
        </div>
      </div>
    </section>
  )
}

function RecordImageShapeOverlay({
  shape,
  strokeWidth,
}: {
  shape: RecordImageShape
  strokeWidth: number
}) {
  const color = normalizeRecordImageColor(shape.color)

  if (shape.type === 'rectangle' && shape.points.length >= 2) {
    const rectangle = getRecordImageRectangle(shape.points)

    return (
      <rect
        fill="none"
        height={rectangle.height}
        stroke={color}
        strokeWidth={strokeWidth}
        width={rectangle.width}
        x={rectangle.x}
        y={rectangle.y}
      />
    )
  }

  if (shape.type !== 'polygon' || shape.points.length === 0) {
    return null
  }

  const points = shape.points.map((point) => `${point.x},${point.y}`).join(' ')

  return (
    <g>
      {shape.points.length >= 3 && <polygon fill="none" points={points} stroke={color} strokeWidth={strokeWidth} />}
      {shape.points.length >= 2 && (
        <polyline fill="none" points={points} stroke={color} strokeDasharray="0.015 0.01" strokeWidth={strokeWidth} />
      )}
    </g>
  )
}

async function createRecordImageSlotFromFile(file: File): Promise<RecordImageSlot> {
  const dataUrl = await resizeImageFileToDataUrl(file)

  return {
    dataUrl,
    fit: 'contain',
    name: file.name,
    offsetX: 0,
    offsetY: 0,
    polygon: [],
    shapes: [],
    strokeWidth: 2,
    zoom: 1,
  }
}

async function resizeImageFileToDataUrl(file: File) {
  const dataUrl = await readFileAsDataUrl(file)
  const image = await loadImageElement(dataUrl)
  const maxWidth = 1600
  const maxHeight = 1000
  const scale = Math.min(1, maxWidth / image.naturalWidth, maxHeight / image.naturalHeight)
  const width = Math.max(1, Math.round(image.naturalWidth * scale))
  const height = Math.max(1, Math.round(image.naturalHeight * scale))

  if (scale >= 1 && dataUrl.length < 1_500_000) {
    return dataUrl
  }

  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')

  canvas.width = width
  canvas.height = height
  context?.drawImage(image, 0, 0, width, height)

  return canvas.toDataURL('image/jpeg', 0.88)
}

function readFileAsDataUrl(file: File): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onload = () => {
      if (typeof reader.result === 'string') {
        resolve(reader.result)
      } else {
        reject(new Error('Could not read image file.'))
      }
    }
    reader.onerror = () => reject(reader.error ?? new Error('Could not read image file.'))
    reader.readAsDataURL(file)
  })
}

function loadImageElement(dataUrl: string): Promise<HTMLImageElement> {
  return new Promise((resolve, reject) => {
    const image = new Image()

    image.onload = () => resolve(image)
    image.onerror = () => reject(new Error('Could not load image file.'))
    image.src = dataUrl
  })
}

function getRecordImageZoom(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value) ? clampNumber(value, recordImageMinZoom, recordImageMaxZoom) : 1
}

function normalizeRecordImageStrokeWidth(value: unknown) {
  return typeof value === 'number' && Number.isFinite(value)
    ? clampNumber(value, recordImageMinStrokeWidth, recordImageMaxStrokeWidth)
    : 2
}

function getRecordImageSvgStrokeWidth(value: unknown) {
  return normalizeRecordImageStrokeWidth(value) / 700
}

function getRecordImageFrame(slot: RecordImageSlot) {
  const zoom = getRecordImageZoom(slot.zoom)

  return {
    size: zoom,
    x: 0.5 + slot.offsetX - zoom / 2,
    y: 0.5 + slot.offsetY - zoom / 2,
  }
}

function getRecordImageSvgTransform(slot: RecordImageSlot) {
  const frame = getRecordImageFrame(slot)

  return `translate(${frame.x} ${frame.y}) scale(${frame.size})`
}

function getRecordImagePointFromPointer(event: MouseEvent<SVGSVGElement> | PointerEvent<SVGSVGElement>, slot: RecordImageSlot, allowClamp = false) {
  const rect = event.currentTarget.getBoundingClientRect()
  const frame = getRecordImageFrame(slot)
  const overlayX = (event.clientX - rect.left) / rect.width
  const overlayY = (event.clientY - rect.top) / rect.height
  const imageX = (overlayX - frame.x) / frame.size
  const imageY = (overlayY - frame.y) / frame.size

  if (!allowClamp && (imageX < 0 || imageX > 1 || imageY < 0 || imageY > 1)) {
    return null
  }

  return {
    x: clampNumber(imageX, 0, 1),
    y: clampNumber(imageY, 0, 1),
  }
}

function addRecordImagePolygonPoint(shapes: RecordImageShape[], point: RecordImagePoint, color: string): RecordImageShape[] {
  const safeColor = normalizeRecordImageColor(color)
  const nextShapes = shapes.slice()
  const previousShape = nextShapes[nextShapes.length - 1]

  if (previousShape?.type === 'polygon' && normalizeRecordImageColor(previousShape.color) === safeColor) {
    nextShapes[nextShapes.length - 1] = {
      ...previousShape,
      color: safeColor,
      points: [...previousShape.points, point],
    }

    return nextShapes
  }

  return [...nextShapes, { color: safeColor, points: [point], type: 'polygon' as const }]
}

function undoRecordImageShape(shapes: RecordImageShape[]): RecordImageShape[] {
  const nextShapes = shapes.slice()
  const previousShape = nextShapes[nextShapes.length - 1]

  if (!previousShape) {
    return nextShapes
  }

  if (previousShape.type === 'polygon' && previousShape.points.length > 1) {
    nextShapes[nextShapes.length - 1] = {
      ...previousShape,
      points: previousShape.points.slice(0, -1),
    }

    return nextShapes
  }

  return nextShapes.slice(0, -1)
}

function getRecordImageRectangle(points: RecordImagePoint[]) {
  const [startPoint, endPoint] = points
  const x = Math.min(startPoint.x, endPoint.x)
  const y = Math.min(startPoint.y, endPoint.y)

  return {
    height: Math.abs(startPoint.y - endPoint.y),
    width: Math.abs(startPoint.x - endPoint.x),
    x,
    y,
  }
}

function normalizeRecordImageColor(value: unknown) {
  return typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value) ? value : defaultRecordImageShapeColor
}

function clampNumber(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

interface MeasurementReportData {
  date: string | null
  deviationPercent: number | null
  ffPercent: number | null
  impMeasuredA: number | null
  impNominalA: number | null
  impTranslatedA: number | null
  inverter: string | null
  iscMeasuredA: number | null
  iscNominalA: number | null
  iscPercent: number | null
  iscTranslatedA: number | null
  irradianceWm2: number | null
  model: string | null
  moduleCount: number | null
  pfPercent: number | null
  pmaxMeasuredW: number | null
  pmaxNominalW: number | null
  pmaxTranslatedW: number | null
  status: 'PASS' | 'FAIL' | 'N/A'
  stringName: string | null
  temperatureC: number | null
  time: string | null
  tolMinusPercent: number | null
  tolPlusPercent: number | null
  vmpMeasuredV: number | null
  vmpNominalV: number | null
  vmpTranslatedV: number | null
  vocMeasuredV: number | null
  vocNominalV: number | null
  vocPercent: number | null
  vocTranslatedV: number | null
}

function MeasurementReportTable({ report }: { report: MeasurementReportData }) {
  const statusClass =
    report.status === 'FAIL'
      ? 'bg-red-600 text-white'
      : report.status === 'N/A'
        ? 'bg-slate-300 text-slate-900'
        : 'bg-emerald-700 text-white'

  return (
    <div className="overflow-x-auto rounded-lg border border-slate-200 bg-white p-1">
      <table className="w-full min-w-[1360px] table-fixed border-collapse border border-slate-500 text-center text-slate-950">
        <tbody>
          <tr>
            <th className="border border-slate-500 bg-slate-50 px-3 py-4 text-left text-4xl font-bold" colSpan={2} rowSpan={2}>
              System
            </th>
            <th className="border border-slate-500 bg-slate-50 px-3 py-2 text-xl font-bold" colSpan={4}>
              Model: {report.model ?? '-'}
            </th>
            <th className="border border-slate-500 bg-slate-50 px-3 py-2 text-xl font-bold" colSpan={2}>
              Date
            </th>
            <th className="border border-slate-500 bg-slate-50 px-3 py-2 text-xl font-bold" colSpan={2}>
              Time
            </th>
            <th className={`border border-slate-500 px-3 py-2 text-5xl font-semibold ${statusClass}`} colSpan={3} rowSpan={2}>
              {report.status}
            </th>
          </tr>
          <tr>
            <td className="border border-slate-500 bg-white px-3 py-3 text-lg font-semibold">Inv.</td>
            <td className="border border-slate-500 bg-white px-3 py-3 text-lg">{report.inverter ?? '-'}</td>
            <td className="border border-slate-500 bg-white px-3 py-3 text-lg font-semibold">String.</td>
            <td className="border border-slate-500 bg-white px-3 py-3 text-lg">{report.stringName ?? '-'}</td>
            <td className="border border-slate-500 bg-white px-3 py-3" colSpan={2}>{report.date ?? '-'}</td>
            <td className="border border-slate-500 bg-white px-3 py-3" colSpan={2}>{report.time ?? '-'}</td>
          </tr>

          <tr>
            <th className="border border-slate-500 bg-white px-3 py-3 text-lg font-semibold" rowSpan={2}>
              Number<br />of<br />modules
            </th>
            <th className="border border-slate-500 bg-slate-100 px-3 py-2 text-lg font-bold">Irr.</th>
            <th className="border border-slate-500 bg-amber-100 px-3 py-2 text-lg font-bold">Temp.</th>
            <th className="border border-slate-500 bg-yellow-50 px-3 py-2 text-lg font-bold" colSpan={3}>Imp</th>
            <th className="border border-slate-500 bg-emerald-50 px-3 py-2 text-lg font-bold" colSpan={3}>Vmp</th>
            <th className="border border-slate-500 bg-slate-100 px-3 py-2 text-lg font-bold" colSpan={4}>Isc</th>
          </tr>
          <tr>
            <td className="border border-slate-500 bg-slate-100 px-3 py-2">(W/m^2)</td>
            <td className="border border-slate-500 bg-amber-100 px-3 py-2">(°C)</td>
            <ReportSubHeader color="bg-yellow-50" label="Measured" unit="(Amps)" />
            <ReportSubHeader color="bg-yellow-50" label="Translated" unit="to STC" />
            <ReportSubHeader color="bg-yellow-50" label="Nominal" unit="at STC" />
            <ReportSubHeader color="bg-emerald-50" label="Measured" unit="(Volts)" />
            <ReportSubHeader color="bg-emerald-50" label="Translated" unit="to STC" />
            <ReportSubHeader color="bg-emerald-50" label="Nominal" unit="at STC" />
            <ReportSubHeader color="bg-slate-100" label="Measured" unit="(Amps)" />
            <ReportSubHeader color="bg-slate-100" label="Translated" unit="to STC" />
            <ReportSubHeader color="bg-slate-100" label="Nominal" unit="at STC" />
            <td className="border border-slate-500 bg-slate-100 px-3 py-2">%</td>
          </tr>
          <tr>
            <ReportValue color="bg-white" value={formatReportNumber(report.moduleCount, 0)} />
            <ReportValue color="bg-slate-100" value={formatReportNumber(report.irradianceWm2, 0)} />
            <ReportValue color="bg-amber-100" value={formatReportNumber(report.temperatureC, 1)} />
            <ReportValue color="bg-yellow-50" value={formatReportNumber(report.impMeasuredA, 2)} />
            <ReportValue color="bg-yellow-50" value={formatReportNumber(report.impTranslatedA, 2)} />
            <ReportValue color="bg-yellow-50" value={formatReportNumber(report.impNominalA, 2)} />
            <ReportValue color="bg-emerald-50" value={formatReportNumber(report.vmpMeasuredV, 0)} />
            <ReportValue color="bg-emerald-50" value={formatReportNumber(report.vmpTranslatedV, 0)} />
            <ReportValue color="bg-emerald-50" value={formatReportNumber(report.vmpNominalV, 0)} />
            <ReportValue color="bg-slate-100" value={formatReportNumber(report.iscMeasuredA, 2)} />
            <ReportValue color="bg-slate-100" value={formatReportNumber(report.iscTranslatedA, 2)} />
            <ReportValue color="bg-slate-100" value={formatReportNumber(report.iscNominalA, 2)} />
            <ReportValue color="bg-slate-100" value={formatReportPercent(report.iscPercent, 1)} />
          </tr>

          <tr>
            <th className="border border-slate-500 bg-orange-50 px-3 py-2 text-lg font-bold" colSpan={4}>Voc</th>
            <th className="border border-slate-500 bg-red-50 px-3 py-2 text-lg font-bold" colSpan={3}>Pmax</th>
            <th className="border border-slate-500 bg-sky-100 px-3 py-2 font-bold" rowSpan={2}>PF</th>
            <th className="border border-slate-500 bg-sky-100 px-3 py-2 font-bold" rowSpan={2}>FF</th>
            <th className="border border-slate-500 bg-sky-100 px-3 py-2 font-bold">Tol +<br /><span className="font-normal">( % )</span></th>
            <th className="border border-slate-500 bg-sky-100 px-3 py-2 font-bold" colSpan={2}>Tol -<br /><span className="font-normal">( % )</span></th>
            <th className="border border-slate-500 bg-sky-100 px-3 py-2 font-bold">Deviation<br />from<br />expected</th>
          </tr>
          <tr>
            <ReportSubHeader color="bg-orange-50" label="Measured" unit="(Volts)" />
            <ReportSubHeader color="bg-orange-50" label="Translated" unit="to STC" />
            <ReportSubHeader color="bg-orange-50" label="Nominal" unit="at STC" />
            <ReportSubHeader color="bg-orange-50" label="%" unit="" />
            <ReportSubHeader color="bg-red-50" label="Measured" unit="(Watts)" />
            <ReportSubHeader color="bg-red-50" label="Translated" unit="to STC" />
            <ReportSubHeader color="bg-red-50" label="Nominal" unit="at STC" />
            <td className="border border-slate-500 bg-sky-100 px-3 py-2">Equipment</td>
            <td className="border border-slate-500 bg-sky-100 px-3 py-2" colSpan={2}>Equipment + Degradation</td>
            <td className="border border-slate-500 bg-sky-100 px-3 py-2">(%)</td>
          </tr>
          <tr>
            <ReportValue color="bg-orange-50" value={formatReportNumber(report.vocMeasuredV, 0)} />
            <ReportValue color="bg-orange-50" value={formatReportNumber(report.vocTranslatedV, 0)} />
            <ReportValue color="bg-orange-50" value={formatReportNumber(report.vocNominalV, 0)} />
            <ReportValue color="bg-orange-50" value={formatReportPercent(report.vocPercent, 1)} />
            <ReportValue color="bg-red-50" value={formatReportNumber(report.pmaxMeasuredW, 0)} />
            <ReportValue color="bg-red-50" value={formatReportNumber(report.pmaxTranslatedW, 0)} />
            <ReportValue color="bg-red-50" value={formatReportNumber(report.pmaxNominalW, 0)} />
            <ReportValue color="bg-sky-100" value={formatReportPercent(report.pfPercent, 1)} />
            <ReportValue color="bg-sky-100" value={formatReportPercent(report.ffPercent, 1)} />
            <ReportValue color="bg-sky-100" value={formatReportPercent(report.tolPlusPercent, 1)} />
            <ReportValue color="bg-sky-100" colSpan={2} value={formatReportPercent(report.tolMinusPercent, 1)} />
            <ReportValue color="bg-sky-100" value={formatReportPercent(report.deviationPercent, 1)} />
          </tr>
        </tbody>
      </table>
    </div>
  )
}

function ReportSubHeader({ color, label, unit }: { color: string; label: string; unit: string }) {
  return (
    <td className={`border border-slate-500 px-3 py-2 ${color}`}>
      {label}
      {unit && (
        <>
          <br />
          <span>{unit}</span>
        </>
      )}
    </td>
  )
}

function ReportValue({ color, colSpan, value }: { color: string; colSpan?: number; value: string }) {
  return (
    <td className={`border border-slate-500 px-3 py-3 text-right text-lg ${color}`} colSpan={colSpan}>
      {value}
    </td>
  )
}

function buildMeasurementReportData(
  record: MeasurementRecord,
  pvModule: PvModule | null,
  toleranceRows: ToleranceRow[],
): MeasurementReportData {
  const summary = record.measurementSummary
  const stcSummary = convertMeasurementSummaryToStc(record, pvModule)
  const moduleCount = record.modulesInString
  const tolPlusPercent = getToleranceTotal(toleranceRows, 'Tol+')
  const tolMinusPercent = getToleranceTotal(toleranceRows, 'Tol-')
  const impTranslatedA = stcSummary?.imppA ?? null
  const vmpTranslatedV = multiplyNullable(stcSummary?.vmppV ?? null, moduleCount)
  const iscTranslatedA = stcSummary?.iscA ?? null
  const vocTranslatedV = multiplyNullable(stcSummary?.vocV ?? null, moduleCount)
  const pmaxTranslatedW = multiplyNullable(stcSummary?.pmaxW ?? null, moduleCount)
  const impNominalA = pvModule?.maximumPowerCurrentA ?? null
  const vmpNominalV = multiplyNullable(pvModule?.maximumPowerVoltageV ?? null, moduleCount)
  const iscNominalA = pvModule?.shortCircuitCurrentA ?? null
  const vocNominalV = multiplyNullable(pvModule?.openCircuitVoltageV ?? null, moduleCount)
  const pmaxNominalW = multiplyNullable(pvModule?.ratedMaximumPowerW ?? null, moduleCount)
  const pfPercent = ratioPercent(pmaxTranslatedW, pmaxNominalW)
  const ffPercent = ratioPercent(summary.pmaxW, multiplyNullable(summary.vocV, summary.iscA))
  const deviationPercent = pfPercent === null ? null : Math.abs(100 - pfPercent)

  return {
    date: record.measurementDate,
    deviationPercent,
    ffPercent,
    impMeasuredA: summary.imppA,
    impNominalA,
    impTranslatedA,
    inverter: record.inverter,
    iscMeasuredA: summary.iscA,
    iscNominalA,
    iscPercent: ratioPercent(iscTranslatedA, iscNominalA),
    iscTranslatedA,
    irradianceWm2: record.irradianceWm2,
    model: pvModule?.model ?? record.moduleModel,
    moduleCount,
    pfPercent,
    pmaxMeasuredW: summary.pmaxW,
    pmaxNominalW,
    pmaxTranslatedW,
    status: getReportStatus(pfPercent, tolPlusPercent, tolMinusPercent),
    stringName: record.stringName ?? record.arrayLocation,
    temperatureC: record.cellTemperatureC,
    time: record.measurementTime,
    tolMinusPercent,
    tolPlusPercent,
    vmpMeasuredV: summary.vmppV,
    vmpNominalV,
    vmpTranslatedV,
    vocMeasuredV: summary.vocV,
    vocNominalV,
    vocPercent: ratioPercent(vocTranslatedV, vocNominalV),
    vocTranslatedV,
  }
}

function buildIvCurveReportSummaryRows(
  records: MeasurementRecord[],
  pvModules: PvModule[],
  selectedPvModuleId: string,
  toleranceRows: ToleranceRow[],
): IvCurveReportSummaryRow[] {
  const selectedModule = pvModules.find((pvModule) => pvModule.id === selectedPvModuleId) ?? null
  const summaryRows = new Map<string, IvCurveReportSummaryRow>()

  records.forEach((record) => {
    const inverter = getReportInverterLabel(record)
    const currentRow =
      summaryRows.get(inverter) ??
      ({
        degradationRatePercent: 0,
        failCount: 0,
        inverter,
        passCount: 0,
        totalStrings: 0,
      } satisfies IvCurveReportSummaryRow)
    const pvModule = selectedModule ?? findPvModuleForRecord(record, pvModules)
    const report = buildMeasurementReportData(record, pvModule, toleranceRows)

    currentRow.totalStrings += 1

    if (report.status === 'PASS') {
      currentRow.passCount += 1
    } else {
      currentRow.failCount += 1
    }

    summaryRows.set(inverter, currentRow)
  })

  return [...summaryRows.values()]
    .map((row) => ({
      ...row,
      degradationRatePercent: row.totalStrings > 0 ? (row.failCount / row.totalStrings) * 100 : 0,
    }))
    .sort((left, right) => naturalCollator.compare(left.inverter, right.inverter))
}

function buildIvCurveReportDetailRows(
  records: MeasurementRecord[],
  pvModules: PvModule[],
  selectedPvModuleId: string,
  toleranceRows: ToleranceRow[],
): IvCurveReportDetailRow[] {
  const selectedModule = pvModules.find((pvModule) => pvModule.id === selectedPvModuleId) ?? null

  return sortRecordsByHierarchy(records)
    .map((record) => {
      const pvModule = selectedModule ?? findPvModuleForRecord(record, pvModules)
      const report = buildMeasurementReportData(record, pvModule, toleranceRows)

      return {
        deviationPercent: report.deviationPercent,
        ffPercent: report.ffPercent,
        inverter: getReportInverterLabel(record),
        iscPercent: report.iscPercent,
        moduleNumber: record.modulesInString,
        note: blankNoneText(record.measurementNote),
        pfPercent: report.pfPercent,
        sourcePath: record.sourcePath,
        status: report.status,
        stringName: getStringKey(record),
        tolMinusPercent: report.tolMinusPercent,
        tolPlusPercent: report.tolPlusPercent,
        vocPercent: report.vocPercent,
      }
    })
}

function getReportStatus(
  pfPercent: number | null,
  tolPlusPercent: number | null,
  tolMinusPercent: number | null,
): 'PASS' | 'FAIL' | 'N/A' {
  if (pfPercent === null || tolPlusPercent === null || tolMinusPercent === null) {
    return 'N/A'
  }

  return pfPercent <= 100 + tolPlusPercent && pfPercent >= 100 - tolMinusPercent ? 'PASS' : 'FAIL'
}

function getToleranceTotal(rows: ToleranceRow[], label: string) {
  const row = rows.find((item) => item.label === label)

  return row ? getToleranceCellValue(row, TOLERANCE_TOTAL_KEY) : null
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

function formatReportNumber(value: number | null, digits: number) {
  if (value === null || !Number.isFinite(value)) {
    return '-'
  }

  return value.toLocaleString(undefined, {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  })
}

function formatReportPercent(value: number | null, digits: number) {
  return value === null || !Number.isFinite(value) ? '-' : `${formatReportNumber(value, digits)}%`
}

function formatNullableInteger(value: number | null) {
  return value === null || !Number.isFinite(value) ? '-' : value.toFixed(0)
}

function blankNoneText(value: string | null | undefined) {
  const text = value?.trim() ?? ''
  const normalizedText = text.toLowerCase()

  return normalizedText === 'none' || normalizedText === '(none)' ? '' : text
}

function formatCompactPercent(value: number) {
  if (!Number.isFinite(value)) {
    return '-'
  }

  return `${Number.isInteger(value) ? value.toFixed(0) : value.toFixed(1)}%`
}

function formatDuration(elapsedMs: number) {
  if (!Number.isFinite(elapsedMs) || elapsedMs <= 0) {
    return '0s'
  }

  const totalSeconds = Math.round(elapsedMs / 1000)
  const minutes = Math.floor(totalSeconds / 60)
  const seconds = totalSeconds % 60

  return minutes > 0 ? `${minutes}m ${seconds}s` : `${seconds}s`
}

function getExportErrorMessage(error: unknown) {
  if (error instanceof Error) {
    return error.message
  }

  return typeof error === 'string' ? error : 'Unknown export error'
}

interface CurveChartProps {
  ariaLabel: string
  chart: CurveChartData
  compact?: boolean
  strokeClassName: string
  title: string
  yAxisLabel: string
}

function CurveChart({
  ariaLabel,
  chart,
  compact = false,
  strokeClassName,
  title,
  yAxisLabel,
}: CurveChartProps) {
  const measuredLegendClass = title.startsWith('I') ? 'bg-sky-700' : 'bg-amber-600'

  return (
    <div className="min-w-0 rounded-lg border border-slate-200 bg-slate-50 p-4">
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2">
        <h3 className="m-0 text-sm font-bold text-slate-800">{title}</h3>
        <div className="flex flex-wrap items-center gap-3">
          <span className="flex items-center gap-1.5 text-xs font-medium text-slate-600">
            <span className={`h-0.5 w-7 rounded-full ${measuredLegendClass}`} />
            Measured
          </span>
          {chart.referencePoints && (
            <span className="flex items-center gap-1.5 text-xs font-medium text-slate-600">
              <span className="h-0 w-7 border-t-2 border-dashed border-green-800" />
              Reference
            </span>
          )}
        </div>
      </div>
      <svg
        className={`${compact ? 'min-h-[250px]' : 'min-h-[320px]'} w-full rounded-md border border-slate-200 bg-white`}
        viewBox={`0 0 ${chart.width} ${chart.height}`}
        role="img"
        aria-label={ariaLabel}
      >
        <rect className="fill-slate-50" x={0} y={0} width={chart.width} height={chart.height} />
        <rect
          className="fill-white stroke-slate-200 stroke-[1]"
          x={chart.left}
          y={chart.top}
          width={chart.right - chart.left}
          height={chart.bottom - chart.top}
        />

        {chart.xTicks.map((tick) => (
          <g key={`x-${tick.value}`}>
            <line
              className="stroke-slate-200 stroke-[1]"
              x1={tick.x}
              y1={chart.top}
              x2={tick.x}
              y2={chart.bottom}
            />
            <text
              className="fill-slate-600 text-[11px] font-medium"
              x={tick.x}
              y={chart.bottom + 22}
              textAnchor="middle"
            >
              {tick.label}
            </text>
          </g>
        ))}

        {chart.yTicks.map((tick) => (
          <g key={`y-${tick.value}`}>
            <line
              className="stroke-slate-200 stroke-[1]"
              x1={chart.left}
              y1={tick.y}
              x2={chart.right}
              y2={tick.y}
            />
            <text
              className="fill-slate-600 text-[11px] font-medium"
              x={chart.left - 10}
              y={tick.y + 4}
              textAnchor="end"
            >
              {tick.label}
            </text>
          </g>
        ))}

        <line className="stroke-slate-500 stroke-[1.4]" x1={chart.left} y1={chart.bottom} x2={chart.right} y2={chart.bottom} />
        <line className="stroke-slate-500 stroke-[1.4]" x1={chart.left} y1={chart.top} x2={chart.left} y2={chart.bottom} />
        {chart.referencePoints && (
          <polyline
            className="fill-none stroke-green-800 stroke-[2.4] [stroke-linecap:round] [stroke-linejoin:round]"
            points={chart.referencePoints}
            strokeDasharray="8 6"
          />
        )}
        <polyline className={`fill-none ${strokeClassName} stroke-[3] [stroke-linecap:round] [stroke-linejoin:round]`} points={chart.points} />

        <text className="fill-slate-600 text-base font-bold" x={chart.right} y={chart.height - 8}>
          V
        </text>
        <text className="fill-slate-600 text-base font-bold" x={8} y={chart.top + 8}>
          {yAxisLabel}
        </text>
      </svg>
    </div>
  )
}

interface CurveChartData {
  width: number
  height: number
  left: number
  right: number
  top: number
  bottom: number
  points: string
  referencePoints: string | null
  xTicks: Array<{
    value: number
    x: number
    label: string
  }>
  yTicks: Array<{
    value: number
    y: number
    label: string
  }>
}

function buildCurveChart(
  measurements: IvPoint[],
  referenceMeasurements: IvPoint[],
  yField: 'amps' | 'watts',
): CurveChartData | null {
  if (measurements.length === 0) {
    return null
  }

  const width = 720
  const height = 360
  const left = 52
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
  const valuePoints = measurements.map((point) => {
      const x = left + (point.volts / maxVolts) * (right - left)
      const y = bottom - (point[yField] / maxY) * (bottom - top)

      return { x, y }
    })
  const points = valuePoints.map((point) => `${point.x.toFixed(2)},${point.y.toFixed(2)}`).join(' ')
  const referencePoints =
    referenceMeasurements.length > 0
      ? referenceMeasurements
          .map((point) => {
            const x = left + (point.volts / maxVolts) * (right - left)
            const y = bottom - (point[yField] / maxY) * (bottom - top)

            return `${x.toFixed(2)},${y.toFixed(2)}`
          })
          .join(' ')
      : null
  const xTicks = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = (maxVolts / tickCount) * index
    const x = left + (value / maxVolts) * (right - left)

    return { value, x, label: formatChartTick(value) }
  })
  const yTicks = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = (maxY / tickCount) * index
    const y = bottom - (value / maxY) * (bottom - top)

    return { value, y, label: formatChartTick(value) }
  })
  return { width, height, left, right, top, bottom, points, referencePoints, xTicks, yTicks }
}

function formatChartTick(value: number) {
  if (Math.abs(value) >= 100) {
    return value.toFixed(0)
  }

  if (Math.abs(value) >= 10) {
    return value.toFixed(1)
  }

  return value.toFixed(2)
}

function getCachedProjectModule(folderPath: string, modules: PvModule[]) {
  const entry = readProjectModuleCache()[folderPath]

  if (!entry) {
    return null
  }

  return (
    modules.find((pvModule) => pvModule.id === entry.moduleId) ??
    modules.find((pvModule) => normalizeModuleModelForCache(pvModule.model) === normalizeModuleModelForCache(entry.moduleModel)) ??
    null
  )
}

function cacheProjectModule(folderPath: string, pvModule: PvModule) {
  const cache = readProjectModuleCache()

  cache[folderPath] = {
    moduleId: pvModule.id,
    moduleModel: pvModule.model,
  }

  writeProjectModuleCache(cache)
}

function clearCachedProjectModule(folderPath: string) {
  const cache = readProjectModuleCache()

  delete cache[folderPath]
  writeProjectModuleCache(cache)
}

function applyRecordNoteOverrides(importResult: DataImportResult, overrides: RecordNoteOverrides): DataImportResult {
  return {
    ...importResult,
    records: importResult.records.map((record) => {
      const noteKey = getRecordNoteOverrideKey(importResult.folderPath, record)

      if (!Object.prototype.hasOwnProperty.call(overrides, noteKey)) {
        return record
      }

      const note = overrides[noteKey]

      return {
        ...record,
        measurementNote: note.trim() ? note : null,
      }
    }),
  }
}

function getRecordNoteOverrideKey(folderPath: string, record: MeasurementRecord) {
  return `${folderPath}::${record.relativePath || record.sourcePath}`
}

function getRecordImageOverrideKey(folderPath: string, record: MeasurementRecord) {
  return `${folderPath}::${record.relativePath || record.sourcePath}`
}

function getRecordImageSlots(overrides: RecordImageOverrides, folderPath: string, record: MeasurementRecord) {
  return normalizeRecordImageSlots(overrides[getRecordImageOverrideKey(folderPath, record)])
}

function createEmptyRecordImageSlots(): Array<RecordImageSlot | null> {
  return [null, null]
}

function readProjectModuleCache(): ProjectModuleCache {
  try {
    const parsed = JSON.parse(localStorage.getItem(projectModuleCacheKey) ?? '{}') as unknown

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      return {}
    }

    return Object.entries(parsed).reduce<ProjectModuleCache>((cache, [folderPath, entry]) => {
      if (!entry || typeof entry !== 'object') {
        return cache
      }

      const source = entry as Partial<ProjectModuleCacheEntry>
      const moduleId = typeof source.moduleId === 'string' ? source.moduleId : ''
      const moduleModel = typeof source.moduleModel === 'string' ? source.moduleModel : ''

      if (folderPath && moduleId && moduleModel) {
        cache[folderPath] = { moduleId, moduleModel }
      }

      return cache
    }, {})
  } catch {
    return {}
  }
}

function writeProjectModuleCache(cache: ProjectModuleCache) {
  try {
    localStorage.setItem(projectModuleCacheKey, JSON.stringify(cache))
  } catch {
    // Ignore storage errors; the app can still fall back to CSV module matching.
  }
}

function normalizeModuleModelForCache(model: string) {
  return model.trim().toLowerCase().replace(/\s+/g, ' ')
}

function invokeWithTimeout<T>(channel: string, timeoutMs: number, ...args: unknown[]): Promise<T> {
  let timeoutId: ReturnType<typeof setTimeout> | null = null
  const timeout = new Promise<never>((_resolve, reject) => {
    timeoutId = setTimeout(() => {
      reject(new Error(`Timed out after ${Math.round(timeoutMs / 1000)} seconds.`))
    }, timeoutMs)
  })

  return Promise.race([window.ipcRenderer.invoke(channel, ...args) as Promise<T>, timeout]).finally(() => {
    if (timeoutId) {
      clearTimeout(timeoutId)
    }
  })
}

function readProjectInfo(): ProjectInfo {
  try {
    return normalizeProjectInfo(JSON.parse(localStorage.getItem(projectInfoStorageKey) ?? 'null'), defaultProjectInfo)
  } catch {
    return { ...defaultProjectInfo }
  }
}

function writeProjectInfo(info: ProjectInfo) {
  try {
    localStorage.setItem(projectInfoStorageKey, JSON.stringify(normalizeProjectInfo(info, defaultProjectInfo)))
  } catch {
    // Ignore storage errors; the current project info stays in memory.
  }
}

function readProjectInfoLibrary(): ProjectInfoLibraryItem[] {
  try {
    const rawValue = localStorage.getItem(projectInfoLibraryStorageKey)

    if (rawValue === null) {
      return sortProjectInfoLibrary(seededProjectInfoLibrary.map((item) => ({ ...item })))
    }

    const parsed = JSON.parse(rawValue) as unknown

    if (!Array.isArray(parsed)) {
      return []
    }

    return sortProjectInfoLibrary(
      parsed
        .map(normalizeProjectInfoLibraryItem)
        .filter((item): item is ProjectInfoLibraryItem => item !== null),
    )
  } catch {
    return sortProjectInfoLibrary(seededProjectInfoLibrary.map((item) => ({ ...item })))
  }
}

function writeProjectInfoLibrary(items: ProjectInfoLibraryItem[]) {
  try {
    localStorage.setItem(
      projectInfoLibraryStorageKey,
      JSON.stringify(
        sortProjectInfoLibrary(items)
          .map((item) => normalizeProjectInfoLibraryItem(item))
          .filter(Boolean),
      ),
    )
  } catch {
    // Ignore storage errors; saved templates still stay in memory for the current session.
  }
}

function normalizeProjectInfo(source: unknown, fallback: ProjectInfo): ProjectInfo {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return { ...fallback }
  }

  const entry = source as Partial<ProjectInfo>

  return {
    projectName: normalizeProjectInfoText(entry.projectName, fallback.projectName),
    investorName: normalizeProjectInfoText(entry.investorName, fallback.investorName),
    investorNameEnglish: normalizeProjectInfoText(entry.investorNameEnglish, fallback.investorNameEnglish),
    factoryOwnerName: normalizeProjectInfoText(entry.factoryOwnerName, fallback.factoryOwnerName),
    factoryOwnerNameEnglish: normalizeProjectInfoText(
      entry.factoryOwnerNameEnglish,
      fallback.factoryOwnerNameEnglish,
    ),
    reportTitle: normalizeProjectInfoText(entry.reportTitle, fallback.reportTitle),
    measurementTitle: normalizeProjectInfoText(entry.measurementTitle, fallback.measurementTitle),
    companyName: normalizeProjectInfoText(entry.companyName, fallback.companyName),
    companyAddress: normalizeProjectInfoText(entry.companyAddress, fallback.companyAddress),
    preparedBy: normalizeProjectInfoText(entry.preparedBy, fallback.preparedBy),
    checkedBy: normalizeProjectInfoText(entry.checkedBy, fallback.checkedBy),
    approvedBy: normalizeProjectInfoText(entry.approvedBy, fallback.approvedBy),
    ownerApproval: normalizeProjectInfoText(entry.ownerApproval, fallback.ownerApproval),
    consultantApproval: normalizeProjectInfoText(entry.consultantApproval, fallback.consultantApproval),
    contractorEpcApproval: normalizeProjectInfoText(entry.contractorEpcApproval, fallback.contractorEpcApproval),
    testerApproval: normalizeProjectInfoText(entry.testerApproval, fallback.testerApproval),
    applicableStandards: normalizeProjectInfoText(entry.applicableStandards, fallback.applicableStandards),
  }
}

function normalizeProjectInfoLibraryItem(source: unknown): ProjectInfoLibraryItem | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null
  }

  const entry = source as Partial<ProjectInfoLibraryItem>
  const id = normalizeProjectInfoText(entry.id, '')

  if (!id) {
    return null
  }

  const info = normalizeProjectInfo(entry, defaultProjectInfo)

  return {
    ...info,
    id,
    label: getProjectInfoLibraryLabel(info),
    updatedAt: normalizeProjectInfoText(entry.updatedAt, new Date().toISOString()),
  }
}

function normalizeProjectInfoText(value: unknown, fallback: string) {
  const text = typeof value === 'string' ? value : fallback

  return normalizeDisplayText(text)
}

function createProjectInfoFromLibraryItem(item: ProjectInfoLibraryItem): ProjectInfo {
  return normalizeProjectInfo(item, defaultProjectInfo)
}

function getProjectInfoLibraryLabel(info: ProjectInfo) {
  return (
    info.factoryOwnerName.trim() ||
    info.factoryOwnerNameEnglish.trim() ||
    info.projectName.trim() ||
    info.investorName.trim() ||
    `Project Info ${new Date().toLocaleDateString()}`
  )
}

function sortProjectInfoLibrary(items: ProjectInfoLibraryItem[]) {
  return [...items].sort((left, right) => {
    return (
      naturalCollator.compare(getProjectInfoInvestorSortKey(left), getProjectInfoInvestorSortKey(right)) ||
      naturalCollator.compare(getProjectInfoLibraryLabel(left), getProjectInfoLibraryLabel(right)) ||
      naturalCollator.compare(left.updatedAt, right.updatedAt)
    )
  })
}

function getProjectInfoInvestorSortKey(info: ProjectInfo) {
  return info.investorName.trim() || info.investorNameEnglish.trim() || 'Unknown investor'
}

function createProjectInfoLibraryId() {
  return globalThis.crypto?.randomUUID?.() ?? `project-info-${Date.now()}`
}

function readAppSession(): AppSessionState | null {
  try {
    return normalizeAppSession(JSON.parse(localStorage.getItem(appSessionStorageKey) ?? 'null'))
  } catch {
    return null
  }
}

function writeAppSession(session: AppSessionState) {
  try {
    localStorage.setItem(appSessionStorageKey, JSON.stringify(session))
  } catch {
    // Ignore storage errors; the app can still be used without session restore.
  }
}

function normalizeAppSession(source: unknown): AppSessionState | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null
  }

  const entry = source as Partial<AppSessionState>

  if (entry.version !== 1) {
    return null
  }

  return {
    version: 1,
    activeView: normalizeAppView(entry.activeView),
    folderPath: typeof entry.folderPath === 'string' && entry.folderPath.trim() ? entry.folderPath : null,
    isChatOpen: typeof entry.isChatOpen === 'boolean' ? entry.isChatOpen : true,
    selectedInverter: typeof entry.selectedInverter === 'string' ? entry.selectedInverter : '',
    selectedPvModuleId: typeof entry.selectedPvModuleId === 'string' ? entry.selectedPvModuleId : '',
    selectedRecordPath: typeof entry.selectedRecordPath === 'string' ? entry.selectedRecordPath : '',
    selectedRecordRelativePath:
      typeof entry.selectedRecordRelativePath === 'string' ? entry.selectedRecordRelativePath : '',
    selectedSystemGroup: typeof entry.selectedSystemGroup === 'string' ? entry.selectedSystemGroup : '',
  }
}

function normalizeAppView(value: unknown): AppView {
  return value === 'project-info' || value === 'pv-module' || value === 'home' ? value : 'home'
}

function readRecordNoteOverrides(): RecordNoteOverrides {
  try {
    const parsed = JSON.parse(localStorage.getItem(recordNoteOverridesStorageKey) ?? '{}') as unknown

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      return {}
    }

    return Object.entries(parsed).reduce<RecordNoteOverrides>((overrides, [key, value]) => {
      if (typeof value === 'string') {
        overrides[key] = value
      }

      return overrides
    }, {})
  } catch {
    return {}
  }
}

function writeRecordNoteOverrides(overrides: RecordNoteOverrides) {
  try {
    localStorage.setItem(recordNoteOverridesStorageKey, JSON.stringify(overrides))
  } catch {
    // Ignore storage errors; edited notes still stay in memory for the current session.
  }
}

function readRecordImageOverrides(): RecordImageOverrides {
  try {
    const parsed = JSON.parse(localStorage.getItem(recordImageOverridesStorageKey) ?? '{}') as unknown

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      return {}
    }

    return Object.entries(parsed).reduce<RecordImageOverrides>((overrides, [key, value]) => {
      const slots = normalizeRecordImageSlots(value)

      if (slots.some(Boolean)) {
        overrides[key] = slots
      }

      return overrides
    }, {})
  } catch {
    return {}
  }
}

function writeRecordImageOverrides(overrides: RecordImageOverrides) {
  try {
    localStorage.setItem(recordImageOverridesStorageKey, JSON.stringify(overrides))
  } catch {
    // Ignore storage errors; images still stay in memory for the current session.
  }
}

function normalizeRecordImageSlots(source: unknown): Array<RecordImageSlot | null> {
  if (!Array.isArray(source)) {
    return createEmptyRecordImageSlots()
  }

  return [normalizeRecordImageSlot(source[0]), normalizeRecordImageSlot(source[1])]
}

function normalizeRecordImageSlot(source: unknown): RecordImageSlot | null {
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    return null
  }

  const entry = source as Partial<RecordImageSlot>

  if (typeof entry.dataUrl !== 'string' || !entry.dataUrl.startsWith('data:image/')) {
    return null
  }

  const legacyPolygon = normalizeRecordImagePolygon(entry.polygon)
  const shapes = normalizeRecordImageShapes(entry.shapes)

  return {
    dataUrl: entry.dataUrl,
    fit: 'contain',
    name: typeof entry.name === 'string' ? entry.name : 'Image',
    offsetX: typeof entry.offsetX === 'number' && Number.isFinite(entry.offsetX) ? clampNumber(entry.offsetX, -1, 1) : 0,
    offsetY: typeof entry.offsetY === 'number' && Number.isFinite(entry.offsetY) ? clampNumber(entry.offsetY, -1, 1) : 0,
    polygon: [],
    shapes: shapes.length > 0 ? shapes : legacyPolygon.length > 0 ? [{ color: defaultRecordImageShapeColor, points: legacyPolygon, type: 'polygon' }] : [],
    strokeWidth: normalizeRecordImageStrokeWidth(entry.strokeWidth),
    zoom: getRecordImageZoom(entry.zoom),
  }
}

function normalizeRecordImageShapes(source: unknown): RecordImageShape[] {
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
      const points = normalizeRecordImagePolygon(entry.points)

      if (!type || points.length === 0 || (type === 'rectangle' && points.length < 2)) {
        return null
      }

      return {
        color: normalizeRecordImageColor(entry.color),
        points: type === 'rectangle' ? points.slice(0, 2) : points,
        type,
      }
    })
    .filter((shape): shape is RecordImageShape => shape !== null)
}

function normalizeRecordImagePolygon(source: unknown): RecordImagePoint[] {
  if (!Array.isArray(source)) {
    return []
  }

  return source
    .map((point) => {
      if (!point || typeof point !== 'object' || Array.isArray(point)) {
        return null
      }

      const entry = point as Partial<RecordImagePoint>
      const x = typeof entry.x === 'number' && Number.isFinite(entry.x) ? clampNumber(entry.x, 0, 1) : null
      const y = typeof entry.y === 'number' && Number.isFinite(entry.y) ? clampNumber(entry.y, 0, 1) : null

      return x === null || y === null ? null : { x, y }
    })
    .filter((point): point is RecordImagePoint => point !== null)
}

function readToleranceRows(): ToleranceRow[] {
  try {
    return normalizeToleranceRows(JSON.parse(localStorage.getItem(toleranceRowsStorageKey) ?? 'null'))
  } catch {
    return cloneToleranceRows()
  }
}

function writeToleranceRows(rows: ToleranceRow[]) {
  try {
    localStorage.setItem(toleranceRowsStorageKey, JSON.stringify(normalizeToleranceRows(rows)))
  } catch {
    // Ignore storage errors; edited values still stay in memory for the current session.
  }
}

export default App
