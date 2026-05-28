import type { IvPoint, MeasurementRecord, MeasurementSummary, PvModule } from '../types/data'

const STC_IRRADIANCE_WM2 = 1000
const STC_CELL_TEMPERATURE_C = 25
const EMPTY_MEASUREMENT_SUMMARY: MeasurementSummary = {
  pmaxW: null,
  vmppV: null,
  imppA: null,
  vocV: null,
  iscA: null,
}

interface StcInputs {
  irradianceWm2: number
  cellTemperatureC: number
  alphaPercentPerC: number
  betaPercentPerC: number
  gammaPercentPerC: number
}

export function findPvModuleForRecord(record: MeasurementRecord | null, modules: PvModule[]) {
  if (!record?.moduleModel) {
    return null
  }

  const exactModelKey = normalizeModuleModel(record.moduleModel)
  const compactModelKey = compactModuleModel(record.moduleModel)

  return (
    modules.find((pvModule) => normalizeModuleModel(pvModule.model) === exactModelKey) ??
    modules.find((pvModule) => compactModuleModel(pvModule.model) === compactModelKey) ??
    null
  )
}

export function getStcConversionProblem(record: MeasurementRecord, pvModule: PvModule | null) {
  if (record.ivMeasurements.length === 0) {
    return 'No IV curve data.'
  }

  return getStcInputProblem(record, pvModule)
}

export function getStcSummaryConversionProblem(record: MeasurementRecord, pvModule: PvModule | null) {
  if (!hasMeasurementSummary(record.measurementSummary)) {
    return 'No measurement summary data.'
  }

  return getStcInputProblem(record, pvModule)
}

function getStcInputProblem(record: MeasurementRecord, pvModule: PvModule | null) {
  if (!pvModule) {
    return 'PV module model is not available in the library.'
  }

  if (!isPositiveFiniteNumber(record.irradianceWm2)) {
    return 'Measured irradiance must be greater than 0 W/m^2.'
  }

  if (!isFiniteNumber(record.cellTemperatureC)) {
    return 'Measured cell temperature is missing.'
  }

  if (!isPositiveFiniteNumber(record.modulesInString)) {
    return 'Modules in string must be greater than 0.'
  }

  if (!isFiniteNumber(pvModule.temperatureCoefficientIscPercentPerC)) {
    return 'PV module alpha coefficient is missing.'
  }

  if (!isFiniteNumber(pvModule.temperatureCoefficientVocPercentPerC)) {
    return 'PV module beta coefficient is missing.'
  }

  if (!isFiniteNumber(pvModule.temperatureCoefficientPmaxPercentPerC)) {
    return 'PV module gamma coefficient is missing.'
  }

  const inputs = readStcInputs(record, pvModule)

  if (!inputs) {
    return 'STC conversion inputs are incomplete.'
  }

  if (!isValidStcFactor(getCurrentTemperatureFactor(inputs))) {
    return 'PV module alpha coefficient creates an invalid STC current factor.'
  }

  if (!isValidStcFactor(getVoltageTemperatureFactor(inputs))) {
    return 'PV module beta coefficient creates an invalid STC voltage factor.'
  }

  if (!isValidStcFactor(getPowerTemperatureFactor(inputs))) {
    return 'PV module gamma coefficient creates an invalid STC power factor.'
  }

  return null
}

export function convertMeasurementSummaryToStc(
  record: MeasurementRecord,
  pvModule: PvModule | null,
): MeasurementSummary | null {
  const inputs = readStcInputs(record, pvModule)
  const moduleCount = record.modulesInString
  const summary = record.measurementSummary ?? EMPTY_MEASUREMENT_SUMMARY

  if (!inputs || !isPositiveFiniteNumber(moduleCount) || getStcSummaryConversionProblem(record, pvModule)) {
    return null
  }

  const irradianceFactor = STC_IRRADIANCE_WM2 / inputs.irradianceWm2
  const currentTemperatureFactor = getCurrentTemperatureFactor(inputs)
  const voltageTemperatureFactor = getVoltageTemperatureFactor(inputs)
  const powerTemperatureFactor = getPowerTemperatureFactor(inputs)

  return {
    pmaxW: convertPowerToStc(summary.pmaxW, irradianceFactor, powerTemperatureFactor, moduleCount),
    vmppV: convertVoltageToStc(summary.vmppV, voltageTemperatureFactor, moduleCount),
    imppA: convertCurrentToStc(summary.imppA, irradianceFactor, currentTemperatureFactor),
    vocV: convertVoltageToStc(summary.vocV, voltageTemperatureFactor, moduleCount),
    iscA: convertCurrentToStc(summary.iscA, irradianceFactor, currentTemperatureFactor),
  }
}

export function convertIvMeasurementsToStc(record: MeasurementRecord, pvModule: PvModule | null): IvPoint[] {
  const inputs = readStcInputs(record, pvModule)
  const moduleCount = record.modulesInString

  if (!inputs || !isPositiveFiniteNumber(moduleCount) || getStcConversionProblem(record, pvModule)) {
    return []
  }

  const irradianceFactor = STC_IRRADIANCE_WM2 / inputs.irradianceWm2
  const currentTemperatureFactor = getCurrentTemperatureFactor(inputs)
  const voltageTemperatureFactor = getVoltageTemperatureFactor(inputs)
  const powerTemperatureFactor = getPowerTemperatureFactor(inputs)

  return record.ivMeasurements.map((point) => ({
    volts: point.volts / voltageTemperatureFactor / moduleCount,
    amps: (point.amps * irradianceFactor) / currentTemperatureFactor,
    watts: (point.watts * irradianceFactor) / powerTemperatureFactor / moduleCount,
  }))
}

export function buildPvModuleReferenceMeasurements(pvModule: PvModule | null): IvPoint[] {
  const iscA = pvModule?.shortCircuitCurrentA ?? null
  const vocV = pvModule?.openCircuitVoltageV ?? null

  if (!isPositiveFiniteNumber(iscA) || !isPositiveFiniteNumber(vocV)) {
    return []
  }

  return NORMALIZED_REFERENCE_IV_CURVE.map(([currentRatio, voltageRatio]) => {
    const amps = currentRatio * iscA
    const volts = voltageRatio * vocV

    return {
      volts,
      amps,
      watts: volts * amps,
    }
  })
}

export function getStcBasisLabel() {
  return `STC (${STC_IRRADIANCE_WM2} W/m^2, ${STC_CELL_TEMPERATURE_C} Deg C)`
}

function convertCurrentToStc(
  value: number | null,
  irradianceFactor: number,
  currentTemperatureFactor: number,
) {
  return value === null ? null : (value * irradianceFactor) / currentTemperatureFactor
}

function convertVoltageToStc(
  value: number | null,
  voltageTemperatureFactor: number,
  moduleCount: number,
) {
  return value === null ? null : value / voltageTemperatureFactor / moduleCount
}

function convertPowerToStc(
  value: number | null,
  irradianceFactor: number,
  powerTemperatureFactor: number,
  moduleCount: number,
) {
  return value === null ? null : (value * irradianceFactor) / powerTemperatureFactor / moduleCount
}

function hasMeasurementSummary(summary: MeasurementSummary | null | undefined) {
  if (!summary) {
    return false
  }

  return Object.values(summary).some((value) => value !== null)
}

function readStcInputs(record: MeasurementRecord, pvModule: PvModule | null): StcInputs | null {
  if (
    !pvModule ||
    !isPositiveFiniteNumber(record.irradianceWm2) ||
    !isFiniteNumber(record.cellTemperatureC) ||
    !isFiniteNumber(pvModule.temperatureCoefficientIscPercentPerC) ||
    !isFiniteNumber(pvModule.temperatureCoefficientVocPercentPerC) ||
    !isFiniteNumber(pvModule.temperatureCoefficientPmaxPercentPerC)
  ) {
    return null
  }

  return {
    irradianceWm2: record.irradianceWm2,
    cellTemperatureC: record.cellTemperatureC,
    alphaPercentPerC: pvModule.temperatureCoefficientIscPercentPerC,
    betaPercentPerC: pvModule.temperatureCoefficientVocPercentPerC,
    gammaPercentPerC: pvModule.temperatureCoefficientPmaxPercentPerC,
  }
}

function getCurrentTemperatureFactor(inputs: StcInputs) {
  return 1 + (inputs.alphaPercentPerC / 100) * (inputs.cellTemperatureC - STC_CELL_TEMPERATURE_C)
}

function getVoltageTemperatureFactor(inputs: StcInputs) {
  return 1 + (inputs.betaPercentPerC / 100) * (inputs.cellTemperatureC - STC_CELL_TEMPERATURE_C)
}

function getPowerTemperatureFactor(inputs: StcInputs) {
  return 1 + (inputs.gammaPercentPerC / 100) * (inputs.cellTemperatureC - STC_CELL_TEMPERATURE_C)
}

function isValidStcFactor(value: number) {
  return Number.isFinite(value) && value !== 0
}

function isFiniteNumber(value: number | null): value is number {
  return typeof value === 'number' && Number.isFinite(value)
}

function isPositiveFiniteNumber(value: number | null): value is number {
  return isFiniteNumber(value) && value > 0
}

function normalizeModuleModel(model: string) {
  return model.trim().toLowerCase().replace(/\s+/g, ' ')
}

function compactModuleModel(model: string) {
  return model.toLowerCase().replace(/[^a-z0-9]/g, '')
}

const NORMALIZED_REFERENCE_IV_CURVE: Array<[number, number]> = [
  [1, 0],
  [1, 0.016774],
  [1, 0.033548],
  [1, 0.050323],
  [1, 0.067097],
  [1, 0.083871],
  [1, 0.100645],
  [1, 0.117419],
  [1, 0.134194],
  [1, 0.150968],
  [1, 0.167742],
  [1, 0.184516],
  [1, 0.20129],
  [1, 0.218065],
  [1, 0.234839],
  [1, 0.251613],
  [1, 0.268387],
  [1, 0.285161],
  [1, 0.301935],
  [1, 0.31871],
  [1, 0.335484],
  [1, 0.352258],
  [1, 0.369032],
  [1, 0.385806],
  [1, 0.402581],
  [1, 0.419355],
  [0.999999, 0.436129],
  [0.999998, 0.452903],
  [0.999997, 0.469677],
  [0.999995, 0.486452],
  [0.999991, 0.503226],
  [0.999984, 0.52],
  [0.999973, 0.536774],
  [0.999955, 0.553548],
  [0.999926, 0.570323],
  [0.999879, 0.587097],
  [0.999804, 0.603871],
  [0.999689, 0.620645],
  [0.999511, 0.637419],
  [0.999242, 0.654194],
  [0.998836, 0.670968],
  [0.998231, 0.687742],
  [0.997341, 0.704516],
  [0.996039, 0.72129],
  [0.994154, 0.738065],
  [0.991448, 0.754839],
  [0.987593, 0.771613],
  [0.982143, 0.788387],
  [0.974497, 0.805161],
  [0.963842, 0.821935],
  [0.949097, 0.83871],
  [0.945679, 0.841935],
  [0.942045, 0.845161],
  [0.938185, 0.848387],
  [0.934082, 0.851613],
  [0.929725, 0.854839],
  [0.925098, 0.858065],
  [0.920185, 0.86129],
  [0.914971, 0.864516],
  [0.909437, 0.867742],
  [0.903565, 0.870968],
  [0.897336, 0.874194],
  [0.890731, 0.877419],
  [0.883727, 0.880645],
  [0.876302, 0.883871],
  [0.868432, 0.887097],
  [0.860094, 0.890323],
  [0.85126, 0.893548],
  [0.841903, 0.896774],
  [0.831994, 0.9],
  [0.821503, 0.903226],
  [0.810397, 0.906452],
  [0.798644, 0.909677],
  [0.786208, 0.912903],
  [0.773052, 0.916129],
  [0.759137, 0.919355],
  [0.744422, 0.922581],
  [0.728864, 0.925806],
  [0.712418, 0.929032],
  [0.695037, 0.932258],
  [0.676671, 0.935484],
  [0.657268, 0.93871],
  [0.636773, 0.941935],
  [0.615129, 0.945161],
  [0.592275, 0.948387],
  [0.56815, 0.951613],
  [0.542686, 0.954839],
  [0.515813, 0.958065],
  [0.48746, 0.96129],
  [0.45755, 0.964516],
  [0.426003, 0.967742],
  [0.392736, 0.970968],
  [0.357661, 0.974194],
  [0.320687, 0.977419],
  [0.281716, 0.980645],
  [0.240649, 0.983871],
  [0.197379, 0.987097],
  [0.151798, 0.990323],
  [0.10378, 0.993548],
  [0, 1],
]
