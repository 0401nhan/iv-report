export interface ToleranceColumn {
  key: string
  label: string
  digits: number
}

export interface ToleranceRow {
  label: string
  values: Record<string, number>
}

export const TOLERANCE_TOTAL_KEY = 'degradationTotal'

export const TOLERANCE_COLUMNS: ToleranceColumn[] = [
  { key: 'temperatureSensor', label: 'A (the temp. sensor tolerance)', digits: 1 },
  { key: 'radiationSensor', label: 'B (The radiation sensor tolerance)', digits: 1 },
  { key: 'pmaxStcTolerance', label: 'C ( The Pmax STC Tolerance)', digits: 1 },
  { key: 'initialModuleDegradation', label: 'D (The initial module degradation)', digits: 1 },
  { key: 'dcCableLoss', label: 'E(DC cable loss)', digits: 1 },
  { key: 'otherLoss', label: 'F(Other loss)', digits: 1 },
  { key: 'degradationPerTime', label: 'Degradation per time (Data sheet)', digits: 2 },
  { key: TOLERANCE_TOTAL_KEY, label: 'Degradation total', digits: 1 },
]

export const TOLERANCE_ROWS: ToleranceRow[] = [
  {
    label: 'Tol+',
    values: {
      temperatureSensor: 1,
      radiationSensor: 1,
      pmaxStcTolerance: 3,
      initialModuleDegradation: 0,
      dcCableLoss: 0,
      otherLoss: 0,
      degradationPerTime: 0,
      [TOLERANCE_TOTAL_KEY]: 5,
    },
  },
  {
    label: 'Tol-',
    values: {
      temperatureSensor: 1,
      radiationSensor: 1,
      pmaxStcTolerance: 0,
      initialModuleDegradation: 1,
      dcCableLoss: 0,
      otherLoss: 0,
      degradationPerTime: 1.6,
      [TOLERANCE_TOTAL_KEY]: 4.6,
    },
  },
]

export function cloneToleranceRows(rows: ToleranceRow[] = TOLERANCE_ROWS): ToleranceRow[] {
  return rows.map((row) => ({
    label: row.label,
    values: { ...row.values },
  }))
}

export function normalizeToleranceRows(source: unknown): ToleranceRow[] {
  const defaults = cloneToleranceRows()

  if (!Array.isArray(source)) {
    return defaults
  }

  return defaults.map((defaultRow) => {
    const sourceRow = source.find((row): row is Partial<ToleranceRow> => {
      return Boolean(row) && typeof row === 'object' && (row as Partial<ToleranceRow>).label === defaultRow.label
    })

    if (!sourceRow?.values || typeof sourceRow.values !== 'object') {
      return defaultRow
    }

    const values = { ...defaultRow.values }

    TOLERANCE_COLUMNS.forEach((column) => {
      const value = sourceRow.values?.[column.key]

      if (typeof value === 'number' && Number.isFinite(value)) {
        values[column.key] = value
      }
    })

    values[TOLERANCE_TOTAL_KEY] = calculateToleranceRowTotal({ label: defaultRow.label, values })

    return {
      label: defaultRow.label,
      values,
    }
  })
}

export function calculateToleranceRowTotal(row: ToleranceRow) {
  return TOLERANCE_COLUMNS.reduce((total, column) => {
    if (column.key === TOLERANCE_TOTAL_KEY) {
      return total
    }

    const value = row.values[column.key]

    return total + (Number.isFinite(value) ? value : 0)
  }, 0)
}

export function getToleranceCellValue(row: ToleranceRow, columnKey: string) {
  return columnKey === TOLERANCE_TOTAL_KEY ? calculateToleranceRowTotal(row) : row.values[columnKey]
}

export function formatTolerancePercent(value: number, digits: number) {
  return `${value.toFixed(digits)}%`
}
