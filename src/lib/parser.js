import * as XLSX from 'xlsx'

const MAX_CHART_POINTS = 50000
const MBAR_PER_PSI = 68.9476
const OIL_PRESSURE_MBAR_THRESHOLD = 600

const SIGNALS = {
  kw: { label: 'Generator kW', aliases: ['generator kw', 'kw'], unit: 'kW', min: -10000, max: 10000 },
  rpm: { label: 'RPM', aliases: ['rpm', 'engine speed'], unit: 'rpm', min: 0, max: 5000 },
  frequency: { label: 'Frequency', aliases: ['generator frequency', 'frequency hz'], exclude: ['switch'], unit: 'Hz', min: 0, max: 90 },
  v12: { label: 'Voltage L1-L2', aliases: ['voltage l1-l2', 'v l1-l2', 'vab'], unit: 'V', min: 0, max: 15000 },
  v23: { label: 'Voltage L2-L3', aliases: ['voltage l2-l3', 'v l2-l3', 'vbc'], unit: 'V', min: 0, max: 15000 },
  v31: { label: 'Voltage L3-L1', aliases: ['voltage l3-l1', 'v l3-l1', 'vca'], unit: 'V', min: 0, max: 15000 },
  i1: { label: 'Current L1', aliases: ['current l1', 'amps l1', 'i l1'], unit: 'A', min: 0, max: 20000 },
  i2: { label: 'Current L2', aliases: ['current l2', 'amps l2', 'i l2'], unit: 'A', min: 0, max: 20000 },
  i3: { label: 'Current L3', aliases: ['current l3', 'amps l3', 'i l3'], unit: 'A', min: 0, max: 20000 },
  powerFactor: { label: 'Power Factor', aliases: ['power factor'], unit: '', min: -1, max: 1 },
  kva: { label: 'Generator kVA', aliases: ['generator kva', 'kva'], unit: 'kVA', min: -10000, max: 10000 },
  coolantTemp: {
    label: 'Coolant Temp',
    aliases: ['t-coolant', 'coolant temp', 'coolant temperature', 'coolant temperature at engine output', 'engine coolant temperature physical'],
    preferred: ['coolant temperature at engine output', 'engine coolant temperature physical', 't-coolant'],
    unit: 'deg',
    min: -40,
    max: 260,
  },
  oilPressure: {
    label: 'Oil Pressure',
    aliases: ['p-oil', 'oil pressure value filtered', 'engine oil pressure after oil filter', 'oil pressure'],
    preferred: ['p-oil', 'oil pressure value filtered', 'engine oil pressure after oil filter'],
    exclude: ['raw', 'voltage', 'adc', 'physical value'],
    unit: 'PSI',
    min: -50,
    max: 15000,
  },
  oilTemp: {
    label: 'Oil Temp',
    aliases: ['t-oil', 'sensed value of oil temperature', 'oil temperature in oil sump', 'oil temp', 'oil temperature'],
    preferred: ['t-oil', 'sensed value of oil temperature', 'oil temperature in oil sump'],
    exclude: ['raw voltage'],
    unit: 'deg',
    min: -40,
    max: 300,
  },
  engineState: { label: 'Engine State', aliases: ['engine state'], unit: '', min: -100, max: 1000 },
  engineOperation: { label: 'Engine Operation', aliases: ['engine operation status'], unit: '', min: null, max: null, type: 'text' },
  breakerState: { label: 'Breaker State', aliases: ['breaker state'], unit: '', min: -100, max: 1000 },
  starter1: { label: 'Starter 1', aliases: ['starter 1', 'starter'], unit: '', min: 0, max: 1 },
  speedRequest: { label: 'Speed Request', aliases: ['speed request'], exclude: ['fan external'], unit: '%', min: 0, max: 120 },
  voltageRequest: { label: 'Voltage Request', aliases: ['voltage request'], unit: '%', min: 0, max: 120 },
  intakeTemp: { label: 'Intake Manifold Temp', aliases: ['t-intmanifold', 'int manifold', 'intake manifold'], unit: 'deg', min: -40, max: 260 },
  preLubePump: { label: 'Pre-Lube Pump', aliases: ['pre-lube pump', 'pre lube pump'], unit: '', min: 0, max: 1 },
  batteryVoltage: {
    label: 'Battery Voltage',
    aliases: ['sensed battery voltage', 'battery voltage before defect', 'battery voltage'],
    preferred: ['sensed battery voltage', 'battery voltage before defect', 'battery voltage'],
    unit: 'V',
    min: 0,
    max: 60,
    scale: (value) => (value > 1000 ? value / 1000 : value),
  },
  fuelPressure: {
    label: 'Fuel Pressure',
    aliases: ['fuel pressure before the fine filter', 'fuel pressure after the fine filter', 'fuel pressure'],
    preferred: ['fuel pressure before the fine filter', 'fuel pressure after the fine filter', 'fuel pressure'],
    exclude: ['raw', 'voltage'],
    unit: 'kPa',
    min: -100,
    max: 200000,
  },
  railPressure: {
    label: 'Rail Pressure',
    aliases: ['maximum rail pressure', 'raw value of rail pressure', 'rail pressure'],
    preferred: ['maximum rail pressure', 'rail pressure'],
    unit: 'kPa',
    min: 0,
    max: 2500000,
  },
  railPressureSetpoint: { label: 'Rail Pressure Setpoint', aliases: ['rail pressure setpoint'], unit: 'kPa', min: 0, max: 2500000 },
  calculatedLoad: { label: 'Calculated Load', aliases: ['calculated load value'], unit: '%', min: 0, max: 150 },
  actualTorque: { label: 'Actual Torque', aliases: ['actual engine torque'], unit: 'Nm', min: -100000, max: 100000 },
  turboSpeed: { label: 'Turbo Speed', aliases: ['turbo charger speed3', 'turbo charger speed'], preferred: ['turbo charger speed', 'turbo charger speed3'], unit: 'rpm', min: 0, max: 250000 },
  turboTemp: { label: 'Turbo Upstream Temp', aliases: ['turbo upstream temperature', 'turbo upstream tempreature'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempA1: { label: 'Exhaust Temp A1', aliases: ['exhaust gas tempreature for a1'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempA2: { label: 'Exhaust Temp A2', aliases: ['exhaust gas tempreature for a2'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempA3: { label: 'Exhaust Temp A3', aliases: ['exhaust gas tempreature for a3'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempA4: { label: 'Exhaust Temp A4', aliases: ['exhaust gas tempreature for a4'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempB1: { label: 'Exhaust Temp B1', aliases: ['exhaust gas tempreature for b1'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempB2: { label: 'Exhaust Temp B2', aliases: ['exhaust gas tempreature for b2'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempB3: { label: 'Exhaust Temp B3', aliases: ['exhaust gas tempreature for b3'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
  exhaustTempB4: { label: 'Exhaust Temp B4', aliases: ['exhaust gas tempreature for b4'], exclude: ['raw'], unit: 'deg', min: -50, max: 1000 },
}

const DATE_PATTERNS = [
  /^(\d{1,2})\.(\d{1,2})\.(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?$/,
  /^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?/,
  /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2}):(\d{2})(?:\.(\d{1,3}))?/,
]

export function parseLogFile({ buffer, text, fileName }) {
  const extension = (fileName.split('.').pop() || '').toLowerCase()
  if (['xlsx', 'xls', 'xlsm'].includes(extension)) {
    const workbook = XLSX.read(buffer, { cellDates: true })
    return analyzeWorkbook(workbook, fileName)
  }
  return analyzeTextLog(text, fileName)
}

function analyzeWorkbook(workbook, fileName) {
  const sheets = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name]
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: false, blankrows: false })
    return buildSheetProfile(name, rows)
  })

  const candidates = sheets
    .map((sheet) => ({ sheet, score: scoreSheet(sheet) }))
    .filter((item) => item.score > 0)
    .sort((a, b) => b.score - a.score)

  if (!candidates.length) {
    return {
      fileName,
      sheets,
      error: 'No timestamped data sheet was detected. Open Raw Tabs to inspect the workbook layout.',
    }
  }

  const selectedSheet = candidates[0].sheet
  return buildAnalysis(fileName, sheets, selectedSheet)
}

function analyzeTextLog(text, fileName) {
  const rows = parseDelimitedText(text)
  const sheet = buildSheetProfile('Imported Log', rows)
  return buildAnalysis(fileName, [sheet], sheet)
}

function buildSheetProfile(name, rows) {
  const normalizedRows = rows.map((row) => (Array.isArray(row) ? row.map(cleanCell) : []))
  const dimensions = normalizedRows.reduce(
    (acc, row) => {
      acc.rows += row.some((cell) => cell !== null && cell !== '')
      acc.cols = Math.max(acc.cols, row.length)
      return acc
    },
    { rows: 0, cols: 0 },
  )
  const headerIndex = detectHeaderIndex(normalizedRows)
  const headers = headerIndex >= 0 ? normalizedRows[headerIndex].map((header, index) => normalizeHeader(header, index)) : []
  const timeColumn = headers.findIndex((header) => isTimeHeader(header))
  return {
    name,
    rows: dimensions.rows,
    cols: dimensions.cols,
    headerIndex,
    headers,
    timeColumn,
    sampleRows: normalizedRows.slice(Math.max(0, headerIndex), Math.max(6, headerIndex + 6)),
    tableRows: normalizedRows,
    dataRows: headerIndex >= 0 ? normalizedRows.slice(headerIndex + 1) : [],
  }
}

function buildAnalysis(fileName, sheets, selectedSheet) {
  const channelMap = mapChannels(selectedSheet.headers)
  const parsed = normalizeRows(selectedSheet, channelMap)
  const series = parsed.series.sort((a, b) => a.timestamp - b.timestamp)
  convertOilPressureToPsi(series)
  const enriched = enrichSeries(series)
  const stats = computeStats(enriched, parsed.invalidCounts)
  const events = detectEvents(enriched, stats)
  const chartData = downsample(enriched, MAX_CHART_POINTS)
  const parameterStats = computeParameterStats(selectedSheet, channelMap)
  const availableSignals = Object.entries(channelMap)
    .filter(([, index]) => index >= 0)
    .map(([key, index]) => ({ key, index, ...SIGNALS[key], source: selectedSheet.headers[index] }))

  return {
    fileName,
    selectedSheet: selectedSheet.name,
    sheets: sheets.map(({ dataRows, ...sheet }) => sheet),
    channels: availableSignals,
    stats,
    events,
    parameterStats,
    samples: enriched,
    chartData,
    chartMeta: {
      fullSampleCount: enriched.length,
      chartPointCount: chartData.length,
      downsampled: chartData.length < enriched.length,
    },
    latestRows: enriched.slice(-250).reverse(),
    rawPreview: selectedSheet.dataRows.slice(0, 200),
    headers: selectedSheet.headers,
    invalidCounts: parsed.invalidCounts,
  }
}

function computeParameterStats(sheet, channelMap) {
  const mappedByIndex = Object.fromEntries(
    Object.entries(channelMap)
      .filter(([, index]) => index >= 0)
      .map(([signal, index]) => [index, SIGNALS[signal].label]),
  )

  return sheet.headers.map((header, columnIndex) => {
    if (columnIndex === sheet.timeColumn) {
      return {
        index: columnIndex,
        header,
        mappedSignal: 'Timestamp',
        type: 'time',
        count: sheet.dataRows.length,
        min: null,
        avg: null,
        max: null,
        samples: [],
      }
    }

    const numericValues = []
    const textValues = new Set()
    let populated = 0

    for (const row of sheet.dataRows) {
      const value = row[columnIndex]
      if (value === null || value === undefined || value === '') continue
      populated += 1
      const numeric = parseNumeric(value)
      if (numeric !== null && Math.abs(numeric) < 900000000) {
        numericValues.push(numeric)
      } else if (textValues.size < 8) {
        textValues.add(String(value))
      }
    }

    const isNumeric = numericValues.length >= Math.max(2, populated * 0.7)
    return {
      index: columnIndex,
      header,
      mappedSignal: mappedByIndex[columnIndex] || '',
      type: isNumeric ? 'numeric' : 'text',
      count: populated,
      min: isNumeric ? Math.min(...numericValues) : null,
      avg: isNumeric ? average(numericValues) : null,
      max: isNumeric ? Math.max(...numericValues) : null,
      samples: isNumeric ? [] : Array.from(textValues),
    }
  })
}

function normalizeRows(sheet, channelMap) {
  const series = []
  const invalidCounts = {}
  let previousTimestamp = null

  sheet.dataRows.forEach((row, rowIndex) => {
    const parsedDate = parseDate(row[sheet.timeColumn], previousTimestamp)
    if (!parsedDate) return
    previousTimestamp = parsedDate.getTime()

    const point = {
      rowNumber: sheet.headerIndex + rowIndex + 2,
      timestamp: parsedDate.getTime(),
      timeLabel: parsedDate.toLocaleTimeString(),
      isoTime: parsedDate.toISOString(),
    }

    Object.entries(channelMap).forEach(([signal, columnIndex]) => {
      if (columnIndex < 0) return
      const raw = row[columnIndex]
      if (SIGNALS[signal]?.type === 'text') {
        point[signal] = raw === null || raw === undefined ? null : String(raw)
        return
      }
      const value = applyScale(signal, parseNumeric(raw))
      const clean = cleanNumeric(signal, value)
      if (value !== null && clean === null) {
        invalidCounts[signal] = (invalidCounts[signal] || 0) + 1
      }
      point[signal] = clean
    })

    series.push(point)
  })

  return { series, invalidCounts }
}

function convertOilPressureToPsi(series) {
  const values = series.map((point) => point.oilPressure).filter(isFiniteNumber)
  if (!values.length) return
  const avg = average(values)
  if (avg <= OIL_PRESSURE_MBAR_THRESHOLD) return
  for (const point of series) {
    if (isFiniteNumber(point.oilPressure)) {
      point.oilPressure = point.oilPressure / MBAR_PER_PSI
    }
  }
}

function enrichSeries(series) {
  let previous = null
  const firstTimestamp = series[0]?.timestamp || 0
  return series.map((point) => {
    const voltageValues = [point.v12, point.v23, point.v31].filter(isFiniteNumber)
    const currentValues = [point.i1, point.i2, point.i3].filter(isFiniteNumber)
    const exhaustTemps = [
      point.exhaustTempA1,
      point.exhaustTempA2,
      point.exhaustTempA3,
      point.exhaustTempA4,
      point.exhaustTempB1,
      point.exhaustTempB2,
      point.exhaustTempB3,
      point.exhaustTempB4,
    ].filter(isFiniteNumber)
    const voltageAvg = average(voltageValues)
    const currentAvg = average(currentValues)
    const enriched = {
      ...point,
      elapsedSeconds: firstTimestamp ? (point.timestamp - firstTimestamp) / 1000 : 0,
      voltageAvg,
      currentAvg,
      exhaustTempAvg: average(exhaustTemps),
      exhaustTempSpread: exhaustTemps.length ? Math.max(...exhaustTemps) - Math.min(...exhaustTemps) : null,
      voltageImbalance: imbalancePercent(voltageValues),
      currentImbalance: imbalancePercent(currentValues),
      engineRunning: (point.rpm || 0) > 900 || (point.frequency || 0) > 30,
      stableRunning: (point.rpm || 0) > 1200 || (point.frequency || 0) > 50,
    }

    if (previous) {
      const seconds = (point.timestamp - previous.timestamp) / 1000
      if (seconds > 0 && seconds < 10) {
        if (isFiniteNumber(point.oilPressure) && isFiniteNumber(previous.oilPressure)) {
          enriched.oilPressureRate = (point.oilPressure - previous.oilPressure) / seconds
        }
        if (isFiniteNumber(point.rpm) && isFiniteNumber(previous.rpm)) {
          enriched.rpmRate = (point.rpm - previous.rpm) / seconds
        }
      }
    }
    previous = point
    return enriched
  })
}

function computeStats(series, invalidCounts) {
  const started = series.find((point) => point.engineRunning)
  const stopped = [...series].reverse().find((point) => point.engineRunning)
  const first = series[0]
  const last = series[series.length - 1]
  const statsBySignal = {}

  Object.keys(SIGNALS).forEach((signal) => {
    const values = series.map((point) => point[signal]).filter(isFiniteNumber)
    if (!values.length) return
    statsBySignal[signal] = {
      ...SIGNALS[signal],
      min: Math.min(...values),
      max: Math.max(...values),
      avg: average(values),
      count: values.length,
    }
  })
  if (statsBySignal.oilPressure) {
    statsBySignal.oilPressure.unit = 'PSI'
  }

  return {
    sampleCount: series.length,
    startTime: first?.isoTime || null,
    endTime: last?.isoTime || null,
    durationMs: first && last ? last.timestamp - first.timestamp : 0,
    runDurationMs: started && stopped ? stopped.timestamp - started.timestamp : 0,
    runningSamples: series.filter((point) => point.engineRunning).length,
    invalidCount: Object.values(invalidCounts).reduce((sum, count) => sum + count, 0),
    invalidCounts,
    bySignal: statsBySignal,
    maxVoltageImbalance: maxOf(series, 'voltageImbalance'),
    maxCurrentImbalance: maxOf(series, 'currentImbalance'),
    minOilPressureWhileRunning: minOf(series.filter((point) => point.stableRunning), 'oilPressure'),
    maxCoolantTemp: maxOf(series, 'coolantTemp'),
    maxOilTemp: maxOf(series, 'oilTemp'),
  }
}

function detectEvents(series, stats) {
  const events = []
  let previous = null
  let wasRunning = false
  const activeAlerts = {}
  const lowOilThreshold = 20
  const rapidOilDropThreshold = -40

  series.forEach((point) => {
    if (point.engineRunning && !wasRunning) {
      events.push(makeEvent('info', 'Engine running detected', point, `RPM ${fmt(point.rpm)}, frequency ${fmt(point.frequency)} Hz`))
    }
    if (!point.engineRunning && wasRunning) {
      events.push(makeEvent('info', 'Engine stopped', point, `RPM ${fmt(point.rpm)}, frequency ${fmt(point.frequency)} Hz`))
    }

    alertOnTransition(
      activeAlerts,
      events,
      'lowOil',
      point.stableRunning && point.oilPressure !== null && point.oilPressure < lowOilThreshold,
      makeEvent('critical', 'Low oil pressure while running', point, `${fmt(point.oilPressure)} ${stats.bySignal.oilPressure?.unit || ''}`),
    )
    if (point.oilPressureRate !== undefined && point.oilPressureRate <= rapidOilDropThreshold) {
      events.push(makeEvent('critical', 'Rapid oil pressure drop', point, `${fmt(point.oilPressureRate)} ${stats.bySignal.oilPressure?.unit || ''}/sec`))
    }
    alertOnTransition(
      activeAlerts,
      events,
      'highCoolant',
      point.engineRunning && point.coolantTemp !== null && point.coolantTemp >= 110,
      makeEvent('warning', 'High coolant temperature', point, `${fmt(point.coolantTemp)} deg`),
    )
    alertOnTransition(
      activeAlerts,
      events,
      'highOilTemp',
      point.engineRunning && point.oilTemp !== null && point.oilTemp >= 130,
      makeEvent('warning', 'High oil temperature', point, `${fmt(point.oilTemp)} deg`),
    )
    alertOnTransition(
      activeAlerts,
      events,
      'frequencyBand',
      point.stableRunning && point.frequency !== null && (point.frequency < 59 || point.frequency > 61),
      makeEvent('warning', 'Frequency outside 59-61 Hz', point, `${fmt(point.frequency)} Hz`),
    )
    if (point.engineRunning && point.rpm !== null && point.rpm > 1850) {
      events.push(makeEvent('warning', 'Possible overspeed', point, `${fmt(point.rpm)} RPM`))
    }
    alertOnTransition(
      activeAlerts,
      events,
      'highExhaustTemp',
      point.exhaustTempAvg !== null && point.exhaustTempAvg >= 650,
      makeEvent('warning', 'High average exhaust temperature', point, `${fmt(point.exhaustTempAvg)} deg`),
    )
    alertOnTransition(
      activeAlerts,
      events,
      'exhaustTempSpread',
      point.exhaustTempSpread !== null && point.exhaustTempSpread >= 125,
      makeEvent('warning', 'Exhaust temperature spread', point, `${fmt(point.exhaustTempSpread)} deg`),
    )
    alertOnTransition(
      activeAlerts,
      events,
      'voltageImbalance',
      point.stableRunning && point.voltageAvg > 100 && point.voltageImbalance !== null && point.voltageImbalance > 3,
      makeEvent('warning', 'Voltage phase imbalance', point, `${fmt(point.voltageImbalance)}%`),
    )

    for (const signal of ['engineState', 'engineOperation', 'breakerState', 'starter1', 'preLubePump']) {
      if (previous && point[signal] !== null && previous[signal] !== null && point[signal] !== previous[signal]) {
        events.push(makeEvent('state', `${SIGNALS[signal].label} changed`, point, `${previous[signal]} -> ${point[signal]}`))
      }
    }

    wasRunning = point.engineRunning
    previous = point
  })

  if (stats.invalidCount > 0) {
    events.unshift({
      severity: 'data',
      title: 'Invalid or sentinel values removed',
      time: stats.endTime,
      rowNumber: null,
      detail: `${stats.invalidCount} values were excluded from charts and statistics.`,
    })
  }

  return collapseRepeatedEvents(events).slice(0, 300)
}

function alertOnTransition(activeAlerts, events, key, condition, event) {
  if (condition && !activeAlerts[key]) {
    events.push(event)
    activeAlerts[key] = true
  } else if (!condition) {
    activeAlerts[key] = false
  }
}

function collapseRepeatedEvents(events) {
  const collapsed = []
  let previousKey = ''
  let repeat = 0

  events.forEach((event) => {
    const key = `${event.title}|${event.severity}`
    if (key === previousKey) {
      repeat += 1
      const last = collapsed[collapsed.length - 1]
      last.repeatCount = repeat + 1
      last.detail = `${last.baseDetail || last.detail} (${last.repeatCount} consecutive samples)`
      return
    }
    repeat = 0
    previousKey = key
    collapsed.push({ ...event, baseDetail: event.detail })
  })

  return collapsed.map(({ baseDetail, ...event }) => event)
}

function mapChannels(headers) {
  const lowerHeaders = headers.map((header) => String(header || '').toLowerCase())
  const map = {}
  Object.entries(SIGNALS).forEach(([key, config]) => {
    map[key] = bestHeaderIndex(lowerHeaders, config)
  })
  return map
}

function bestHeaderIndex(lowerHeaders, config) {
  let best = { index: -1, score: 0 }
  lowerHeaders.forEach((header, index) => {
    let score = 0
    for (const alias of config.aliases || []) {
      if (header.includes(alias)) score = Math.max(score, 10 + alias.length / 10)
    }
    for (const alias of config.preferred || []) {
      if (header.includes(alias)) score += 40 - config.preferred.indexOf(alias) * 5 + alias.length / 10
    }
    for (const rejected of config.exclude || []) {
      if (header.includes(rejected)) score -= 25
    }
    if (score > best.score) best = { index, score }
  })
  return best.score > 0 ? best.index : -1
}

function scoreSheet(sheet) {
  if (sheet.headerIndex < 0 || sheet.timeColumn < 0) return 0
  const dataRows = sheet.dataRows.length
  const mapped = Object.values(mapChannels(sheet.headers)).filter((index) => index >= 0).length
  const name = sheet.name.toLowerCase()
  const sourceBonus = name.includes('quickconnection') ? 2500 : name === 'summary' ? 500 : 0
  return dataRows + mapped * 500 + sourceBonus
}

function detectHeaderIndex(rows) {
  let bestIndex = -1
  let bestScore = 0

  rows.slice(0, 50).forEach((row, index) => {
    const cells = row.map((cell) => String(cell || '').trim()).filter(Boolean)
    const timeScore = cells.some(isTimeHeader) ? 6 : 0
    const signalScore = cells.filter((cell) => Object.values(SIGNALS).some((signal) => signal.aliases.some((alias) => cell.toLowerCase().includes(alias)))).length
    const widthScore = Math.min(cells.length, 8)
    const score = timeScore + signalScore * 2 + widthScore
    if (score > bestScore) {
      bestScore = score
      bestIndex = index
    }
  })

  return bestScore >= 8 ? bestIndex : -1
}

function parseDelimitedText(text) {
  const lines = String(text || '').replace(/^\uFEFF/, '').split(/\r?\n/).filter((line) => line.trim())
  const delimiters = ['\t', ',', ';', '|']
  const delimiter = delimiters
    .map((candidate) => ({ candidate, count: lines.slice(0, 20).reduce((sum, line) => sum + line.split(candidate).length, 0) }))
    .sort((a, b) => b.count - a.count)[0]?.candidate || ','
  return lines.map((line) => splitDelimitedLine(line, delimiter))
}

function splitDelimitedLine(line, delimiter) {
  if (delimiter !== ',') return line.split(delimiter).map(cleanCell)
  const cells = []
  let current = ''
  let quoted = false
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]
    if (char === '"') {
      quoted = !quoted
    } else if (char === ',' && !quoted) {
      cells.push(cleanCell(current))
      current = ''
    } else {
      current += char
    }
  }
  cells.push(cleanCell(current))
  return cells
}

function isTimeHeader(value) {
  const normalized = String(value || '').toLowerCase()
  return normalized.includes('time') || normalized.includes('date') || normalized.includes('timestamp')
}

function parseDate(value, previousTimestamp = null) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  const text = String(value || '').trim()
  if (!text) return null

  const timeOnly = text.match(/^(\d{1,2}):(\d{2}):(\d{2})[\s:.](\d{1,3})$/)
  if (timeOnly) {
    const [, hour, minute, second, ms = '0'] = timeOnly
    const date = new Date(2000, 0, 1, Number(hour), Number(minute), Number(second), Number(ms.padEnd(3, '0')))
    if (previousTimestamp && date.getTime() < previousTimestamp - 12 * 60 * 60 * 1000) {
      date.setDate(date.getDate() + 1)
    }
    return date
  }

  let match = text.match(DATE_PATTERNS[0])
  if (match) {
    const [, day, month, year, hour, minute, second, ms = '0'] = match
    return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second), Number(ms.padEnd(3, '0')))
  }

  match = text.match(DATE_PATTERNS[1])
  if (match) {
    const [, year, month, day, hour, minute, second, ms = '0'] = match
    return new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second), Number(ms.padEnd(3, '0')))
  }

  match = text.match(DATE_PATTERNS[2])
  if (match) {
    const [, month, day, rawYear, hour, minute, second, ms = '0'] = match
    const year = rawYear.length === 2 ? Number(`20${rawYear}`) : Number(rawYear)
    return new Date(year, Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second), Number(ms.padEnd(3, '0')))
  }

  const date = new Date(text)
  return Number.isNaN(date.getTime()) ? null : date
}

function cleanNumeric(signal, value) {
  if (value === null) return null
  const config = SIGNALS[signal]
  if (!config) return value
  if (config.min !== null && config.min !== undefined && value < config.min) return null
  if (config.max !== null && config.max !== undefined && value > config.max) return null
  if (Math.abs(value) > 900000000) return null
  return value
}

function applyScale(signal, value) {
  if (value === null) return null
  const scale = SIGNALS[signal]?.scale
  return typeof scale === 'function' ? scale(value) : value
}

function parseNumeric(value) {
  if (value === null || value === undefined || value === '') return null
  if (typeof value === 'number' && Number.isFinite(value)) return value
  const normalized = String(value).replace(/,/g, '').replace(/[^\d.+\-eE]/g, '')
  if (!normalized || normalized === '-' || normalized === '.') return null
  const number = Number(normalized)
  return Number.isFinite(number) ? number : null
}

function normalizeHeader(value, index) {
  const text = String(value || '').replace(/\s+/g, ' ').trim()
  return text || `Column ${index + 1}`
}

function cleanCell(value) {
  if (value === undefined || value === null) return null
  if (value instanceof Date) return value
  const text = String(value).replace(/\uFEFF/g, '').trim().replace(/^"|"$/g, '').trim()
  return text === '' ? null : text
}

function makeEvent(severity, title, point, detail) {
  return {
    severity,
    title,
    detail,
    rowNumber: point.rowNumber,
    time: point.isoTime,
  }
}

function downsample(series, maxPoints) {
  if (series.length <= maxPoints) return series
  const step = Math.ceil(series.length / maxPoints)
  const sampled = []
  for (let index = 0; index < series.length; index += step) sampled.push(series[index])
  if (sampled[sampled.length - 1] !== series[series.length - 1]) sampled.push(series[series.length - 1])
  return sampled
}

function average(values) {
  if (!values.length) return null
  return values.reduce((sum, value) => sum + value, 0) / values.length
}

function imbalancePercent(values) {
  if (values.length < 2) return null
  const avg = average(values)
  if (!avg) return null
  return ((Math.max(...values) - Math.min(...values)) / avg) * 100
}

function maxOf(series, key) {
  const values = series.map((point) => point[key]).filter(isFiniteNumber)
  return values.length ? Math.max(...values) : null
}

function minOf(series, key) {
  const values = series.map((point) => point[key]).filter(isFiniteNumber)
  return values.length ? Math.min(...values) : null
}

function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value)
}

function fmt(value) {
  return value === null || value === undefined ? 'n/a' : Number(value).toFixed(2)
}
