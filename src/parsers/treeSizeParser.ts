import Papa from 'papaparse'
import { Workbook, type Worksheet } from 'exceljs'
import { detectFormat } from './fileDetector'
import type { TreeNode, ParsedTreeSizeRow } from '../types'

// ─── Public API ───────────────────────────────────────────────────────────────

export async function parseTreeSizeFile(file: File): Promise<TreeNode> {
  const format = detectFormat(file)
  let rows: ParsedTreeSizeRow[]

  if (format === 'csv') {
    rows = await parseCsv(file)
  } else if (format === 'xlsx') {
    rows = await parseExcel(file)
  } else {
    throw new Error(`Unsupported file type: ${file.name}. Please upload a .csv or .xlsx file.`)
  }

  if (rows.length === 0) {
    throw new Error('The file contains no data rows. Please check the export format.')
  }

  return buildTree(rows)
}

// ─── CSV parser ───────────────────────────────────────────────────────────────

async function parseCsv(file: File): Promise<ParsedTreeSizeRow[]> {
  return new Promise((resolve, reject) => {
    Papa.parse<Record<string, string>>(file, {
      header: true,
      skipEmptyLines: true,
      complete: (result) => {
        try {
          validateHeaders(result.meta.fields ?? [])
          resolve(result.data.map(normalizeRow))
        } catch (err) {
          reject(err)
        }
      },
      error: (err: Error) => reject(err),
    })
  })
}

// ─── Excel parser ─────────────────────────────────────────────────────────────

async function parseExcel(file: File): Promise<ParsedTreeSizeRow[]> {
  const buffer = await file.arrayBuffer()
  const workbook = new Workbook()
  await workbook.xlsx.load(buffer)

  // Use the first worksheet that has data; TreeSize sometimes adds a cover sheet
  const sheet = workbook.worksheets.find((ws) => ws.rowCount > 1) ?? workbook.worksheets[0]
  if (!sheet) throw new Error('No worksheets found in the Excel file.')

  // TreeSize reports sometimes prepend metadata rows before the column headers.
  // Scan the first 10 rows to find the actual header row.
  const headerRowNum = findHeaderRowNum(sheet)
  if (headerRowNum === -1) {
    throw new Error(
      `Could not find a header row. Expected a column named one of: ${PATH_COLS.join(', ')}.`
    )
  }

  const headers: string[] = []
  sheet.getRow(headerRowNum).eachCell({ includeEmpty: true }, (cell, colNum) => {
    headers[colNum - 1] = String(cell.value ?? '').trim()
  })

  const rows: ParsedTreeSizeRow[] = []
  sheet.eachRow((row, rowNum) => {
    if (rowNum <= headerRowNum) return
    const record: Record<string, string> = {}
    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const header = headers[colNum - 1]
      if (header) record[header] = cellText(cell.value)
    })
    const parsed = normalizeRow(record)
    if (parsed.path) rows.push(parsed)
  })

  return rows
}

/** Scan up to the first 10 rows for the row containing a known path column name. */
function findHeaderRowNum(sheet: Worksheet): number {
  for (let r = 1; r <= Math.min(10, sheet.rowCount); r++) {
    const row = sheet.getRow(r)
    let found = false
    row.eachCell((cell) => {
      const v = String(cell.value ?? '').trim()
      if (PATH_COLS.includes(v)) found = true
    })
    if (found) return r
  }
  return -1
}

function cellText(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') {
    const v = value as Record<string, unknown>
    // Formula cell — use cached result
    if ('result' in v) return String(v.result ?? '').trim()
    // Rich text cell (ExcelJS returns { richText: [{text, font}, ...] })
    if ('richText' in v && Array.isArray(v.richText)) {
      return (v.richText as Array<{ text?: unknown }>).map((r) => String(r.text ?? '')).join('').trim()
    }
    // Hyperlink cell
    if ('text' in v) return String(v.text ?? '').trim()
    // Error cell
    if ('error' in v) return ''
  }
  return String(value).trim()
}

// ─── Normalisation ────────────────────────────────────────────────────────────

// TreeSize exports can use different column name conventions across versions.
// We try several common variants.
const PATH_COLS = ['Full Path', 'Path', 'Folder', 'Directory', 'Name', 'Share', 'Drive', 'Shared Drive']
const SIZE_COLS = ['Size', 'Size (Bytes)', 'Allocated', 'Size in Bytes', 'Total Size', 'Used Space']
const FILES_COLS = ['Files', '# Files', 'File Count', 'Number of Files', 'Total Files']
const FOLDERS_COLS = ['Folders', '# Folders', 'Subfolder Count', 'Subfolders', 'Total Folders']
const DATE_COLS = ['Last Change', 'Last Modified', 'Modified', 'Date Modified', 'Last Accessed']

function findCol(record: Record<string, string>, candidates: string[]): string {
  for (const c of candidates) {
    if (record[c] !== undefined) return record[c]
  }
  return ''
}

function parseBytes(raw: string): number {
  if (!raw) return 0
  // Strip non-numeric except decimal point
  const cleaned = raw.replace(/[^0-9.]/g, '')
  const num = parseFloat(cleaned)
  if (isNaN(num)) return 0

  const lower = raw.toLowerCase()
  if (lower.includes('tb')) return Math.round(num * 1e12)
  if (lower.includes('gb')) return Math.round(num * 1e9)
  if (lower.includes('mb')) return Math.round(num * 1e6)
  if (lower.includes('kb')) return Math.round(num * 1e3)
  return Math.round(num)
}

function normalizeRow(record: Record<string, string>): ParsedTreeSizeRow {
  const rawDate = findCol(record, DATE_COLS)
  const lastModified = rawDate ? new Date(rawDate) : undefined
  return {
    path: findCol(record, PATH_COLS).replace(/\\/g, '/').replace(/\/+$/, ''),
    sizeBytes: parseBytes(findCol(record, SIZE_COLS)),
    fileCount: parseInt(findCol(record, FILES_COLS).replace(/[^0-9]/g, '') || '0', 10),
    folderCount: parseInt(findCol(record, FOLDERS_COLS).replace(/[^0-9]/g, '') || '0', 10),
    lastModified: lastModified && !isNaN(lastModified.getTime()) ? lastModified : undefined,
  }
}

function validateHeaders(headers: string[]): void {
  const hasPath = PATH_COLS.some((c) => headers.includes(c))
  if (!hasPath) {
    throw new Error(
      `Could not find a path column. Expected one of: ${PATH_COLS.join(', ')}. ` +
      `Found: ${headers.join(', ')}`
    )
  }
}

// ─── Tree builder ─────────────────────────────────────────────────────────────

function buildTree(rows: ParsedTreeSizeRow[]): TreeNode {
  // Sort by path length so parents always come before children
  const sorted = [...rows].sort((a, b) => a.path.localeCompare(b.path))

  const nodeMap = new Map<string, TreeNode>()

  for (const row of sorted) {
    const parts = row.path.split('/').filter(Boolean)
    const name = parts[parts.length - 1] ?? row.path
    const depth = parts.length - 1

    const node: TreeNode = {
      path: row.path,
      name,
      depth,
      sizeBytes: row.sizeBytes,
      fileCount: row.fileCount,
      folderCount: row.folderCount,
      lastModified: row.lastModified,
      children: [],
    }

    nodeMap.set(row.path, node)

    // Find parent
    if (parts.length > 1) {
      const parentPath = parts.slice(0, -1).join('/')
      const parent = nodeMap.get(parentPath) ?? nodeMap.get('/' + parentPath)
      if (parent) {
        parent.children.push(node)
      }
    }
  }

  // Find root(s) — nodes with no parent in the map
  const roots = sorted.filter((r) => {
    const parts = r.path.split('/').filter(Boolean)
    if (parts.length <= 1) return true
    const parentPath = parts.slice(0, -1).join('/')
    return !nodeMap.has(parentPath) && !nodeMap.has('/' + parentPath)
  })

  if (roots.length === 1) {
    return nodeMap.get(roots[0].path)!
  }

  // Multiple roots → synthetic root
  const syntheticRoot: TreeNode = {
    path: '',
    name: 'Root',
    depth: -1,
    sizeBytes: roots.reduce((s, r) => s + (nodeMap.get(r.path)?.sizeBytes ?? 0), 0),
    fileCount: roots.reduce((s, r) => s + (nodeMap.get(r.path)?.fileCount ?? 0), 0),
    folderCount: roots.length,
    children: roots.map((r) => nodeMap.get(r.path)!),
  }
  return syntheticRoot
}
