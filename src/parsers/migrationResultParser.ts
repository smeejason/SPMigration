import JSZip from 'jszip'
import Papa from 'papaparse'
import type { MigrationResultItem, MigrationResultStatus, MigrationResultSummary, ReviewNode } from '../types'

// ─── Public API ───────────────────────────────────────────────────────────────

export async function parseMigrationResultZip(file: File): Promise<MigrationResultSummary> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer())

  // Find ItemReport (not failure) and ItemFailureReport
  let itemReportText = ''
  let failureReportText = ''

  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    const base = name.split('/').pop() ?? name
    if (/ItemFailureReport/i.test(base)) {
      failureReportText = await entry.async('text')
    } else if (/ItemReport/i.test(base)) {
      itemReportText = await entry.async('text')
    }
  }

  if (!itemReportText) throw new Error('No ItemReport CSV found in ZIP')

  // Build error code map from failure report
  const failureErrorCodes = new Map<string, string>()
  if (failureReportText) {
    const failRows = parseCsvText(stripBom(failureReportText))
    for (const row of failRows) {
      const src = row['Source'] ?? ''
      const code = row['Error code'] ?? ''
      if (src && code) failureErrorCodes.set(src, code)
    }
  }

  // Parse item report
  const itemRows = parseCsvText(stripBom(itemReportText))
  if (itemRows.length === 0) throw new Error('ItemReport CSV has no data rows')

  // Deduplicate by Source — keep highest-priority status across incremental rounds
  const itemMap = new Map<string, MigrationResultItem>()
  for (const row of itemRows) {
    const source = row['Source'] ?? ''
    if (!source) continue

    const status = normalizeStatus(row['Status'] ?? '')
    const item: MigrationResultItem = {
      source,
      destination: row['Destination'] ?? '',
      itemName: row['Item name'] ?? '',
      itemType: (row['Type'] ?? '').toLowerCase() === 'folder' ? 'Folder' : 'File',
      status,
      resultCategory: row['Result category'] ?? '',
      message: row['Message'] ?? '',
      errorCode: failureErrorCodes.get(source) ?? '',
      fileSizeBytes: parseInt((row['Item size (bytes)'] ?? '0').replace(/[^0-9]/g, '') || '0', 10),
      isRecycleBin: source.includes('$RECYCLE.BIN'),
      sourcePath: normalizePath(source),
    }

    const existing = itemMap.get(source)
    if (!existing || statusPriority(status) > statusPriority(existing.status)) {
      itemMap.set(source, item)
    }
  }

  const items = Array.from(itemMap.values())

  let migratedCount = 0, failedCount = 0, skippedCount = 0, partialCount = 0
  for (const item of items) {
    if (item.status === 'Migrated') migratedCount++
    else if (item.status === 'Failed') failedCount++
    else if (item.status === 'Skipped') skippedCount++
    else if (item.status === 'Partial') partialCount++
  }

  return { items, migratedCount, failedCount, skippedCount, partialCount, totalCount: items.length }
}

// ─── Tree builder ─────────────────────────────────────────────────────────────

export function buildReviewTree(items: MigrationResultItem[]): ReviewNode {
  const nodeMap = new Map<string, ReviewNode>()

  function getOrCreate(path: string, name: string, depth: number): ReviewNode {
    if (nodeMap.has(path)) return nodeMap.get(path)!
    const node: ReviewNode = { path, name, depth, children: [], migratedCount: 0, failedCount: 0, skippedCount: 0, partialCount: 0, totalCount: 0 }
    nodeMap.set(path, node)
    return node
  }

  // Pass 1: create all nodes from item paths
  for (const item of items) {
    const parts = item.sourcePath.split('/').filter(Boolean)
    for (let i = 0; i < parts.length; i++) {
      const path = parts.slice(0, i + 1).join('/')
      getOrCreate(path, parts[i], i)
    }
  }

  // Pass 2: increment leaf counts
  for (const item of items) {
    const node = nodeMap.get(item.sourcePath)
    if (!node) continue
    node.totalCount++
    if (item.status === 'Migrated') node.migratedCount++
    else if (item.status === 'Failed') node.failedCount++
    else if (item.status === 'Skipped') node.skippedCount++
    else if (item.status === 'Partial') node.partialCount++
  }

  // Pass 3: link children to parents, bubble counts up
  const roots: ReviewNode[] = []
  for (const node of nodeMap.values()) {
    const parts = node.path.split('/')
    if (parts.length === 1) {
      roots.push(node)
    } else {
      const parentPath = parts.slice(0, -1).join('/')
      const parent = nodeMap.get(parentPath)
      if (parent) {
        parent.children.push(node)
      } else {
        roots.push(node)
      }
    }
  }

  // Bubble counts from leaves up — process in deepest-first order
  const sorted = [...nodeMap.values()].sort((a, b) => b.depth - a.depth)
  for (const node of sorted) {
    const parts = node.path.split('/')
    if (parts.length > 1) {
      const parentPath = parts.slice(0, -1).join('/')
      const parent = nodeMap.get(parentPath)
      if (parent) {
        parent.migratedCount += node.migratedCount
        parent.failedCount += node.failedCount
        parent.skippedCount += node.skippedCount
        parent.partialCount += node.partialCount
        parent.totalCount += node.totalCount
      }
    }
  }

  // Sort children: folders first, then files, both alphabetically
  const folderFirst = (a: ReviewNode, b: ReviewNode): number => {
    const aIsFolder = a.children.length > 0 ? 0 : 1
    const bIsFolder = b.children.length > 0 ? 0 : 1
    if (aIsFolder !== bIsFolder) return aIsFolder - bIsFolder
    return a.name.localeCompare(b.name)
  }
  for (const node of nodeMap.values()) {
    node.children.sort(folderFirst)
  }

  roots.sort(folderFirst)

  // Return synthetic root if multiple roots, or single root
  if (roots.length === 1) return roots[0]

  return {
    path: '',
    name: '',
    depth: -1,
    children: roots,
    migratedCount: roots.reduce((s, n) => s + n.migratedCount, 0),
    failedCount: roots.reduce((s, n) => s + n.failedCount, 0),
    skippedCount: roots.reduce((s, n) => s + n.skippedCount, 0),
    partialCount: roots.reduce((s, n) => s + n.partialCount, 0),
    totalCount: roots.reduce((s, n) => s + n.totalCount, 0),
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function parseCsvText(text: string): Record<string, string>[] {
  const result = Papa.parse<Record<string, string>>(text, {
    header: true,
    skipEmptyLines: true,
  })
  return result.data
}

function stripBom(text: string): string {
  return text.startsWith('\uFEFF') ? text.slice(1) : text
}

function normalizeStatus(raw: string): MigrationResultStatus {
  const s = raw.trim().toLowerCase()
  if (s === 'migrated') return 'Migrated'
  if (s === 'failed') return 'Failed'
  if (s === 'skipped') return 'Skipped'
  if (s === 'partial success' || s === 'partial') return 'Partial'
  return 'Skipped'
}

function statusPriority(s: MigrationResultStatus): number {
  return s === 'Migrated' ? 4 : s === 'Partial' ? 3 : s === 'Failed' ? 2 : 1
}

function normalizePath(raw: string): string {
  return raw
    .replace(/^\\\\/, '')   // strip leading \\
    .replace(/\\/g, '/')    // backslash → forward slash
    .replace(/\/+/g, '/')   // collapse consecutive slashes
    .replace(/^\//, '')     // strip leading /
    .replace(/\/$/, '')     // strip trailing /
    .trim()
}
