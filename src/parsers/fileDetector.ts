export type SupportedFormat = 'csv' | 'xlsx' | 'unknown'

export function detectFormat(file: File): SupportedFormat {
  const name = file.name.toLowerCase()
  if (name.endsWith('.csv')) return 'csv'
  if (name.endsWith('.xlsx') || name.endsWith('.xls')) return 'xlsx'
  // Fallback: sniff MIME type
  if (file.type === 'text/csv' || file.type === 'text/plain') return 'csv'
  if (
    file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    file.type === 'application/vnd.ms-excel'
  ) {
    return 'xlsx'
  }
  return 'unknown'
}
