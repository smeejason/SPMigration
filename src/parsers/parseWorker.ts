import { parseTreeSizeFile } from './treeSizeParser'

// Web Worker — runs parsing off the main thread so the UI stays responsive
self.onmessage = async (e: MessageEvent<File>) => {
  try {
    const tree = await parseTreeSizeFile(e.data)
    self.postMessage({ ok: true, tree })
  } catch (err) {
    self.postMessage({ ok: false, error: (err as Error).message })
  }
}
