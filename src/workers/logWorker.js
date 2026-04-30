import { parseLogFile } from '../lib/parser.js'

self.onmessage = async (event) => {
  const { fileName, buffer, text } = event.data
  try {
    const result = parseLogFile({ fileName, buffer, text })
    self.postMessage({ type: 'success', result })
  } catch (error) {
    self.postMessage({
      type: 'error',
      message: error instanceof Error ? error.message : 'Unable to parse log file.',
    })
  }
}
