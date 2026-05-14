import type { QuickDocumentApi } from '../../preload'

declare global {
  interface Window {
    quickDocument: QuickDocumentApi
  }
}

export {}
