import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App.jsx'
import './index.css'

// Suppress known harmless non-fatal CFB/ZIP streaming warnings from xlsx libraries (e.g. "Bad uncompressed size")
if (typeof window !== 'undefined') {
  const originalConsoleError = console.error
  console.error = function (...args) {
    const msg = typeof args[0] === 'string' ? args[0] : ''
    if (
      msg.includes('Bad uncompressed size:') ||
      msg.includes('Bad compressed size:') ||
      msg.includes('Bad CRC32 checksum:')
    ) {
      // Demote CFB library streaming notices to debug so they don't trigger false-positive error alerts
      if (console.debug) console.debug('[XLSX Notice]', ...args)
      return
    }
    originalConsoleError.apply(console, args)
  }
}

ReactDOM.createRoot(document.getElementById('root')).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)

