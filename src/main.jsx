import React from 'react'
import ReactDOM from 'react-dom/client'
import AppShell from './AppShell.jsx'
import ErrorBoundary from './ErrorBoundary.jsx'
ReactDOM.createRoot(document.getElementById('root')).render(<ErrorBoundary><AppShell /></ErrorBoundary>)
