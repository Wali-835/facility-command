import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App.tsx'
import AssetPage from './AssetPage.jsx'

const isAssetPage = new URLSearchParams(window.location.search).get("asset");

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    {isAssetPage ? <AssetPage /> : <App />}
  </StrictMode>,
)
