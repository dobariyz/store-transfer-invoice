import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  preview: {
    allowedHosts: ["store-transfer-invoice-production.up.railway.app"]
  },
  server: {
    allowedHosts: ["store-transfer-invoice-production.up.railway.app"]
  }
})
