import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          // React core — sempre cacheado entre deploys
          if (id.includes('node_modules/react') || id.includes('node_modules/react-dom') || id.includes('node_modules/scheduler')) {
            return 'vendor-react';
          }
          // Supabase client — raramente muda
          if (id.includes('node_modules/@supabase')) {
            return 'vendor-supabase';
          }
          // Recharts + D3 — carregados sob demanda pela aba Faturamento
          if (id.includes('node_modules/recharts') || id.includes('node_modules/d3') || id.includes('node_modules/victory') || id.includes('node_modules/internmap') || id.includes('node_modules/robust-predicates')) {
            return 'vendor-recharts';
          }
          // XLSX — carregado sob demanda no export
          if (id.includes('node_modules/xlsx')) {
            return 'vendor-xlsx';
          }
        },
      },
    },
    chunkSizeWarningLimit: 600,
    // Minificação mais agressiva
    minify: 'terser',
    terserOptions: {
      compress: {
        drop_console: true,
        passes: 2,
      },
    },
  },
})
