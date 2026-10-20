import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          // React core separado — cached pelo browser entre deploys
          'vendor-react': ['react', 'react-dom'],
          // Supabase client separado
          'vendor-supabase': ['@supabase/supabase-js'],
        },
      },
    },
    // Aumenta o limite para não spammar warnings, mas mantemos o split
    chunkSizeWarningLimit: 600,
  },
})
