import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  build: {
    minify: 'terser',
    terserOptions: {
      compress: {
        // Desabilita otimizações que reordenam declarações e causam TDZ
        sequences: false,
        join_vars: false,
        collapse_vars: false,
        reduce_vars: false,
        hoist_vars: false,
        hoist_funs: false,
      },
      mangle: true,
    },
  },
})
