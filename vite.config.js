import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: './',
  optimizeDeps: {
    include: ['xlsx-js-style'],
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks: {
          'vendor-react': ['react', 'react-dom'],
          'vendor-xlsx': ['xlsx', 'xlsx-js-style'],
          'vendor-neon': ['@neondatabase/postgrest-js'],
        }
      }
    },
    chunkSizeWarningLimit: 800,
  }
})
