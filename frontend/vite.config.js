import { defineConfig } from 'vite'

export default defineConfig({
  base: './',
  server: {
    port: 5174,
    host: true
  },
  build: {
    outDir: 'dist',
    assetsDir: 'assets'
  }
})
