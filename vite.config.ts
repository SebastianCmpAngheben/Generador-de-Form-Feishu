import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // Asegura que las rutas a los scripts en el HTML sean relativas
  build: {
    outDir: 'dist', // Fuerza a que la carpeta de salida se llame estrictamente 'dist'
    emptyOutDir: true, // Limpia la carpeta dist antes de cada nueva compilación
    rollupOptions: {
      input: {
        main: './index.html', // Establece el index.html de la raíz como el punto de entrada principal
      },
      output: {
        // Evita que los archivos JS y CSS se metan en subcarpetas profundas, manteniéndolos ordenados en dist
        assetFileNames: 'assets/[name]-[hash][extname]',
        chunkFileNames: 'assets/[name]-[hash].js',
        entryFileNames: 'assets/[name]-[hash].js',
      },
    },
  },
})