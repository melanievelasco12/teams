import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, '.', 'VITE_')
  const backendUrl = env.VITE_LTI_BACKEND_URL

  return {
    plugins: [react()],
    build: {
      outDir: 'server/public',
      emptyOutDir: false,
    },
    server: {
      allowedHosts: ['spousal-engraved-ability.ngrok-free.dev'],
      proxy: backendUrl
        ? {
            '/api': backendUrl,
            '/lti': backendUrl,
          }
        : undefined,
    },
  }
})
