import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// Served from the root of the custom domain https://entra.2cloud.nl/ on GitHub
// Pages. Set VITE_BASE (e.g. /EntraApp/) to host it under a sub-path instead,
// such as a fork published at https://<owner>.github.io/<repo>/.
const base = process.env.VITE_BASE ?? '/';

export default defineConfig({
  base,
  plugins: [react()],
  build: {
    outDir: 'dist',
    sourcemap: false,
  },
});
