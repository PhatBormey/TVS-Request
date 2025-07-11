import path from 'path';
import { defineConfig, loadEnv } from 'vite';

export default defineConfig(({ mode }) => {
    const env = loadEnv(mode, '.', '');
    return {
      define: {
        'process.env.API_KEY': JSON.stringify(env.GEMINI_API_KEY),
        'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY)
      },
      resolve: {
        alias: {
          '@': path.resolve(__dirname, '.'),
        }
      },
      server: {
      host: '0.0.0.0', // 👈 this allows external access (like Render)
      port: 5173,
      allowedHosts: ['tvs-request.onrender.com'], // ✅ Add this line       // 👈 optional, or use process.env.PORT
    }
    };
});
