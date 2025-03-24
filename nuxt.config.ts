export default defineNuxtConfig({
  ssr: true,
  vite: {
    build: {
      commonjsOptions: {
        transformMixedEsModules: true,
        include: [
          /node_modules\/@mescius\/spread-sheets-vue/,
          /node_modules\/@mescius\/spread-sheets-resources-ko/
        ]
      }
    }
  },
  compatibilityDate: '2025-03-24'
})