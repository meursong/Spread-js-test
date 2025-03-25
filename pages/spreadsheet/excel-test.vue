<template>
  <div>
    <h1 class="text-2xl font-bold mb-6">엑셀 파일 가져오기/내보내기</h1>
    <ClientOnly>
      <Suspense>
        <template #default>
          <ExcelTestSheet @sheet-loaded="onSheetLoaded" />
        </template>
        <template #fallback>
          <LoadingSpinner>
            스프레드시트를 불러오는 중...
          </LoadingSpinner>
        </template>
      </Suspense>
    </ClientOnly>
  </div>
</template>

<script setup>
const { stopLoading } = inject('loading')

const onSheetLoaded = () => {
  stopLoading()
}

definePageMeta({
  layout: 'default'
})
</script>