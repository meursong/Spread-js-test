<template>
  <div>
    <h1 class="text-2xl font-bold mb-6">엑셀 수식 테스트</h1>
    <ClientOnly>
      <Suspense>
        <template #default>
          <FomulaSheet @sheet-loaded="onSheetLoaded" />
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
// loading composable 주입
const { stopLoading } = inject('loading')

const onSheetLoaded = () => {
  stopLoading()
}

definePageMeta({
  layout: 'default'
})
</script>