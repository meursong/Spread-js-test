<template>
  <div>
    <h1 class="text-2xl font-bold mb-6">데이터 바인딩 테스트</h1>
    <ClientOnly>
      <Suspense>
        <template #default>
          <DataBindingSheet @sheet-loaded="onSheetLoaded" />
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