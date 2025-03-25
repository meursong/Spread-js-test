<template>
  <div>
    <h1 class="text-2xl font-bold mb-6">빈 엑셀 시트</h1>
    <ClientOnly>
      <Suspense>
        <template #default>
          <EmptySheet @sheet-loaded="onSheetLoaded" />
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
import Empty from "~/pages/spreadsheet/empty.vue";

const { stopLoading } = inject('loading')

const onSheetLoaded = () => {
  stopLoading()
}

definePageMeta({
  layout: 'default'
})
</script>