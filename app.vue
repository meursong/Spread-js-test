<template>
  <LoadingSpinner v-if="loading" />
  <NuxtLayout>
    <NuxtPage />
  </NuxtLayout>
</template>

<script setup>
import {useLoading} from "~/composables/useLoading.js";
const router = useRouter();

const { loading, startLoading, stopLoading } = useLoading()

// provide를 사용하여 하위 컴포넌트에서도 접근 가능하도록 설정
provide('loading', {
  loading,
  startLoading,
  stopLoading
})

router.beforeEach((to, from, next) => {
  startLoading()
  next()
})

router.afterEach(() => {
  // 컴포넌트가 마운트될 때까지 약간의 지연
  setTimeout(() => {
    stopLoading()
  }, 100)
})

</script>

<style>
/* 전역 스타일 */
body {
  margin: 0;
  padding: 0;
  font-family: 'Pretendard', sans-serif;
}
</style>