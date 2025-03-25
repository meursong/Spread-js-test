<template>
  <div>
    <h1 class="text-2xl font-bold mb-6">Excel 기능 개발</h1>
    <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
      <NuxtLink to="/pages/spreadsheet/data-types"
                class="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow">
        <h3 class="text-xl font-semibold mb-2">데이터 타입 테스트</h3>
      </NuxtLink>

      <NuxtLink to="/pages/spreadsheet/styles"
                class="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow">
        <h3 class="text-xl font-semibold mb-2">스타일링 테스트</h3>
      </NuxtLink>

      <!-- 추가 카드들 -->
    </div>
  </div>
</template>