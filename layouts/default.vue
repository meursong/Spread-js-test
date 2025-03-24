<template>
  <div class="app-container">
    <!-- 사이드 네비게이션 -->
    <nav class="navigation">
      <div class="nav-header">
        <h1 class="nav-title">SpreadJS 테스트</h1>
      </div>
      <div class="nav-links">
        <NuxtLink to="/" class="nav-item">
          <span class="nav-text">홈</span>
        </NuxtLink>
        <NuxtLink to="/spreadsheet/excel-test" class="nav-item">
          <span class="nav-text">엑셀 테스트</span>
        </NuxtLink>
        <NuxtLink to="/spreadsheet/data-types" class="nav-item">
          <span class="nav-text">데이터 타입</span>
        </NuxtLink>
        <NuxtLink to="/spreadsheet/data-binding" class="nav-item">
          <span class="nav-text">데이터 바인딩</span>
        </NuxtLink>
        <NuxtLink to="/spreadsheet/formulas" class="nav-item">
          <span class="nav-text">수식</span>
        </NuxtLink>
      </div>
    </nav>

    <!-- 메인 컨텐츠 영역 -->
    <client-only>
      <main class="content">
        <slot />
      </main>
    </client-only>
  </div>
</template>

<style scoped>
.app-container {
  display: flex;
  min-height: 100vh;
}

.navigation {
  width: 240px;
  background-color: #f5f5f5;
  border-right: 1px solid #e0e0e0;
  display: flex;
  flex-direction: column;
}

.nav-header {
  padding: 20px;
  border-bottom: 1px solid #e0e0e0;
}

.nav-title {
  font-size: 1.25rem;
  font-weight: 600;
  color: #333;
  margin: 0;
}

.nav-links {
  padding: 20px 10px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}

.nav-item {
  padding: 12px 16px;
  text-decoration: none;
  color: #333;
  border-radius: 8px;
  transition: all 0.2s ease;
  font-size: 0.95rem;
}

.nav-item:hover {
  background-color: #e0e0e0;
}

.nav-text {
  font-weight: 500;
}

.content {
  flex: 1;
  padding: 24px;
  background-color: #fff;
}

/* Nuxt의 active 클래스에 대한 스타일 */
.router-link-active {
  background-color: #4472C4;
  color: white;
}

/* 반응형 디자인 */
@media (max-width: 768px) {
  .navigation {
    width: 200px;
  }
}

@media (max-width: 640px) {
  .app-container {
    flex-direction: column;
  }

  .navigation {
    width: 100%;
    border-right: none;
    border-bottom: 1px solid #e0e0e0;
  }

  .nav-links {
    padding: 10px;
  }
}
</style>
<script setup lang="ts">
</script>