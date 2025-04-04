<template>
  <div class="excel-ribbon">
    <!-- 리본 탭 영역 -->
    <div class="ribbon-tabs">
      <div class="ribbon-tab active">삽입</div>
      <div class="ribbon-tab">페이지 레이아웃</div>
      <div class="ribbon-tab">수식</div>
      <div class="ribbon-tab">데이터</div>
      <div class="ribbon-tab">검토</div>
      <div class="ribbon-tab">보기</div>
    </div>

    <!-- 리본 컨텐츠 영역 (홈 탭) -->
    <div class="ribbon-content">
      <!-- 클립보드 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">클립보드</div>
        <div class="ribbon-controls">
          <div class="ribbon-button-group">
            <div class="ribbon-button small">
              <div class="icon">📋</div>
              <div class="label">붙여넣기</div>
            </div>
            <div class="ribbon-button small">
              <div class="icon">✂️</div>
              <div class="label">잘라내기</div>
            </div>
          </div>
          <div class="ribbon-button-group">
            <div class="ribbon-button small">
              <div class="icon">📄</div>
              <div class="label">복사</div>
            </div>
            <div class="ribbon-button small">
              <div class="icon">🧹</div>
              <div class="label">서식 복사</div>
            </div>
          </div>
        </div>
      </div>

      <!-- 글꼴 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">글꼴</div>
        <div class="ribbon-controls">
          <div class="font-controls">
            <select class="font-dropdown">
              <option>맑은 고딕</option>
              <option>Arial</option>
              <option>Calibri</option>
              <option>Times New Roman</option>
            </select>
            <select class="size-dropdown">
              <option>11</option>
              <option>12</option>
              <option>14</option>
              <option>16</option>
              <option>18</option>
            </select>
          </div>
          <div class="format-controls">
            <button class="format-button" title="굵게">B</button>
            <button class="format-button" title="기울임꼴">I</button>
            <button class="format-button" title="밑줄">U</button>
            <button class="format-button" title="취소선">S</button>
            <div class="color-picker" title="글꼴 색">A</div>
            <div class="color-picker" title="배경색">■</div>
            <button class="format-button" title="테두리">□</button>
          </div>
        </div>
      </div>

      <!-- 정렬 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">정렬</div>
        <div class="ribbon-controls">
          <div class="ribbon-button-group">
            <button class="format-button" title="위쪽 맞춤">▲</button>
            <button class="format-button" title="가운데 맞춤">■</button>
            <button class="format-button" title="아래쪽 맞춤">▼</button>
          </div>
          <div class="ribbon-button-group">
            <button class="format-button" title="왼쪽 맞춤">◀</button>
            <button class="format-button" title="가운데 맞춤">■</button>
            <button class="format-button" title="오른쪽 맞춤">▶</button>
          </div>
          <div class="ribbon-button-group">
            <button class="format-button" title="텍스트 내림">↓</button>
            <button class="format-button" title="병합하고 가운데 맞춤">⊞</button>
          </div>
        </div>
      </div>

      <!-- 숫자 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">숫자</div>
        <div class="ribbon-controls">
          <select class="format-dropdown">
            <option>일반</option>
            <option>숫자</option>
            <option>날짜</option>
            <option>통화</option>
            <option>백분율</option>
          </select>
          <div class="ribbon-button-group">
            <button class="format-button" title="통화 서식">₩</button>
            <button class="format-button" title="백분율 서식">%</button>
            <button class="format-button" title="천 단위 구분 기호">1,000</button>
          </div>
          <div class="ribbon-button-group">
            <button class="format-button" title="소수 자릿수 늘림">.0→.00</button>
            <button class="format-button" title="소수 자릿수 줄임">.00→.0</button>
          </div>
        </div>
      </div>

      <!-- 스타일 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">스타일</div>
        <div class="ribbon-controls">
          <div class="ribbon-button-group">
            <button class="format-button" title="조건부 서식">조건부 서식</button>
            <button class="format-button" title="표 서식">표 서식</button>
            <button class="format-button" title="셀 스타일">셀 스타일</button>
          </div>
        </div>
      </div>

      <!-- 셀 그룹 -->
      <div class="ribbon-group">
        <div class="ribbon-group-title">셀</div>
        <div class="ribbon-controls">
          <div class="ribbon-button-group">
            <button class="format-button" title="삽입">삽입</button>
            <button class="format-button" title="삭제">삭제</button>
            <button class="format-button" title="서식">서식</button>
          </div>
        </div>
      </div>
    </div>

    <!-- 수식 바 -->
    <div class="formula-bar">
      <div class="name-box">A1</div>
      <div class="formula-label">fx</div>
      <input type="text" class="formula-input" placeholder="" />
    </div>
  </div>
</template>

<script setup>
// 껍데기만 필요하므로 스크립트 부분은 간단하게 유지
</script>

<style scoped>
.excel-ribbon {
  font-family: 'Segoe UI', '맑은 고딕', sans-serif;
  background-color: #f3f3f3;
  border-bottom: 1px solid #d4d4d4;
  user-select: none;
}

.ribbon-tabs {
  display: flex;
  background-color: #217346;
  color: white;
}

.ribbon-tab {
  padding: 8px 15px;
  cursor: pointer;
}

.ribbon-tab.active {
  background-color: #f3f3f3;
  color: #217346;
  font-weight: bold;
}

.ribbon-content {
  display: flex;
  padding: 5px;
  border-bottom: 1px solid #d4d4d4;
}

.ribbon-group {
  display: flex;
  flex-direction: column;
  border-right: 1px solid #e1e1e1;
  padding: 0 5px;
  margin-right: 5px;
  min-width: 80px;
  position: relative;
}

.ribbon-group-title {
  font-size: 11px;
  text-align: center;
  color: #666;
  margin-top: 2px;
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
}

.ribbon-controls {
  display: flex;
  flex-direction: column;
  padding: 2px 0 18px 0;
}

.ribbon-button {
  display: flex;
  flex-direction: column;
  align-items: center;
  padding: 3px 5px;
  margin: 2px;
  cursor: pointer;
  border-radius: 3px;
}

.ribbon-button:hover {
  background-color: #e5e5e5;
}

.ribbon-button.big {
  width: 45px;
}

.ribbon-button.small {
  flex-direction: row;
  font-size: 11px;
}

.ribbon-button .icon {
  font-size: 16px;
}

.ribbon-button .label {
  font-size: 11px;
  white-space: nowrap;
}

.ribbon-button-group {
  display: flex;
  flex-wrap: wrap;
  margin: 2px 0;
}

.format-button {
  padding: 3px 5px;
  margin: 1px;
  background-color: transparent;
  border: 1px solid transparent;
  cursor: pointer;
  border-radius: 3px;
  font-size: 12px;
  min-width: 24px;
  min-height: 24px;
}

.format-button:hover {
  border-color: #b8b8b8;
  background-color: #e5e5e5;
}

.font-dropdown, .size-dropdown, .format-dropdown {
  padding: 2px;
  font-size: 12px;
  margin: 2px;
  border: 1px solid #b8b8b8;
}

.font-dropdown {
  width: 120px;
}

.size-dropdown {
  width: 45px;
}

.format-dropdown {
  width: 90px;
}

.font-controls {
  display: flex;
  margin-bottom: 2px;
}

.format-controls {
  display: flex;
  flex-wrap: wrap;
}

.color-picker {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 24px;
  height: 24px;
  border: 1px solid transparent;
  cursor: pointer;
  border-radius: 3px;
  margin: 1px;
}

.color-picker:hover {
  border-color: #b8b8b8;
  background-color: #e5e5e5;
}

.formula-bar {
  display: flex;
  align-items: center;
  padding: 3px 5px;
  background-color: #fff;
  border-bottom: 1px solid #d4d4d4;
}

.name-box {
  width: 80px;
  padding: 3px 5px;
  border: 1px solid #b8b8b8;
  margin-right: 5px;
  font-size: 12px;
}

.formula-label {
  padding: 0 5px;
  font-size: 12px;
  color: #666;
}

.formula-input {
  flex-grow: 1;
  padding: 3px 5px;
  border: 1px solid #b8b8b8;
  font-size: 12px;
}
</style>