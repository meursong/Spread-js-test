<template>
  <div>
    <div class="control-panel">
      <button class="action-button" @click="$emit('insert-function', 'SUM')">=SUM()</button>
      <button class="action-button" @click="$emit('insert-function', 'COUNT')">=COUNT()</button>
      <button class="action-button" @click="$emit('insert-function', 'MAX')">=MAX()</button>
      <button class="action-button" @click="$emit('insert-function', 'MIN')">=MIN()</button>
      <!-- 수식 입력 바 -->
      <div class="formula-bar">
        <span>fx</span>
        <input
            :value="formulaText"
            @input="$emit('update:formulaText', $event.target.value)"
            @keyup.enter="$emit('apply-formula')"
            @click="$emit('copy-formula')"
            placeholder="수식을 입력하세요"
        />
      </div>
    </div>

    <div class="control-panel">
      <button class="action-button" @click="$emit('merge-cells')">선택 영역 병합</button>
      <button class="action-button" @click="$emit('unmerge-cells')">병합 해제</button>
    </div>

    <!-- 텍스트 서식 패널 추가 -->
    <div class="control-panel formatting-panel">
      <div class="format-group">
        <select
            class="format-control"
            :value="selectedFont"
            @change="$emit('update:selectedFont', $event.target.value)"
        >
          <option value="맑은 고딕">맑은 고딕</option>
          <option value="굴림">굴림</option>
          <option value="돋움">돋움</option>
          <option value="바탕">바탕</option>
          <option value="Arial">Arial</option>
          <option value="Verdana">Verdana</option>
          <option value="Times New Roman">Times New Roman</option>
        </select>

        <select
            class="format-control"
            :value="fontSize"
            @change="$emit('update:fontSize', Number($event.target.value))"
        >
          <option v-for="size in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36]" :key="size" :value="size">
            {{ size }}pt
          </option>
        </select>
      </div>

      <div class="format-group">
        <button
            class="format-button"
            :class="{active: isBold}"
            @click="$emit('toggle-bold')"
            title="굵게"
        >
          <strong>B</strong>
        </button>
        <button
            class="format-button"
            :class="{active: isItalic}"
            @click="$emit('toggle-italic')"
            title="기울임"
        >
          <em>I</em>
        </button>
        <button
            class="format-button"
            :class="{active: isUnderline}"
            @click="$emit('toggle-underline')"
            title="밑줄"
        >
          <u>U</u>
        </button>
      </div>

      <div class="format-group">
        <!-- 텍스트 정렬 버튼 -->
        <button class="format-button" @click="$emit('apply-alignment', 'left')" title="왼쪽 정렬">
          ◀
        </button>
        <button class="format-button" @click="$emit('apply-alignment', 'center')" title="가운데 정렬">
          ■
        </button>
        <button class="format-button" @click="$emit('apply-alignment', 'right')" title="오른쪽 정렬">
          ▶
        </button>
      </div>

      <!-- 색상 선택 도구 -->
      <div class="format-group">
        <div class="color-picker">
          <label>글자색:</label>
          <input
              type="color"
              :value="textColor"
              @change="$emit('update:textColor', $event.target.value)"
          />
        </div>

        <div class="color-picker">
          <label>배경색:</label>
          <input
              type="color"
              :value="backgroundColor"
              @change="$emit('update:backgroundColor', $event.target.value)"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
// props 정의
const props = defineProps({
  formulaText: String,
  selectedFont: String,
  fontSize: Number,
  isBold: Boolean,
  isItalic: Boolean,
  isUnderline: Boolean,
  textColor: String,
  backgroundColor: String
});

// 이벤트 정의
const emit = defineEmits([
  'insert-function',
  'update:formulaText',
  'apply-formula',
  'copy-formula',
  'merge-cells',
  'unmerge-cells',
  'update:selectedFont',
  'update:fontSize',
  'toggle-bold',
  'toggle-italic',
  'toggle-underline',
  'apply-alignment',
  'update:textColor',
  'update:backgroundColor'
]);
</script>

<style scoped>
.control-panel {
  display: flex;
  margin-bottom: 8px;
  flex-wrap: wrap;
  align-items: center;
  gap: 5px;
}

.formatting-panel {
  background-color: #f5f5f5;
  padding: 8px;
  border-radius: 4px;
}

.format-group {
  display: flex;
  gap: 4px;
  align-items: center;
  margin-right: 10px;
}

.action-button, .format-button {
  padding: 6px 10px;
  background-color: #f0f0f0;
  border: 1px solid #ccc;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
}

.action-button:hover, .format-button:hover {
  background-color: #e0e0e0;
}

.format-button {
  min-width: 28px;
  text-align: center;
  padding: 5px 8px;
}

.format-button.active {
  background-color: #d0d0d0;
  border-color: #a0a0a0;
}

.format-control {
  padding: 5px;
  border: 1px solid #ccc;
  border-radius: 4px;
}

.color-picker {
  display: flex;
  align-items: center;
  gap: 4px;
}

.color-picker label {
  font-size: 12px;
}

.color-picker input[type="color"] {
  width: 24px;
  height: 24px;
  border: 1px solid #ccc;
  padding: 0;
  cursor: pointer;
}

.formula-bar {
  display: flex;
  align-items: center;
  gap: 5px;
  margin-left: 10px;
  flex: 1;
}

.formula-bar input {
  flex: 1;
  padding: 5px;
  border: 1px solid #ccc;
  border-radius: 4px;
}
</style>