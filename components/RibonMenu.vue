<template>
  <div>
    <div class="control-panel">
      <button class="action-button" @click="handleInsertFunction('SUM')">=SUM()</button>
      <button class="action-button" @click="handleInsertFunction('COUNT')">=COUNT()</button>
      <button class="action-button" @click="handleInsertFunction('MAX')">=MAX()</button>
      <button class="action-button" @click="handleInsertFunction('MIN')">=MIN()</button>
      <!-- 수식 입력 바 -->
      <div class="formula-bar">
        <span>fx</span>
        <input
            v-model="localFormulaText"
            @keyup.enter="handleApplyFormula"
            @click="handleCopyFormulaToClipboard"
            placeholder="수식을 입력하세요"
        />
      </div>
    </div>

    <div class="control-panel">
      <button class="action-button" @click="handleMergeCells">선택 영역 병합</button>
      <button class="action-button" @click="handleUnmergeCells">병합 해제</button>
    </div>

    <!-- 텍스트 서식 패널 추가 -->
    <div class="control-panel formatting-panel">
      <div class="format-group">
        <select
            class="format-control"
            v-model="localSelectedFont"
            @change="handleApplyFont"
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
            v-model="localFontSize"
            @change="handleApplyFontSize"
        >
          <option v-for="size in [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36]" :key="size" :value="size">
            {{ size }}pt
          </option>
        </select>
      </div>

      <div class="format-group">
        <button
            class="format-button"
            :class="{active: localIsBold}"
            @click="handleToggleBold"
            title="굵게"
        >
          <strong>B</strong>
        </button>
        <button
            class="format-button"
            :class="{active: localIsItalic}"
            @click="handleToggleItalic"
            title="기울임"
        >
          <em>I</em>
        </button>
        <button
            class="format-button"
            :class="{active: localIsUnderline}"
            @click="handleToggleUnderline"
            title="밑줄"
        >
          <u>U</u>
        </button>
      </div>

      <div class="format-group">
        <!-- 텍스트 정렬 버튼 -->
        <button class="format-button" @click="handleApplyAlignment('left')" title="왼쪽 정렬">
          ◀
        </button>
        <button class="format-button" @click="handleApplyAlignment('center')" title="가운데 정렬">
          ■
        </button>
        <button class="format-button" @click="handleApplyAlignment('right')" title="오른쪽 정렬">
          ▶
        </button>
      </div>

      <!-- 색상 선택 도구 -->
      <div class="format-group">
        <div class="color-picker">
          <label>글자색:</label>
          <input
              type="color"
              v-model="localTextColor"
              @change="handleApplyTextColor"
          />
        </div>

        <div class="color-picker">
          <label>배경색:</label>
          <input
              type="color"
              v-model="localBackgroundColor"
              @change="handleApplyBackgroundColor"
          />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch } from 'vue';
import * as GC from "@mescius/spread-sheets";

// props 정의
const props = defineProps({
  formulaText: String,
  selectedFont: String,
  fontSize: Number,
  isBold: Boolean,
  isItalic: Boolean,
  isUnderline: Boolean,
  textColor: String,
  backgroundColor: String,
  spread: Object,
  currentSelection: Object
});

// 이벤트 정의
const emit = defineEmits([
  'update:formulaText',
  'update:selectedFont',
  'update:fontSize',
  'update:isBold',
  'update:isItalic',
  'update:isUnderline',
  'update:textColor',
  'update:backgroundColor',
  'formula-applied',
  'style-updated'
]);

// 로컬 상태 (양방향 바인딩용)
const localFormulaText = computed({
  get: () => props.formulaText,
  set: (val) => emit('update:formulaText', val)
});

const localSelectedFont = computed({
  get: () => props.selectedFont,
  set: (val) => emit('update:selectedFont', val)
});

const localFontSize = computed({
  get: () => props.fontSize,
  set: (val) => emit('update:fontSize', val)
});

const localIsBold = computed({
  get: () => props.isBold,
  set: (val) => emit('update:isBold', val)
});

const localIsItalic = computed({
  get: () => props.isItalic,
  set: (val) => emit('update:isItalic', val)
});

const localIsUnderline = computed({
  get: () => props.isUnderline,
  set: (val) => emit('update:isUnderline', val)
});

const localTextColor = computed({
  get: () => props.textColor,
  set: (val) => emit('update:textColor', val)
});

const localBackgroundColor = computed({
  get: () => props.backgroundColor,
  set: (val) => emit('update:backgroundColor', val)
});

// 스프레드 시트 관련 함수들
const getActiveSheet = () => {
  if (!props.spread) return null;
  return props.spread.getActiveSheet();
};

// 선택된 셀에 스타일 적용하는 헬퍼 함수
const applyStyleToSelection = (styleUpdater) => {
  const sheet = getActiveSheet();
  if (!sheet || !props.currentSelection) return;

  const { row, col, rowCount, colCount } = props.currentSelection;

  // 선택 영역의 모든 셀에 스타일 적용
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const cellRow = row + r;
      const cellCol = col + c;

      let style = sheet.getStyle(cellRow, cellCol) || new GC.Spread.Sheets.Style();
      // 스타일 업데이터 함수 호출하여 스타일 수정
      styleUpdater(style);
      sheet.setStyle(cellRow, cellCol, style);
    }
  }

  sheet.repaint();
  emit('style-updated');
};

// 수식 관련 함수
const handleInsertFunction = (funcName) => {
  localFormulaText.value = `=${funcName}()`;
};

const handleApplyFormula = () => {
  const sheet = getActiveSheet();
  if (!sheet || !props.currentSelection) return;

  const { row, col } = props.currentSelection;
  sheet.setFormula(row, col, localFormulaText.value);
  emit('formula-applied', localFormulaText.value);
};

const handleCopyFormulaToClipboard = () => {
  const sheet = getActiveSheet();
  if (!sheet || !props.currentSelection) return;

  const { row, col } = props.currentSelection;
  const cellFormula = sheet.getFormula(row, col);
  if (cellFormula) {
    localFormulaText.value = cellFormula;
    // 클립보드에 복사하는 로직 (선택사항)
    navigator.clipboard.writeText(cellFormula).catch(err => {
      console.error('클립보드 복사 실패:', err);
    });
  }
};

// 셀 병합 함수
const handleMergeCells = () => {
  const sheet = getActiveSheet();
  console.log(props.currentSelection)
  if (!sheet || !props.currentSelection) return;

  const { row, col, rowCount, colCount } = props.currentSelection;
  sheet.addSpan(row, col, rowCount, colCount);
  sheet.repaint();
};

const handleUnmergeCells = () => {
  const sheet = getActiveSheet();
  if (!sheet || !props.currentSelection) return;

  const { row, col } = props.currentSelection;
  sheet.removeSpan(row, col);
  sheet.repaint();
};

// 서식 적용 함수들
const handleApplyFont = () => {
  applyStyleToSelection((style) => {
    const fontStyles = [];
    if (localIsBold.value) fontStyles.push('bold');
    if (localIsItalic.value) fontStyles.push('italic');

    const fontStyle = fontStyles.join(' ');
    style.font = `${fontStyle ? fontStyle + ' ' : ''}${localFontSize.value}pt ${localSelectedFont.value}`;
  });
};

const handleApplyFontSize = () => {
  applyStyleToSelection((style) => {
    const fontStyles = [];
    if (localIsBold.value) fontStyles.push('bold');
    if (localIsItalic.value) fontStyles.push('italic');

    const fontStyle = fontStyles.join(' ');
    style.font = `${fontStyle ? fontStyle + ' ' : ''}${localFontSize.value}pt ${localSelectedFont.value}`;
  });
};

const handleToggleBold = () => {
  localIsBold.value = !localIsBold.value;
  applyStyleToSelection((style) => {
    const fontStyles = [];
    if (localIsBold.value) fontStyles.push('bold');
    if (localIsItalic.value) fontStyles.push('italic');

    const fontStyle = fontStyles.join(' ');
    style.font = `${fontStyle ? fontStyle + ' ' : ''}${localFontSize.value}pt ${localSelectedFont.value}`;
  });
};

const handleToggleItalic = () => {
  localIsItalic.value = !localIsItalic.value;
  applyStyleToSelection((style) => {
    const fontStyles = [];
    if (localIsBold.value) fontStyles.push('bold');
    if (localIsItalic.value) fontStyles.push('italic');

    const fontStyle = fontStyles.join(' ');
    style.font = `${fontStyle ? fontStyle + ' ' : ''}${localFontSize.value}pt ${localSelectedFont.value}`;
  });
};

const handleToggleUnderline = () => {
  localIsUnderline.value = !localIsUnderline.value;
  applyStyleToSelection((style) => {
    style.textDecoration = localIsUnderline.value
        ? GC.Spread.Sheets.TextDecorationType.underline
        : GC.Spread.Sheets.TextDecorationType.none;
  });
};

const handleApplyAlignment = (align) => {
  applyStyleToSelection((style) => {
    switch (align) {
      case 'left':
        style.hAlign = GC.Spread.Sheets.HorizontalAlign.left;
        break;
      case 'center':
        style.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
        break;
      case 'right':
        style.hAlign = GC.Spread.Sheets.HorizontalAlign.right;
        break;
    }
  });
};

const handleApplyTextColor = () => {
  applyStyleToSelection((style) => {
    style.foreColor = localTextColor.value;
  });
};

const handleApplyBackgroundColor = () => {
  applyStyleToSelection((style) => {
    style.backColor = localBackgroundColor.value;
  });
};

// 현재 선택이 변경될 때마다 수식 텍스트 업데이트
watch(() => props.currentSelection, (newSelection) => {
  if (!newSelection || !getActiveSheet()) return;

  const sheet = getActiveSheet();
  const { row, col } = newSelection;

  // 수식 표시 업데이트
  const formula = sheet.getFormula(row, col);
  if (formula) {
    localFormulaText.value = formula;
  } else {
    localFormulaText.value = '';
  }
}, { deep: true });
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