<template>
  <div class="container">
    <!-- 컴포넌트화된 툴바 사용 -->
    <RibonMenu
      v-model:formulaText="formulaText"
      v-model:selectedFont="selectedFont"
      v-model:fontSize="fontSize"
      v-model:textColor="textColor"
      v-model:backgroundColor="backgroundColor"
      v-model:isBold="isBold"
      v-model:isItalic="isItalic"
      v-model:isUnderline="isUnderline"
      :spread="spreadRef"
      :currentSelection="currentSelection"
      @formula-applied="onFormulaApplied"
      @style-updated="onStyleUpdated"
    />


    <div class="relative">
      <gc-spread-sheets
          class="spread-host"
          @workbookInitialized="initWorkbook"
          @selectionChanged="onSelectionChanged"
      >
      </gc-spread-sheets>
    </div>
  </div>
</template>

<script setup>
import { ref, inject } from 'vue';
import '@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css';
import { GcSpreadSheets } from '@mescius/spread-sheets-vue';
import * as GC from "@mescius/spread-sheets";
import '@mescius/spread-sheets-resources-ko';

// loading composable 주입
const { startLoading, stopLoading } = inject('loading');

// 한국어 설정
GC.Spread.Common.CultureManager.culture("ko-kr");

// spread 객체를 ref로 관리
const spreadRef = ref(null);

// 서식 설정을 위한 상태 추가
const selectedFont = ref('맑은 고딕');
const fontSize = ref(11);
const isBold = ref(false);
const isItalic = ref(false);
const isUnderline = ref(false);
const textColor = ref('#000000');
const backgroundColor = ref('#ffffff');
const formulaText = ref('');
const currentSelection = ref(null);

// 워크북 초기화
const initWorkbook = (spread) => {
  startLoading();
  try {
    // spread 객체 저장
    spreadRef.value = spread;

    const sheet = spread.getActiveSheet();
  } catch (error) {
    console.error('스프레드시트 초기화 중 오류:', error);
  } finally {
    stopLoading();
  }
};

// 수식 적용 후 콜백
const onFormulaApplied = (formula) => {
  console.log('수식이 적용되었습니다:', formula);
};

// 스타일 업데이트 후 콜백
const onStyleUpdated = () => {
  console.log('스타일이 업데이트되었습니다.');
};

// 선택 영역 변경 이벤트 핸들러
const onSelectionChanged = (e) => {
  const sheet = spreadRef.value.getActiveSheet();
  const selections = sheet.getSelections();

  if (selections && selections.length > 0) {
    currentSelection.value = selections[0];

    // 선택된 셀의 수식 가져오기
    const formula = sheet.getFormula(
      currentSelection.value.row,
      currentSelection.value.col
    );

    formulaText.value = formula ? '=' + formula : '';

    // 서식 정보도 업데이트
    updateFormatInfo();
  }
};

// 선택된 셀의 서식 정보 업데이트 함수
const updateFormatInfo = () => {
  if (!spreadRef.value || !currentSelection.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const { row, col } = currentSelection.value;

  // 스타일 정보 가져오기
  const style = sheet.getStyle(row, col);

  if (style) {
    // 폰트 패밀리
    if (style.font && style.font.family) {
      selectedFont.value = style.font.family;
    }

    // 폰트 크기
    if (style.font && style.font.size) {
      fontSize.value = parseInt(style.font.size);
    }

    // 굵게
    isBold.value = style.font && style.font.bold === true;

    // 기울임
    isItalic.value = style.font && style.font.italic === true;

    // 밑줄
    isUnderline.value = style.font && style.font.underline === true;

    // 텍스트 색상
    if (style.foreColor) {
      textColor.value = style.foreColor;
    }

    // 배경 색상
    if (style.backColor) {
      backgroundColor.value = style.backColor;
    }
  } else {
    // 스타일이 없는 경우 기본값으로 설정
    selectedFont.value = '맑은 고딕';
    fontSize.value = 11;
    isBold.value = false;
    isItalic.value = false;
    isUnderline.value = false;
    textColor.value = '#000000';
    backgroundColor.value = '#ffffff';
  }
};

</script>

<style scoped>
.container {
  display: flex;
  flex-direction: column;
  height: 100%;
}

.relative {
  position: relative;
  flex: 1;
  min-height: 400px;
}

.spread-host {
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
}
</style>
