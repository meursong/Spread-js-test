<template>
  <div class="container">
    <div class="control-panel">
      <button class="action-button" @click="calculateAll">계산 실행</button>
      <button class="action-button" @click="checkFormulas">수식 확인</button>
    </div>
    <div class="relative">
      <gc-spread-sheets
          class="spread-host"
          @workbookInitialized="initWorkbook"
          @valueChanged="onValueChanged"
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

// 샘플 데이터와 수식 설정
const sampleData = [
  { label: '숫자 1', value: 10 },
  { label: '숫자 2', value: 20 },
  { label: 'SUM 수식', formula: '=SUM(B1:B2)' },
  { label: 'AVERAGE 수식', formula: '=AVERAGE(B1:B2)' },
  { label: 'MAX 수식', formula: '=MAX(B1:B2)' },
  { label: 'MIN 수식', formula: '=MIN(B1:B2)' },
  { label: 'COUNT 수식', formula: '=COUNT(B1:B2)' },
  { label: '조건부 수식', formula: '=IF(B1>B2,"숫자1이 크다","숫자2가 크다")' },
  { label: '반올림 수식', formula: '=ROUND(B1/3,2)' },
  { label: '문자결합 수식', formula: '=CONCATENATE("계산값: ",B1)' }
];

// 스타일 설정 함수
const setStyles = (sheet) => {
  sampleData.forEach((item, index) => {
    // 레이블 셀 스타일링
    const labelStyle = new GC.Spread.Sheets.Style();
    labelStyle.backColor = "#4472C4";
    labelStyle.foreColor = "white";
    labelStyle.font = "bold 12pt 맑은 고딕";
    labelStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
    sheet.setStyle(index, 0, labelStyle);

    // 값/수식 셀 스타일링
    const valueStyle = new GC.Spread.Sheets.Style();
    valueStyle.font = "11pt 맑은 고딕";

    if (item.formula) {
      valueStyle.backColor = "#FFE699"; // 수식 셀 강조
    } else {
      valueStyle.backColor = "#E6E6E6";
    }
    sheet.setStyle(index, 1, valueStyle);
  });

  // 컬럼 너비 설정
  sheet.setColumnWidth(0, 150);
  sheet.setColumnWidth(1, 250);
};

// 워크북 초기화
const initWorkbook = (spread) => {
  startLoading();
  try {
    // spread 객체 저장
    spreadRef.value = spread;

    const sheet = spread.getActiveSheet();

    // 자동 계산 설정
    spread.options.calcOnDemand = false; // 자동 계산 활성화

    // 데이터와 수식 설정
    sampleData.forEach((item, index) => {
      sheet.setValue(index, 0, item.label);
      if (item.formula) {
        sheet.setFormula(index, 1, item.formula);
      } else {
        sheet.setValue(index, 1, item.value);
      }
    });

    // 스타일 적용
    setStyles(sheet);

  } catch (error) {
    console.error('스프레드시트 초기화 중 오류:', error);
  } finally {
    stopLoading();
  }
};

// 값 변경 이벤트 핸들러
const onValueChanged = (args) => {
  const { sheet, row, col, newValue } = args;
  console.log(`셀 (${row}, ${col}) 변경됨: ${newValue}`);
};

// 모든 수식 재계산
const calculateAll = () => {
  if (spreadRef.value) {
    spreadRef.value.calculateAll();
  }
};

// 수식 확인
const checkFormulas = () => {
  if (!spreadRef.value) {
    alert('스프레드시트가 초기화되지 않았습니다.');
    return;
  }

  const sheet = spreadRef.value.getActiveSheet();

  let message = "현재 수식 상태:\n\n";
  sampleData.forEach((item, index) => {
    if (item.formula) {
      const formula = sheet.getFormula(index, 1);
      const value = sheet.getValue(index, 1);
      message += `${item.label}:\n수식: ${formula}\n결과: ${value}\n\n`;
    }
  });

  alert(message);
};

</script>

<style scoped>
.container {
  padding: 1rem;
}

.control-panel {
  margin-bottom: 1rem;
  display: flex;
  gap: 1rem;
}

.action-button {
  padding: 0.5rem 1rem;
  background-color: #4472C4;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-weight: bold;
}

.action-button:hover {
  background-color: #365899;
}

.spread-host {
  width: 90vw;
  height: 80vh;
  border: 1px solid #ccc;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}
</style>