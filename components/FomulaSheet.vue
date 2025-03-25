<template>
  <div class="container">
    <div class="control-panel">
      <button class="action-button" @click="insertFunction('SUM')">=SUM()</button>
      <button class="action-button" @click="insertFunction('COUNT')">=COUNT()</button>
      <button class="action-button" @click="insertFunction('MAX')">=MAX()</button>
      <button class="action-button" @click="insertFunction('MIN')">=MIN()</button>
      <!-- 수식 입력 바 -->
      <div class="formula-bar">
        <span>fx</span>
        <input
          v-model="formulaText"
          @keyup.enter="applyFormula"
          @click="copyFormulaToClipboard"
          placeholder="수식을 입력하세요"
        />
      </div>
    </div>
    <div class="control-panel">
      <button class="action-button" @click="mergeCells">선택 영역 병합</button>
      <button class="action-button" @click="unmergeCells">병합 해제</button>
    </div>
    <div class="relative">
      <gc-spread-sheets
          class="spread-host"
          @workbookInitialized="initWorkbook"
          @valueChanged="onValueChanged"
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
  }
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
  console.table(args);
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
// 수식 적용 함수
const applyFormula = () => {
  if (!spreadRef.value || !spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const { row, col } = currentSelection.value;

  try {
    // 수식에서 = 제거하고 적용
    const formula = formulaText.value.startsWith('=')
      ? formulaText.value.substring(1)
      : formulaText.value;

    sheet.setFormula(row, col, formula);

    // 수식 적용 후 계산 실행
    sheet.recalcAll();
  } catch (error) {
    alert('수식 적용 중 오류가 발생했습니다: ' + error.message);
  }
};

const formulaText = ref('');
const currentSelection = ref(null);

// 미리 정의된 함수 삽입
const insertFunction = (funcName) => {
  if (!currentSelection.value) {
    alert('먼저 셀을 선택해주세요.');
    return;
  }

  const selection = currentSelection.value;
  const range = `${getColumnName(selection.col)}${selection.row + 1}:${
    getColumnName(selection.col + selection.colCount - 1)
  }${selection.row + selection.rowCount}`;

  formulaText.value = `=${funcName}(${range})`;
};

// 열 이름 가져오기 (A, B, C, ...)
const getColumnName = (index) => {
  let name = '';
  while (index >= 0) {
    name = String.fromCharCode(65 + (index % 26)) + name;
    index = Math.floor(index / 26) - 1;
  }
  return name;
};

// 수식을 클립보드에 복사하는 함수
const copyFormulaToClipboard = () => {
  if (!formulaText.value) return;

  navigator.clipboard.writeText(formulaText.value)
    .then(() => {
      // 선택적: 복사 성공 표시
      const originalBorderColor = document.querySelector('.formula-bar').style.borderColor;
      document.querySelector('.formula-bar').style.borderColor = '#4CAF50';

      setTimeout(() => {
        document.querySelector('.formula-bar').style.borderColor = originalBorderColor;
      }, 500);
    })
    .catch(err => {
      console.error('클립보드 복사 실패:', err);
    });
};

// 선택된 셀 병합하기
const mergeCells = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const selections = sheet.getSelections();

  if (!selections || selections.length === 0) {
    alert('병합할 셀을 먼저 선택해주세요.');
    return;
  }

  try {
    // 각 선택 영역에 대해 병합 수행
    selections.forEach(selection => {
      const { row, col, rowCount, colCount } = selection;

      // 병합 전에 첫번째 셀의 값 저장 (병합 시 첫 셀 값만 유지됨)
      const firstCellValue = sheet.getValue(row, col);
      const firstCellFormula = sheet.getFormula(row, col);

      // 셀 병합
      sheet.addSpan(row, col, rowCount, colCount);

      // 병합된 셀에 원래 첫 셀의 값이나 수식 적용
      if (firstCellFormula) {
        sheet.setFormula(row, col, firstCellFormula);
      } else if (firstCellValue !== null && firstCellValue !== undefined) {
        sheet.setValue(row, col, firstCellValue);
      }

      // 병합된 셀의 텍스트 정렬 가운데로 설정
      const style = new GC.Spread.Sheets.Style();
      style.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
      style.vAlign = GC.Spread.Sheets.VerticalAlign.center;
      sheet.setStyle(row, col, style);
    });

    // 변경사항 알림
    alert('선택한 영역이 병합되었습니다.');
  } catch (error) {
    console.error('셀 병합 중 오류 발생:', error);
    alert('셀 병합 중 오류가 발생했습니다.');
  }
};


// 셀 병합 해제 기능
const unmergeCells = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const selections = sheet.getSelections();

  if (!selections || selections.length === 0) {
    alert('병합 해제할 셀을 먼저 선택해주세요.');
    return;
  }

  // 선택된 영역의 병합 해제
  selections.forEach(selection => {
    // 선택 영역 내의 모든 셀 검사
    for (let row = selection.row; row < selection.row + selection.rowCount; row++) {
      for (let col = selection.col; col < selection.col + selection.colCount; col++) {
        const span = sheet.getSpan(row, col);
        if (span) {
          // 병합된 셀 발견 시 병합 해제
          sheet.removeSpan(span.row, span.col);
        }
      }
    }
  });
};


</script>

<style scoped>
.container {
  display: flex;
  flex-direction: column;
  gap: 16px;
  padding: 20px;
  background-color: #f8f9fa;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.control-panel {
  display: flex;
  gap: 12px;
  flex-wrap: wrap;
  align-items: center;
  padding: 12px;
  background-color: white;
  border-radius: 6px;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.06);
}

.action-button {
  padding: 8px 16px;
  font-size: 14px;
  font-weight: 500;
  color: #ffffff;
  background: linear-gradient(135deg, #4472C4 0%, #2c5aa0 100%);
  border: none;
  border-radius: 4px;
  cursor: pointer;
  transition: all 0.2s ease;
  min-width: 100px;
  height: 36px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.action-button:hover {
  background: linear-gradient(135deg, #2c5aa0 0%, #1e3c6a 100%);
  transform: translateY(-1px);
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.action-button:active {
  transform: translateY(0);
}

.formula-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-grow: 1;
  background-color: #ffffff;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  padding: 4px 12px;
  height: 36px;
  transition: all 0.2s ease;
}

.formula-bar:focus-within {
  border-color: #4472C4;
  box-shadow: 0 0 0 2px rgba(68, 114, 196, 0.2);
}

.formula-bar span {
  color: #4472C4;
  font-weight: 600;
  font-size: 14px;
  padding-right: 8px;
  border-right: 1px solid #e0e0e0;
}

.formula-bar input {
  flex-grow: 1;
  border: none;
  outline: none;
  font-size: 14px;
  color: #333333;
  background: transparent;
  padding: 0 8px;
}

.formula-bar input::placeholder {
  color: #999999;
}

.relative {
  position: relative;
  height: 600px; /* 스프레드시트 높이 조정 */
  border-radius: 6px;
  overflow: hidden;
  border: 1px solid #e0e0e0;
}

.spread-host {
  width: 100%;
  height: 100%;
}

/* 반응형 디자인 */
@media (max-width: 768px) {
  .control-panel {
    flex-direction: column;
    align-items: stretch;
  }

  .formula-bar {
    width: 100%;
  }

  .action-button {
    width: 100%;
  }
}

/* 다크 모드 지원 */
@media (prefers-color-scheme: dark) {
  .container {
    background-color: #1e1e1e;
  }

  .control-panel {
    background-color: #2d2d2d;
  }

  .formula-bar {
    background-color: #2d2d2d;
    border-color: #404040;
  }

  .formula-bar span {
    color: #7aa2f7;
    border-right-color: #404040;
  }

  .formula-bar input {
    color: #ffffff;
  }

  .formula-bar input::placeholder {
    color: #666666;
  }

  .relative {
    border-color: #404040;
  }
}
</style>
