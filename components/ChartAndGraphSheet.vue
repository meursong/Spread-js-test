<template>
  <div class="container">
    <!-- 컴포넌트화된 툴바 사용 -->
    <RibonMenu
      v-model:formulaText="formulaText"
      v-model:selectedFont="selectedFont"
      v-model:fontSize="fontSize"
      v-model:textColor="textColor"
      v-model:backgroundColor="backgroundColor"
      :isBold="isBold"
      :isItalic="isItalic"
      :isUnderline="isUnderline"
      @insert-function="insertFunction"
      @apply-formula="applyFormula"
      @copy-formula="copyFormulaToClipboard"
      @merge-cells="mergeCells"
      @unmerge-cells="unmergeCells"
      @toggle-bold="toggleBold"
      @toggle-italic="toggleItalic"
      @toggle-underline="toggleUnderline"
      @apply-alignment="applyAlignment"
      @update:textColor="applyTextColor"
      @update:backgroundColor="applyBackgroundColor"
      @update:selectedFont="applyFont"
      @update:fontSize="applyFontSize"
    />

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

// 서식 설정을 위한 상태 추가
const selectedFont = ref('맑은 고딕');
const fontSize = ref(11);
const isBold = ref(false);
const isItalic = ref(false);
const isUnderline = ref(false);
const textColor = ref('#000000');
const backgroundColor = ref('#ffffff');

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

// 선택된 셀의 서식 정보 업데이트
const updateFormatInfo = () => {
  if (!spreadRef.value || !currentSelection.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const { row, col } = currentSelection.value;
  const style = sheet.getStyle(row, col);

  if (style) {
    // 글꼴 정보 파싱
    if (style.font) {
      const fontInfo = style.font.split(' ');

      // 볼드 체크
      isBold.value = style.font.includes('bold');

      // 이탤릭 체크
      isItalic.value = style.font.includes('italic');

      // 글꼴 크기 (예: '11pt 맑은 고딕' 형식에서 추출)
      const sizeMatch = style.font.match(/(\d+)pt/);
      if (sizeMatch) {
        fontSize.value = parseInt(sizeMatch[1]);
      }

      // 글꼴 이름
      const fontFamilyParts = style.font.split('pt ');
      if (fontFamilyParts.length > 1) {
        selectedFont.value = fontFamilyParts[1].replace(/bold|italic|\s+/g, ' ').trim();
      }
    }

    // 색상 정보
    if (style.foreColor) {
      textColor.value = style.foreColor;
    }

    if (style.backColor) {
      backgroundColor.value = style.backColor;
    }
  }
};

// 선택된 셀에 서식 적용 함수
const applyStyleToSelection = (styleUpdater) => {
  if (!spreadRef.value || !currentSelection.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const { row, col, rowCount, colCount } = currentSelection.value;

  // 선택된 영역의 모든 셀에 스타일 적용
  for (let r = row; r < row + rowCount; r++) {
    for (let c = col; c < col + colCount; c++) {
      // 기존 스타일 가져오기
      let style = sheet.getStyle(r, c) || new GC.Spread.Sheets.Style();

      // 스타일 업데이트
      style = styleUpdater(style);

      // 업데이트된 스타일 적용
      sheet.setStyle(r, c, style);
    }
  }
};

// 글꼴 적용
const applyFont = () => {
  applyStyleToSelection((style) => {
    // 기존 글꼴 정보 유지하면서 글꼴만 변경
    const fontParts = (style.font || '11pt 맑은 고딕').split('pt ');
    const fontWeight = isBold.value ? 'bold' : '';
    const fontStyle = isItalic.value ? 'italic' : '';
    const fontWeightAndStyle = (fontWeight + ' ' + fontStyle).trim();

    style.font = `${fontSize.value}pt ${selectedFont.value}`;
    if (fontWeightAndStyle) {
      style.font = `${fontWeightAndStyle} ${style.font}`;
    }

    return style;
  });
};

// 글꼴 크기 적용
const applyFontSize = () => {
  applyStyleToSelection((style) => {
    // 기존 글꼴 정보 유지하면서 크기만 변경
    const fontParts = (style.font || '11pt 맑은 고딕').split('pt ');
    const fontWeight = isBold.value ? 'bold' : '';
    const fontStyle = isItalic.value ? 'italic' : '';
    const fontWeightAndStyle = (fontWeight + ' ' + fontStyle).trim();

    style.font = `${fontSize.value}pt ${selectedFont.value}`;
    if (fontWeightAndStyle) {
      style.font = `${fontWeightAndStyle} ${style.font}`;
    }

    return style;
  });
};

// 굵게 설정
const toggleBold = () => {
  isBold.value = !isBold.value;

  applyStyleToSelection((style) => {
    // 기존 글꼴 정보 유지하면서 굵기만 변경
    const fontParts = (style.font || '11pt 맑은 고딕').split('pt ');
    const fontStyle = isItalic.value ? 'italic' : '';
    const fontWeightAndStyle = ((isBold.value ? 'bold' : '') + ' ' + fontStyle).trim();

    style.font = `${fontSize.value}pt ${selectedFont.value}`;
    if (fontWeightAndStyle) {
      style.font = `${fontWeightAndStyle} ${style.font}`;
    }

    return style;
  });
};

// 기울임꼴 설정
const toggleItalic = () => {
  isItalic.value = !isItalic.value;

  applyStyleToSelection((style) => {
    // 기존 글꼴 정보 유지하면서 기울임만 변경
    const fontParts = (style.font || '11pt 맑은 고딕').split('pt ');
    const fontWeight = isBold.value ? 'bold' : '';
    const fontWeightAndStyle = (fontWeight + ' ' + (isItalic.value ? 'italic' : '')).trim();

    style.font = `${fontSize.value}pt ${selectedFont.value}`;
    if (fontWeightAndStyle) {
      style.font = `${fontWeightAndStyle} ${style.font}`;
    }

    return style;
  });
};

// 밑줄 설정
const toggleUnderline = () => {
  isUnderline.value = !isUnderline.value;

  applyStyleToSelection((style) => {
    style.textDecoration = isUnderline.value
      ? GC.Spread.Sheets.TextDecorationType.underline
      : GC.Spread.Sheets.TextDecorationType.none;
    return style;
  });
};

// 텍스트 정렬 적용
const applyAlignment = (align) => {
  let hAlign;

  switch (align) {
    case 'left':
      hAlign = GC.Spread.Sheets.HorizontalAlign.left;
      break;
    case 'center':
      hAlign = GC.Spread.Sheets.HorizontalAlign.center;
      break;
    case 'right':
      hAlign = GC.Spread.Sheets.HorizontalAlign.right;
      break;
    default:
      hAlign = GC.Spread.Sheets.HorizontalAlign.general;
  }

  applyStyleToSelection((style) => {
    style.hAlign = hAlign;
    return style;
  });
};

// 글자색 적용
const applyTextColor = () => {
  applyStyleToSelection((style) => {
    style.foreColor = textColor.value;
    return style;
  });
};

// 배경색 적용
const applyBackgroundColor = () => {
  applyStyleToSelection((style) => {
    style.backColor = backgroundColor.value;
    return style;
  });
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
  height: 100%;
}

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
