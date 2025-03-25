<template>
  <div class="container">
    <button class="check-button" @click="checkData">작성한 엑셀 데이터 확인</button>
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

// spread 인스턴스를 저장하기 위한 ref 생성
const spreadInstance = ref(null);


// 워크북 초기화
const initWorkbook = (spread) => {
  startLoading();
  try {
    spreadInstance.value = spread;
    const sheet = spread.getActiveSheet();

  } catch (error) {
    console.error('스프레드시트 초기화 중 오류:', error);
  } finally {
    stopLoading();
  }
};

const data = ref({});

// 값 변경 이벤트 핸들러
const onValueChanged = (args) => {
  const { sheet, row, col, newValue } = args;
  if (col === 1) {
    const header = headers[row];
    let processedValue = newValue;

    // 데이터 타입에 따른 값 처리
    switch (header.type) {
      case 'number':
      case 'currency':
        processedValue = Number(newValue);
        break;
      case 'date':
      case 'time':
        processedValue = new Date(newValue);
        break;
      case 'percentage':
        processedValue = Number(newValue) / 100;
        break;
    }

    data.value[header.field] = processedValue;
  }
};


// 데이터 확인
const checkData = () => {
  startLoading();
  try {
    if (!spreadInstance.value) {
      throw new Error('스프레드시트가 초기화되지 않았습니다.');
    }

    const sheet = spreadInstance.value.getActiveSheet();
    const rowCount = sheet.getRowCount();
    const colCount = sheet.getColumnCount();

    console.group('스프레드시트 데이터');

    // 전체 데이터를 2차원 배열로 가져오기
    const allData = [];
    for (let row = 0; row < rowCount; row++) {
      const rowData = [];
      for (let col = 0; col < colCount; col++) {
        // 셀의 실제 값 가져오기
        const value = sheet.getValue(row, col);
        // 셀의 표시 텍스트 가져오기 (서식이 적용된 텍스트)
        const text = sheet.getText(row, col);
        // 셀의 서식 가져오기
        const formatter = sheet.getFormatter(row, col);
        // 셀의 수식 가져오기
        const formula = sheet.getFormula(row, col);
        const formulaInfo = sheet.getFormulaInformation(row, col);

        rowData.push({
          value: value,
          displayText: text,
          formatter: formatter,
          formula: formula,
          formulaInfo: formulaInfo,
        });
      }
      if (rowData.some(cell => cell.value !== null && cell.value !== undefined)) {
        allData.push(rowData);
      }
    }

    // 데이터 출력
    console.log('전체 행 수:', allData.length);
    allData.forEach((row, rowIndex) => {
      console.group(`행 ${rowIndex + 1}`);
      row.forEach((cell, colIndex) => {
        if (cell.value !== null && cell.value !== undefined) {
          console.log(`열 ${colIndex + 1}:`, {
            '실제 값': cell.value,
            '표시 텍스트': cell.displayText,
            '서식': cell.formatter,
            '수식': cell.formula,
            '서식 정보': cell.formulaInfo,
          });
        }
      });
      console.groupEnd();
    });

    // 선택된 범위의 데이터 가져오기
    const selections = sheet.getSelections();
    if (selections && selections.length > 0) {
      console.group('선택된 영역 데이터');
      selections.forEach((selection, index) => {
        const { row, col, rowCount, colCount } = selection;
        console.group(`선택 영역 ${index + 1}`);
        for (let r = row; r < row + rowCount; r++) {
          for (let c = col; c < col + colCount; c++) {
            const value = sheet.getValue(r, c);
            if (value !== null && value !== undefined) {
              console.log(`셀(${r},${c}):`, {
                '값': value,
                '표시 텍스트': sheet.getText(r, c)
              });
            }
          }
        }
        console.groupEnd();
      });
      console.groupEnd();
    }

    console.groupEnd();
    alert('콘솔에서 스프레드시트 데이터를 확인해주세요.');

  } catch (error) {
    console.error('데이터 확인 중 오류 발생:', error);
    alert('데이터 확인 중 오류가 발생했습니다: ' + error.message);
  } finally {
    stopLoading();
  }
};

// 날짜 포맷팅 함수
const formatDate = (date) => {
  if (!date) return '';
  if (typeof date === 'string' && date.includes('OADate')) {
    // OADate를 JavaScript Date로 변환
    date = fromOADate(parseFloat(date.match(/\d+\.?\d*/)[0]));
  }
  const d = new Date(date);
  return new Intl.DateTimeFormat('ko-KR', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).format(d);
};

// 시간 포맷팅 함수
const formatTime = (time) => {
  if (!time) return '';
  if (typeof time === 'string' && time.includes('OADate')) {
    time = fromOADate(parseFloat(time.match(/\d+\.?\d*/)[0]));
  }
  const d = new Date(time);
  return new Intl.DateTimeFormat('ko-KR', {
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  }).format(d);
};

// 통화 포맷팅 함수
const formatCurrency = (value) => {
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW'
  }).format(value);
};

// 백분율 포맷팅 함수
const formatPercentage = (value) => {
  return new Intl.NumberFormat('ko-KR', {
    style: 'percent',
    minimumFractionDigits: 2
  }).format(value);
};

// OADate를 JavaScript Date로 변환하는 함수
const fromOADate = (oaDate) => {
  // OADate의 시작일 (1899년 12월 30일)
  const dateStart = new Date(1899, 11, 30);
  // 일수를 밀리초로 변환
  const milliseconds = Math.round((oaDate) * 86400000);
  return new Date(dateStart.getTime() + milliseconds);
};

</script>

<style scoped>
.container {
  padding: 1rem;
}

.spread-host {
  width: 90vw;
  height: 80vh;
  border: 1px solid #ccc;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.check-button {
  margin-bottom: 1rem;
  padding: 0.5rem 1rem;
  background-color: #4472C4;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-weight: bold;
}

.check-button:hover {
  background-color: #365899;
}
</style>