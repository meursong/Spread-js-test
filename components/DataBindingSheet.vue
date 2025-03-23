<template>
  <div class="container">
    <button class="check-button" @click="checkData">바인딩 데이터 확인</button>
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

// 데이터 모델
const data = ref({
  text: '일반 텍스트입니다',
  number: 12345.67,
  date: new Date(),
  boolean: true,
  currency: 50000,
  percentage: 0.75,
  time: new Date(),
  calculation: 0,
  dropdown: 'KR', // value 값으로 저장됨
  hyperlink: 'https://www.example.com'
});

// 헤더 정보
const headers = [
  {
    label: '텍스트',
    field: 'text',
    type: 'text'
  },
  {
    label: '숫자',
    field: 'number',
    type: 'number',
    formatter: '#,##0.00'
  },
  {
    label: '날짜',
    field: 'date',
    type: 'date',
    formatter: 'yyyy-MM-dd'
  },
  {
    label: '불리언',
    field: 'boolean',
    type: 'boolean'
  },
  {
    label: '통화',
    field: 'currency',
    type: 'currency',
    formatter: '₩#,##0'
  },
  {
    label: '백분율',
    field: 'percentage',
    type: 'percentage',
    formatter: '0.00%'
  },
  {
    label: '시간',
    field: 'time',
    type: 'time',
    formatter: 'HH:mm:ss'
  },
  {
    label: '계산식',
    field: 'calculation',
    type: 'formula',
    formula: '=B2+B5' // 숫자와 통화 값의 합
  },
  {
    label: '드롭다운',
    field: 'dropdown',
    type: 'dropdown',
    items: [
      { value: 'KR', text: '대한민국' },
      { value: 'US', text: '미국' },
      { value: 'JP', text: '일본' },
      { value: 'CN', text: '중국' }
    ]
  },
  {
    label: '하이퍼링크',
    field: 'hyperlink',
    type: 'hyperlink'
  }
];

// 셀 서식 및 유효성 검사 설정
const setCellFormatting = (sheet, header, rowIndex) => {
  const dataCell = sheet.getCell(rowIndex, 1);

  // 기본 스타일 적용
  dataCell.font("11pt 맑은 고딕");

  switch (header.type) {
    case 'number':
    case 'currency':
    case 'percentage':
      dataCell.formatter(header.formatter);
      break;

    case 'date':
    case 'time':
      dataCell.formatter(header.formatter);
      break;

    case 'boolean':
      // 체크박스 스타일 적용
      sheet.setCellType(rowIndex, 1, new GC.Spread.Sheets.CellTypes.CheckBox());
      break;

    case 'dropdown':
      const dropdownList = new GC.Spread.Sheets.CellTypes.ComboBox();
      // text와 value를 매핑한 객체 생성
      const dropdownValues = {};
      header.items.forEach(item => {
        dropdownValues[item.text] = item.value;
      });
      // 표시될 텍스트만 items로 설정
      dropdownList.items(header.items.map(item => item.text));
      sheet.setCellType(rowIndex, 1, dropdownList);
      break;


    case 'hyperlink':
      // 하이퍼링크 설정
      sheet.setHyperlink(rowIndex, 1, {
        url: data.value[header.field],
        tooltip: '링크 클릭',
        linkColor: '#0066cc',
        visitedLinkColor: '#3399ff'
      });
      break;
  }
};

// 스타일 설정 함수 수정
const setStyles = (sheet) => {
  headers.forEach((header, index) => {
    // 헤더 셀 스타일링
    const headerCell = sheet.getCell(index, 0);
    headerCell.backColor("#4472C4")
        .foreColor("white")
        .font("bold 12pt 맑은 고딕")
        .hAlign(GC.Spread.Sheets.HorizontalAlign.center);

    // 데이터 셀 개별 서식 설정
    setCellFormatting(sheet, header, index);
  });

  // 컬럼 너비 자동 조정
  sheet.autoFitColumn(0);
  sheet.autoFitColumn(1);
};

// 워크북 초기화
const initWorkbook = (spread) => {
  startLoading();
  try {
    const sheet = spread.getActiveSheet();

    // 데이터 바인딩 설정
    const source = new GC.Spread.Sheets.Bindings.CellBindingSource(data.value);

    // 헤더와 바인딩 설정
    headers.forEach((header, index) => {
      sheet.setValue(index, 0, header.label); // 헤더 라벨 설정
      sheet.setBindingPath(index, 1, header.field); // 데이터 바인딩
    });

    // 데이터 소스 설정
    sheet.setDataSource(source);

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
  const message = Object.entries(data.value)
    .map(([key, value]) => {
      // 헤더 정보 찾기
      const header = headers.find(h => h.field === key);
      if (!header) return `${key}: ${value}`;

      // 데이터 타입에 따른 포맷팅
      switch (header.type) {
        case 'date':
          return `${header.label}: ${formatDate(value)}`;
        case 'time':
          return `${header.label}: ${formatTime(value)}`;
        case 'currency':
          return `${header.label}: ${formatCurrency(value)}`;
        case 'percentage':
          return `${header.label}: ${formatPercentage(value)}`;
        default:
          return `${header.label}: ${value}`;
      }
    })
    .join('\n');

  alert(message);
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