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

    <!-- 차트 컨트롤 추가 -->
    <div class="chart-controls">
      <div class="control-group">
        <button class="chart-btn" @click="addColumnChart">막대 차트</button>
        <button class="chart-btn" @click="addLineChart">선 차트</button>
        <button class="chart-btn" @click="addPieChart">파이 차트</button>
        <button class="chart-btn" @click="addAreaChart">영역 차트</button>
        <button class="chart-btn" @click="addScatterChart">분산형 차트</button>
      </div>
      <div class="control-group">
        <button class="chart-btn" @click="addSparkline('line')">라인 스파크라인</button>
        <button class="chart-btn" @click="addSparkline('column')">컬럼 스파크라인</button>
        <button class="chart-btn" @click="addSparkline('winloss')">승패 스파크라인</button>
      </div>
      <div class="control-group">
        <button class="chart-btn" @click="increaseChartHeight">차트 높이 증가</button>
        <button class="chart-btn" @click="decreaseChartHeight">차트 높이 감소</button>
      </div>
    </div>

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
import { ref, inject, onMounted } from 'vue';
import '@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css';
import { GcSpreadSheets } from '@mescius/spread-sheets-vue';
import * as GC from "@mescius/spread-sheets";
import '@mescius/spread-sheets-resources-ko';

// 차트 라이브러리 임포트
import '@mescius/spread-sheets-charts';
import '@mescius/spread-sheets-shapes';

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

    setupTimetableHeader(sheet);
    drawTimetable(sheet);

  } catch (error) {
    console.error('스프레드시트 초기화 중 오류:', error);
  } finally {
    stopLoading();
  }
};


// ------------------------------------------------------------------------------------------------------------------------------------------------
// ------------------------------------ 시간표 영역 START ------------------------------------------------------------------------
function setupTimetableHeader(sheet) {
  // 1. 전체 제목 병합 및 설정
  sheet.addSpan(0, 0, 1, 21); // 5요일 × 4교시 + 1 (시간/번호 열)
  sheet.setValue(0, 0, "2020년도 2학기 시간표");

  // 2. '요일' 행 설정
  const days = ["월", "화", "수", "목", "금"];
  for (let i = 0; i < days.length; i++) {
    const startCol = 1 + i * 4; // 교시당 4칸씩
    sheet.addSpan(1, startCol, 1, 4); // 요일 셀 병합
    sheet.setValue(1, startCol, days[i]);
  }

  // 3. '교시 번호' 행 설정
  let col = 1;
  for (let i = 0; i < 5; i++) { // 5요일
    for (let j = 1; j <= 4; j++) { // 교시
      sheet.setValue(2, col, j.toString());
      col++;
    }
  }

  // 4. '시간/번호' 열 고정 텍스트
  sheet.addSpan(1, 0, 2, 1);
  sheet.setValue(1, 0, "시간/번호");

  // 5. 스타일 적용
  sheet.getRange(0, 0, 3, 21).hAlign(GC.Spread.Sheets.HorizontalAlign.center)
      .vAlign(GC.Spread.Sheets.VerticalAlign.center)
      .setBorder(new GC.Spread.Sheets.LineBorder("black", GC.Spread.Sheets.LineStyle.thin), { all: true });

  // 열 너비 설정
  for (let i = 0; i < 21; i++) {
    sheet.setColumnWidth(i, 60);
  }

  // 행 높이 설정
  sheet.setRowHeight(0, 30);
  sheet.setRowHeight(1, 30);
  sheet.setRowHeight(2, 25);
}

const timetableData = [
  // 월요일
  { day: "월", start: 2, span: 3, text: "Open Lab. (B반)\n과102" },
  { day: "월", start: 6, span: 2, text: "약리학\n(오창근)\n과210호" },
  { day: "월", start: 8, span: 3, text: "임상간호술기 (A)\n(이재욱)\n과102호" },

  // 화요일
  { day: "화", start: 2, span: 2, text: "기본간호학2\n(김수현)\n과201호" },
  { day: "화", start: 4, span: 2, text: "여성건강간호학\n(권순자)\n본관203호" },
  { day: "화", start: 6, span: 2, text: "Open Lab. (C반)\n과102호" },

  // 수요일
  { day: "수", start: 2, span: 2, text: "기본간호학실습\n(2C반)\n과204호" },
  { day: "수", start: 6, span: 3, text: "성인간호학실습\n(이영주)\n본203호" }
];

const dayMap = { "월": 1, "화": 2, "수": 3, "목": 4, "금": 5 };

function drawTimetable(sheet) {
  // 헤더 영역은 0~2행을 이미 사용하므로, 3행부터 시작
  const headerRowCount = 3;

  const days = ["시간/요일", "월", "화", "수", "목", "금"];
  const times = [
    "1교시(09:00-09:50)", "2교시(09:50-10:40)", "3교시(10:40-11:30)",
    "4교시(11:30-12:30)", "5교시(12:30-13:20)", "6교시(13:20-14:10)",
    "7교시(14:10-15:00)", "8교시(15:00-15:50)", "9교시(15:50-16:40)",
    "10교시(16:50-17:40)", "11교시(17:40-18:30)"
  ];

  // 시간 정보만 설정 (3행부터)
  for (let row = 0; row < times.length; row++) {
    sheet.setValue(row + headerRowCount, 0, times[row]);
    sheet.setRowHeight(row + headerRowCount, 50);

    // 시간 셀에도 자동 줄바꿈 및 중앙 정렬 적용
    const timeCell = sheet.getCell(row + headerRowCount, 0);
    timeCell.wordWrap(true);
    timeCell.hAlign(GC.Spread.Sheets.HorizontalAlign.center);
    timeCell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
  }

  // 데이터 입력 - 오프셋을 적용하여 시작 행 조정
  timetableData.forEach(item => {
    const col = dayMap[item.day];
    const row = item.start + headerRowCount; // 헤더 행 수만큼 오프셋 추가
    const span = item.span;
    sheet.addSpan(row, col, span, 1);
    sheet.setValue(row, col, item.text);

    // 각 데이터 셀에 대해 자동 줄바꿈 및 중앙 정렬 적용
    const cell = sheet.getCell(row, col);
    cell.wordWrap(true);
    cell.hAlign(GC.Spread.Sheets.HorizontalAlign.center);
    cell.vAlign(GC.Spread.Sheets.VerticalAlign.center);
  });

  // 모든 셀에 대한 기본 스타일 적용 (테두리, 정렬, 자동 줄바꿈)
  const range = sheet.getRange(headerRowCount, 0, times.length, days.length);
  range.setBorder(new GC.Spread.Sheets.LineBorder("black", GC.Spread.Sheets.LineStyle.thin), { all: true })
      .wordWrap(true)
      .hAlign(GC.Spread.Sheets.HorizontalAlign.center)
      .vAlign(GC.Spread.Sheets.VerticalAlign.center);

  // 헤더 영역에도 동일한 스타일 적용 (이미 setupTimetableHeader에서 설정했을 수도 있지만 확실히 하기 위해)
  sheet.getRange(0, 0, headerRowCount, days.length)
      .wordWrap(true)
      .hAlign(GC.Spread.Sheets.HorizontalAlign.center)
      .vAlign(GC.Spread.Sheets.VerticalAlign.center);

  console.log("시간표가 생성되었습니다.");
}

// ------------------------------------------------------------------------------------------------------------------------------------------------
// ------------------------------------ 시간표 영역 END ------------------------------------------------------------------------


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

const createTimetable = (sheet) => {
  const days = ["시간", "월", "화", "수", "목", "금"];
  const times = ["1교시", "2교시", "3교시", "4교시", "5교시"];

  // 요일 헤더 세팅
  days.forEach((day, col) => {
    sheet.setValue(0, col, day);
    sheet.setColumnWidth(col, 100);
  });

  // 시간 및 과목 세팅
  times.forEach((time, rowIndex) => {
    const row = rowIndex + 1;
    sheet.setValue(row, 0, time); // 시간 열

    // 요일별 과목 데이터 입력 (예시)
    const timetableRow = ["자료구조", "", "운영체제", "", ""]; // 월~금
    for (let col = 1; col <= 5; col++) {
      sheet.setValue(row, col, timetableRow[col - 1]);
    }
  });

  // 셀 정렬 및 테두리 설정
  sheet.getRange(0, 0, 6, 6).hAlign(GC.Spread.Sheets.HorizontalAlign.center);
  sheet.getRange(0, 0, 6, 6).vAlign(GC.Spread.Sheets.VerticalAlign.center);
  sheet.getRange(0, 0, 6, 6).setBorder(
      new GC.Spread.Sheets.LineBorder("black", GC.Spread.Sheets.LineStyle.thin),
      { all: true }
  );

  console.log("시간표가 생성되었습니다.");
};





// 차트 추가 함수들
const addColumnChart = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const chartCount = sheet.charts.all().length;

  // 데이터 범위 설정
  const dataRange = "Sheet1!A1:D5"; // 또는 적절한 범위 문자열

  const categoriesRange = new GC.Spread.Sheets.Range(1, 0, 4, 1);

  // 차트 생성
  const chart = sheet.charts.add(`차트_${chartCount + 1}`,
      GC.Spread.Sheets.Charts.ChartType.columnClustered,
      200, 20, 400, 300, // x, y, width, height
      dataRange, categoriesRange);

  // 차트 제목 설정
  chart.title({
    text: "분기별 제품 판매량 (막대 차트)",
    fontSize: 14,
    fontFamily: "맑은 고딕"
  });

  // 범례 설정
  chart.legend({
    position: "bottom"
  });

  console.log('막대 차트가 추가되었습니다.');
};

const addLineChart = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const chartCount = sheet.charts.all().length;

  // 데이터 범위 설정
  const dataRange = new GC.Spread.Sheets.Range(1, 1, 4, 3);
  const categoriesRange = new GC.Spread.Sheets.Range(1, 0, 4, 1);

  // 차트 생성
  const chart = sheet.charts.add(`차트_${chartCount + 1}`,
      GC.Spread.Sheets.Charts.ChartType.line,
      200, 350, 400, 300, // x, y, width, height
      dataRange, categoriesRange);

  // 차트 제목 설정
  chart.title({
    text: "분기별 제품 판매량 (선 차트)",
    fontSize: 14,
    fontFamily: "맑은 고딕"
  });

  // 범례 설정
  chart.legend({
    position: "bottom"
  });

  console.log('선 차트가 추가되었습니다.');
};

const addPieChart = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const chartCount = sheet.charts.all().length;

  // 데이터 범위 설정 (파이차트는 한 시리즈만 가능)
  const dataRange = new GC.Spread.Sheets.Range(1, 1, 4, 1);
  const categoriesRange = new GC.Spread.Sheets.Range(1, 0, 4, 1);

  // 차트 생성
  const chart = sheet.charts.add(`차트_${chartCount + 1}`,
      GC.Spread.Sheets.Charts.ChartType.pie,
      650, 20, 300, 300, // x, y, width, height
      dataRange, categoriesRange);

  // 차트 제목 설정
  chart.title({
    text: "제품 A 분기별 판매량 (파이 차트)",
    fontSize: 14,
    fontFamily: "맑은 고딕"
  });

  // 범례 설정
  chart.legend({
    position: "right"
  });

  console.log('파이 차트가 추가되었습니다.');
};

const addAreaChart = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const chartCount = sheet.charts.all().length;

  // 데이터 범위 설정
  const dataRange = new GC.Spread.Sheets.Range(1, 1, 4, 3);
  const categoriesRange = new GC.Spread.Sheets.Range(1, 0, 4, 1);

  // 차트 생성
  const chart = sheet.charts.add(`차트_${chartCount + 1}`,
      GC.Spread.Sheets.Charts.ChartType.area,
      650, 350, 400, 300, // x, y, width, height
      dataRange, categoriesRange);

  // 차트 제목 설정
  chart.title({
    text: "분기별 제품 판매량 (영역 차트)",
    fontSize: 14,
    fontFamily: "맑은 고딕"
  });

  // 범례 설정
  chart.legend({
    position: "bottom"
  });

  console.log('영역 차트가 추가되었습니다.');
};

const addScatterChart = () => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();
  const chartCount = sheet.charts.all().length;

  // 데이터 범위 설정
  const dataRange = new GC.Spread.Sheets.Range(1, 1, 4, 2);

  // 차트 생성 (분산형은 카테고리가 필요없음)
  const chart = sheet.charts.add(`차트_${chartCount + 1}`,
      GC.Spread.Sheets.Charts.ChartType.scatter,
      200, 670, 400, 300, // x, y, width, height
      dataRange);

  // 차트 제목 설정
  chart.title({
    text: "제품 A와 제품 B의 상관관계 (분산형)",
    fontSize: 14,
    fontFamily: "맑은 고딕"
  });

  // 범례 설정
  chart.legend({
    position: "bottom"
  });

  console.log('분산형 차트가 추가되었습니다.');
};

// 스파크라인 추가 함수
const addSparkline = (type) => {
  if (!spreadRef.value) return;

  const sheet = spreadRef.value.getActiveSheet();

  // 스파크라인 데이터 추가
  sheet.setValue(6, 0, "스파크라인:");

  // 스파크라인 자료
  sheet.setValue(8, 0, "제품 A");
  sheet.setValue(9, 0, "제품 B");
  sheet.setValue(10, 0, "제품 C");

  // 스파크라인 생성
  const sparklineType = GC.Spread.Sheets.Sparklines.SparklineType[type];

  // 제품 A 스파크라인
  sheet.setSparkline(8, 1, {
    dataRange: "B2:B5", // 1분기~4분기 제품 A 데이터
    sparklineType: sparklineType,
    setting: {
      showMarkers: true,
      displayXAxis: true,
      lineWeight: 2
    }
  });

  // 제품 B 스파크라인
  sheet.setSparkline(9, 1, {
    dataRange: "C2:C5", // 1분기~4분기 제품 B 데이터
    sparklineType: sparklineType,
    setting: {
      showMarkers: true,
      displayXAxis: true,
      lineWeight: 2
    }
  });

  // 제품 C 스파크라인
  sheet.setSparkline(10, 1, {
    dataRange: "D2:D5", // 1분기~4분기 제품 C 데이터
    sparklineType: sparklineType,
    setting: {
      showMarkers: true,
      displayXAxis: true,
      lineWeight: 2
    }
  });

  // 스타일 적용
  const headerStyle = new GC.Spread.Sheets.Style();
  headerStyle.font = new GC.Spread.Sheets.Font("맑은 고딕", 11, "bold");

  sheet.setStyle(6, 0, headerStyle);
  sheet.setStyle(8, 0, headerStyle);
  sheet.setStyle(9, 0, headerStyle);
  sheet.setStyle(10, 0, headerStyle);
  sheet.setColumnWidth(1, 200);

  console.log(`${type} 스파크라인이 추가되었습니다.`);
};

</script>

<style scoped>
.container {
  display: flex;
  flex-direction: column;
  height: 100%;
}

.chart-controls {
  display: flex;
  padding: 8px;
  background-color: #f5f5f5;
  border-bottom: 1px solid #ddd;
  gap: 10px;
}

.control-group {
  display: flex;
  gap: 4px;
}

.chart-btn {
  padding: 6px 12px;
  background-color: #4285f4;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-size: 13px;
  font-weight: 500;
}

.chart-btn:hover {
  background-color: #3275e4;
}

.relative {
  position: relative;
  flex: 1;
  min-height: 400px;
}

.spread-host {
  height: 600px !important; /* 원하는 높이로 조정 (기본값보다 크게) */
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
}
</style>