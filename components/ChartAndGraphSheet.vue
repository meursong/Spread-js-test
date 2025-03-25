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

// 샘플 데이터 생성 함수
const createSampleData = (sheet) => {
  // 헤더 설정
  sheet.setValue(0, 0, "분기");
  sheet.setValue(0, 1, "제품 A");
  sheet.setValue(0, 2, "제품 B");
  sheet.setValue(0, 3, "제품 C");

  // 데이터 설정
  sheet.setValue(1, 0, "1분기");
  sheet.setValue(1, 1, 250);
  sheet.setValue(1, 2, 180);
  sheet.setValue(1, 3, 120);

  sheet.setValue(2, 0, "2분기");
  sheet.setValue(2, 1, 300);
  sheet.setValue(2, 2, 200);
  sheet.setValue(2, 3, 140);

  sheet.setValue(3, 0, "3분기");
  sheet.setValue(3, 1, 280);
  sheet.setValue(3, 2, 250);
  sheet.setValue(3, 3, 160);

  sheet.setValue(4, 0, "4분기");
  sheet.setValue(4, 1, 350);
  sheet.setValue(4, 2, 280);
  sheet.setValue(4, 3, 200);

  // 헤더 스타일 적용
  const headerStyle = new GC.Spread.Sheets.Style();
  headerStyle.font = "맑은 고딕 11pt bold";
  headerStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
  headerStyle.backColor = "#EEEEEE";

  sheet.setRowHeight(0, 25);
  for (let col = 0; col < 4; col++) {
    sheet.setStyle(0, col, headerStyle);
    sheet.setColumnWidth(col, 100);
  }
}

// 워크북 초기화
const initWorkbook = (spread) => {
  startLoading();
  try {
    // spread 객체 저장
    spreadRef.value = spread;

    const sheet = spread.getActiveSheet();

    // 샘플 데이터 생성
    createSampleData(sheet);

    // 기본 차트 추가 (예시)
    addColumnChart();

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
  position: absolute;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
}
</style>