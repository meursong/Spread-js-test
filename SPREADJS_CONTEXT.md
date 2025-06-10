# SpreadJS 컨텍스트 문서

## 개요
SpreadJS는 Mescius(구 GrapeCity)에서 개발한 JavaScript 기반의 스프레드시트 솔루션으로, Excel과 유사한 기능을 웹 애플리케이션에서 구현할 수 있는 라이브러리입니다.

## 주요 특징

### 1. 핵심 기능
- **Excel 호환성**: 완전한 Excel 파일(.xlsx) 가져오기/내보내기
- **500+ 함수**: 19개 언어로 500개 이상의 내장 함수 지원
- **고성능 계산 엔진**: 대용량 데이터 처리 최적화
- **차트 및 시각화**: 30+ 종류의 차트 및 스파크라인
- **데이터 바인딩**: 객체와 배열 데이터 실시간 바인딩

### 2. 고급 기능
- **조건부 서식**: 동적 셀 스타일링
- **피벗 테이블**: 데이터 분석 및 요약
- **바코드 지원**: 다양한 바코드 타입 생성
- **PDF 내보내기**: 직접 PDF 변환 기능
- **Camera Shape**: 실시간 스냅샷 뷰
- **AI 어시스턴트**: 자연어로 수식 생성 (개발자 프리뷰)

## API 주요 네임스페이스

### 1. 핵심 네임스페이스
- `GC.Spread.Sheets`: 기본 스프레드시트 작업
- `GC.Spread.Commands`: 실행취소/재실행 유틸리티
- `GC.Spread.Excel`: 클라이언트 측 Excel 가져오기/내보내기

### 2. 고급 기능 네임스페이스
- `GC.Pivot`: 데이터 피벗 및 분석
- `GC.Spread.CalcEngine`: 계산 유틸리티
- `GC.Spread.Sheets.CellTypes`: 사용자 정의 셀 타입 관리
- `GC.Spread.Sheets.Charts`: 차트 생성 및 조작
- `GC.Spread.Sheets.Barcode`: 바코드 통합
- `GC.Spread.Sheets.DataValidation`: 입력 유효성 검사
- `GC.Spread.Sheets.Slicers`: 데이터 필터링
- `GC.Spread.Sheets.Tables`: 테이블 관리

## 현재 프로젝트 구현 분석

### 1. 사용 중인 패키지 (v18.0.5)
```json
{
  "@mescius/spread-sheets": "^18.0.5",
  "@mescius/spread-sheets-charts": "^18.0.5",
  "@mescius/spread-sheets-io": "^18.0.5",
  "@mescius/spread-sheets-resources-ko": "^18.0.2",
  "@mescius/spread-sheets-shapes": "^18.0.5",
  "@mescius/spread-sheets-vue": "^18.0.5"
}
```

### 2. 핵심 컴포넌트 구조

#### SpreadSheet.vue (기본 스프레드시트)
- **목적**: 데이터 타입 테스트 및 기본 기능 시연
- **주요 기능**:
  - 다양한 데이터 타입 (텍스트, 숫자, 통화, 퍼센트, 날짜, 시간)
  - 수식 및 하이퍼링크
  - 데이터 바인딩 (`CellBindingSource`)
  - 스타일링 및 포맷팅

#### ChartAndGraphSheet.vue (차트 및 그래프)
- **목적**: 차트 및 스파크라인 기능 시연
- **지원 차트**: 막대, 선, 파이, 영역, 분산형
- **스파크라인**: 라인, 컬럼, 승패 타입
- **데이터 소스**: 분기별 제품 판매량 데이터

#### ExcelTestSheet.vue (Excel 통합)
- **목적**: Excel 파일 가져오기/내보내기
- **기능**:
  - 파일 업로드 및 가져오기
  - 데이터 수정 및 편집
  - Excel 형식으로 내보내기
  - 동적 행 추가 및 스파크라인 생성

### 3. 구현 패턴

#### 초기화 패턴
```javascript
const initWorkbook = (spread) => {
  startLoading();
  try {
    spreadRef.value = spread;
    const sheet = spread.getActiveSheet();
    // 시트 초기화 로직
  } finally {
    stopLoading();
  }
};
```

#### 데이터 바인딩 패턴
```javascript
const source = new GC.Spread.Sheets.Bindings.CellBindingSource(data);
sheet.setBindingPath(row, col, 'property.path');
sheet.setDataSource(source);
```

#### 이벤트 처리 패턴
```javascript
@workbookInitialized="initWorkbook"
@selectionChanged="onSelectionChanged"
@valueChanged="onValueChanged"
```

### 4. 스타일링 및 포맷팅

#### 셀 포맷터
```javascript
sheet.getCell(row, col).formatter("₩#,##0");     // 통화
sheet.getCell(row, col).formatter("0%");         // 퍼센트
sheet.getCell(row, col).formatter("yyyy-mm-dd"); // 날짜
```

#### 스타일 적용
```javascript
const style = new GC.Spread.Sheets.Style();
style.font = "맑은 고딕 11pt bold";
style.backColor = "#EEEEEE";
style.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
sheet.setStyle(row, col, style);
```

### 5. 차트 생성

#### 기본 차트
```javascript
const chart = sheet.charts.add(
  'Chart_1',
  GC.Spread.Sheets.Charts.ChartType.columnClustered,
  x, y, width, height,
  dataRange, categoriesRange
);
```

#### 스파크라인
```javascript
sheet.setSparkline(row, col, {
  dataRange: "B2:B5",
  sparklineType: GC.Spread.Sheets.Sparklines.SparklineType.line,
  setting: {
    showMarkers: true,
    displayXAxis: true,
    lineWeight: 2
  }
});
```

## 개발 가이드라인

### 1. 초기화 및 설정
```javascript
// 한국어 지역화
GC.Spread.Common.CultureManager.culture("ko-kr");

// 라이선스 설정 (필요시)
// GC.Spread.Sheets.LicenseKey = "your-license-key";
```

### 2. Vue 통합
```vue
<template>
  <gc-spread-sheets 
    class="spread-host" 
    @workbookInitialized="initWorkbook">
  </gc-spread-sheets>
</template>
```

### 3. 상태 관리
- `ref()`를 사용한 spread 인스턴스 관리
- 반응형 상태로 선택 영역 및 포맷 정보 추적
- 로딩 상태 관리를 위한 composable 사용

### 4. 에러 처리
- try-catch 블록을 통한 초기화 보호
- null 체크를 통한 인스턴스 가용성 확인
- 로딩 상태를 통한 사용자 피드백

## 성능 최적화

### 1. 메모리 관리
- 적절한 컴포넌트 라이프사이클 관리
- 이벤트 리스너 정리
- 대용량 데이터 처리 시 청킹

### 2. 렌더링 최적화
- 필요시에만 스프레드시트 표시
- 동적 로딩 패턴 활용
- 셀 업데이트 배치 처리

## 라이선스 및 배포

### 1. 라이선스 모델
- 개발자별 라이선스
- 호스트명별 배포 라이선스
- 상업적/비상업적 용도에 따른 차등 가격

### 2. 프레임워크 지원
- Angular, React, Vue 지원
- AngularJS, Breeze, Knockout 호환
- TypeScript 지원

## 문제 해결

### 1. 일반적인 이슈
- SSR 환경에서 `ssr: false` 설정 필요
- CommonJS 모듈 변환 설정
- 메모리 누수 방지를 위한 적절한 정리

### 2. Nuxt 3 특정 설정
```javascript
// nuxt.config.ts
export default defineNuxtConfig({
  ssr: false,
  build: {
    transpile: ['@mescius/spread-sheets-vue']
  }
})
```

## 추가 리소스

- [공식 문서](https://developer.mescius.com/spreadjs/docs/)
- [API 레퍼런스](https://developer.mescius.com/spreadjs/docs/API-documentation)
- [시작 가이드](https://developer.mescius.com/spreadjs/docs/getstarted)
- [30일 무료 체험](https://developer.mescius.com/spreadjs/pricing)

이 문서는 현재 프로젝트의 SpreadJS 구현을 이해하고 확장하는 데 필요한 모든 컨텍스트를 제공합니다.