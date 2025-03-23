<template>
  <div>
    <gc-spread-sheets class='spread-host' @workbookInitialized="initWorkbook">
    </gc-spread-sheets>
  </div>
</template>

<script setup>
import '@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import { GcSpreadSheets } from '@mescius/spread-sheets-vue'
import * as GC from "@mescius/spread-sheets"
import '@mescius/spread-sheets-resources-ko'

// SpreadJS 라이선싱
// var SpreadJSKey = "xxx";          // 라이선스 입력
// GC.Spread.Sheets.LicenseKey = SpreadJSKey;

GC.Spread.Common.CultureManager.culture("ko-kr");

const initWorkbook = (spread) => {
  let sheet = spread.getActiveSheet();

  // Initial value 바인딩 테스트
  setInitialValues(sheet);

  // style 적용
  setStylesForCells(sheet);

  // 데이터 바인딩
  // 새로운 시트를 추가합니다.
  spread.addSheet(1);

  // 추가한 시트를 가져옵니다.
  const sheet2 = spread.getSheet(1);

  // 데이터를 입력합니다.
  const person = {
    name: '홍길동',
    age: 25,
    gender: '남',
    address: {
      postcode: '10001'
    }
  };

  // 셀 바인딩 소스를 설정합니다.
  const source = new GC.Spread.Sheets.Bindings.CellBindingSource(person);

  // 지정된 시트 영역에 지정된 셀의 바인딩 경로를 설정합니다.
  sheet2.setBindingPath(2, 2, 'name');
  sheet2.setBindingPath(3, 2, 'age');
  sheet2.setBindingPath(4, 2, 'gender');
  sheet2.setBindingPath(5, 2, 'address.postcode');

  // 시트의 데이터 소스를 설정합니다.
  sheet2.setDataSource(source);
}

/**
 * SpreadJS 데이터 바인딩
 */
const setInitialValues = (sheet) => {

  // 제목 설정
  sheet.setValue(0, 0, "데이터 타입 테스트"); // 값 설정

  // 1. 기본 텍스트
  sheet.setValue(1, 0, "Text");
  sheet.setValue(1, 1, "일반 텍스트");

  // 2. 숫자
  sheet.setValue(2, 0, "Number");
  sheet.setValue(2, 1, 12345);
  sheet.getCell(2, 1).formatter("0.00"); // 소수점 2자리

  // 3. 통화
  sheet.setValue(3, 0, "Currency");
  sheet.setValue(3, 1, 50000);
  sheet.getCell(3, 1).formatter("₩#,##0"); // 원화 표시

  // 4. 퍼센트
  sheet.setValue(4, 0, "Percentage");
  sheet.setValue(4, 1, 0.75);
  sheet.getCell(4, 1).formatter("0%"); // 퍼센트 표시

  // 5. 날짜
  sheet.setValue(5, 0, "Date");
  sheet.setValue(5, 1, new Date());
  sheet.getCell(5, 1).formatter("yyyy-mm-dd");

  // 6. 시간
  sheet.setValue(6, 0, "Time");
  sheet.setValue(6, 1, new Date());
  sheet.getCell(6, 1).formatter("hh:mm:ss");

  // 7. 날짜/시간
  sheet.setValue(7, 0, "DateTime");
  sheet.setValue(7, 1, new Date());
  sheet.getCell(7, 1).formatter("yyyy-mm-dd hh:mm:ss"); // 셀 포맷 설정

  // 8. 불리언
  sheet.setValue(8, 0, "Boolean");
  sheet.setValue(8, 1, true);

  // 9. 공식
  sheet.setValue(9, 0, "Formula");
  sheet.setFormula(9, 1, "=SUM(B2:B3)"); // 수식 설정

  // 10. 기본 하이퍼링크
  sheet.setValue(10, 0, "기본 링크");
  sheet.setHyperlink(10, 1, {
    url: 'https://naver.com',
    tooltip: 'grapecity',
    linkColor: '#0066cc',
    visitedLinkColor: '#3399ff',
    target: GC.Spread.Sheets.Hyperlink.HyperlinkTargetType.blank
  }, GC.Spread.Sheets.SheetArea.viewport);
  // If you want to click the hyperlink you set, don't forget to add a value for the cell.
  sheet.setText(10, 1, '네이버 바로가기');

  // 컬럼 너비 자동 조정
  sheet.autoFitColumn(0);
  sheet.autoFitColumn(1);

};

const setStylesForCells = (sheet) => {

  // 제목 스타일링
  const titleCell = sheet.getCell(0, 0);
  titleCell.backColor("#4472C4")
      .foreColor("white")
      .font("bold 14pt 맑은 고림")
      .hAlign(GC.Spread.Sheets.HorizontalAlign.center);
  sheet.setRowHeight(0, 35);

  // 왼쪽 레이블 컬럼 스타일링
  const labelRange = sheet.getRange(1, 0, 10, 1);
  labelRange.backColor("#D9E1F2")
      .foreColor("#2F528F")
      .font("bold 11pt 맑은 고림");

  // 데이터 셀 스타일링
  for (let i = 1; i <= 10; i++) {
    const cell = sheet.getCell(i, 1);

    // 공통 스타일
    cell.borderLeft(new GC.Spread.Sheets.LineBorder("#8EA9DB"))
        .borderRight(new GC.Spread.Sheets.LineBorder("#8EA9DB"))
        .borderTop(new GC.Spread.Sheets.LineBorder("#8EA9DB"))
        .borderBottom(new GC.Spread.Sheets.LineBorder("#8EA9DB"));

    // 데이터 타입별 특별 스타일
    switch(i) {
      case 2: // Number
        cell.backColor("#E2EFD9");
        break;
      case 3: // Currency
        cell.backColor("#E2EFD9")
            .font("bold 11pt 맑은 고림");
        break;
      case 4: // Percentage
        cell.backColor("#E2EFD9");
        break;
      case 5: // Date
      case 6: // Time
      case 7: // DateTime
        cell.backColor("#FFF2CC");
        break;
      case 8: // Boolean
        cell.backColor("#FCE4D6");
        break;
      case 9: // Formula
        cell.backColor("#DEEBF7")
            .font("italic 11pt 맑은 고림");
        break;
      case 10: // Hyperlink
        cell.textDecoration(GC.Spread.Sheets.TextDecorationType.underline);
        break;
    }
  }

  // 전체 테두리 설정
  sheet.getRange(0, 0, 11, 2).setBorder(
      new GC.Spread.Sheets.LineBorder("#8EA9DB", GC.Spread.Sheets.LineStyle.thin),
      { all: true }
  );

  // 열 너비 조정
  sheet.setColumnWidth(0, 220); // 첫 번째 열
  sheet.setColumnWidth(1, 200); // 두 번째 열
}


</script>

<style scoped>
.spread-host {
  width: 90vw;
  height: 80vh;
}
</style>