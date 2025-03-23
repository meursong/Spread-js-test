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


  // // 스타일 설정
  // // B열, C열의 너비를 200으로 설정합니다.
  // sheet.setColumnWidth(1, 200);
  // sheet.setColumnWidth(2, 200);
  // // B2:C2의 배경색과 글자색을 설정합니다.
  // sheet.getRange(1, 1, 1, 2).backColor("rgb(130, 188, 0)").foreColor("rgb(255, 255, 255)");
  // // B4:C4의 배경색을 설정합니다.
  // sheet.getRange(3, 1, 1, 2).backColor("rgb(211, 211, 211)");
  // // B2:C2의 셀을 병합합니다.
  // sheet.addSpan(1, 1, 1, 2);
  // // 각 범위에 테두리를 설정합니다.
  // sheet.getRange(1, 1, 4, 2).setBorder(
  //     new GC.Spread.Sheets.LineBorder("Black", GC.Spread.Sheets.LineStyle.thin),
  //     { all: true }
  // );
  // sheet.getRange(1, 1, 4, 2).setBorder(
  //     new GC.Spread.Sheets.LineBorder("Black", GC.Spread.Sheets.LineStyle.dotted),
  //     { inside: true }
  // );
  // // B2:C2의 병합된 셀에 수직 가운데 정렬을 설정합니다.
  // sheet.getRange(1, 1, 1, 2).hAlign(GC.Spread.Sheets.HorizontalAlign.center);

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
 *
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

</script>

<style scoped>
.spread-host {
  width: 90vw;
  height: 80vh;
}
</style>