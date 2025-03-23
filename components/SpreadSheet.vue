<template>
    <div>
        <gc-spread-sheets Class='spread-host' @workbookInitialized="initWorkbook">
        </gc-spread-sheets>
    </div>
</template>
    
<script>
    import '@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
    import { GcSpreadSheets } from '@mescius/spread-sheets-vue'
    import * as GC from "@mescius/spread-sheets";
    import '@mescius/spread-sheets-resources-ko';

    // SpreadJS 라이선싱
    // var SpreadJSKey = "xxx";          // 라이선스 입력
    // GC.Spread.Sheets.LicenseKey = SpreadJSKey;

    GC.Spread.Common.CultureManager.culture("ko-kr");

    export default {
        components: {
            GcSpreadSheets
        },
        methods: {
            initWorkbook(spread) {
                let sheet = spread.getActiveSheet();

                sheet.setValue(1, 1, "값 설정하기");
                // 값 설정 - Number : B3에 "Number" 라는 텍스트를, C3에 23이라는 숫자를 삽입합니다.
                sheet.setValue(2, 1, "Number");
                sheet.setValue(2, 2, 23);
                // 값 설정 - Text : B4에 "Text" 라는 텍스트를, C4에 "MESCIUS"라는 텍스트를 삽입합니다.
                sheet.setValue(3, 1, "Text");
                sheet.setValue(3, 2, "MESCIUS");
                // 값 설정 - Text : B5에 "Datetime" 이라는 텍스트를, C5에 오늘 날짜를 삽입합니다.
                sheet.setValue(4, 1, "Datetime");
                sheet.getCell(4, 2).value(new Date()).formatter("yyyy-mm-dd");

                // 스타일 설정
                // B열, C열의 너비를 200으로 설정합니다.
                sheet.setColumnWidth(1, 200);
                sheet.setColumnWidth(2, 200);
                // B2:C2의 배경색과 글자색을 설정합니다.
                sheet.getRange(1, 1, 1, 2).backColor("rgb(130, 188, 0)").foreColor("rgb(255, 255, 255)");
                // B4:C4의 배경색을 설정합니다.
                sheet.getRange(3, 1, 1, 2).backColor("rgb(211, 211, 211)");
                // B2:C2의 셀을 병합합니다.
                sheet.addSpan(1, 1, 1, 2);
                // 각 범위에 테두리를 설정합니다.
                sheet.getRange(1, 1, 4, 2).setBorder(new GC.Spread.Sheets.LineBorder("Black", GC.Spread.Sheets.LineStyle.thin), {
                    all: true
                });
                sheet.getRange(1, 1, 4, 2).setBorder(new GC.Spread.Sheets.LineBorder("Black", GC.Spread.Sheets.LineStyle.dotted), {
                    inside: true
                });
                // B2:C2의 병합된 셀에 수직 가운데 정렬을 설정합니다.
                sheet.getRange(1, 1, 1, 2).hAlign(GC.Spread.Sheets.HorizontalAlign.center);

                // 데이터 바인딩
                // 새로운 시트를 추가합니다.
                spread.addSheet(1);

                // 추가한 시트를 가져옵니다.
                var sheet2 = spread.getSheet(1);

                // 데이터를 입력합니다.
                var person = { name: '홍길동', age: 25, gender: '남', address: { postcode: '10001' } };

                // 셀 바인딩 소스를 설정합니다.
                var source = new GC.Spread.Sheets.Bindings.CellBindingSource(person);

                // 지정된 시트 영역에 지정된 셀의 바인딩 경로를 설정합니다.
                sheet2.setBindingPath(2, 2, 'name');
                sheet2.setBindingPath(3, 2, 'age');
                sheet2.setBindingPath(4, 2, 'gender');
                sheet2.setBindingPath(5, 2, 'address.postcode');

                // 시트의 데이터 소스를 설정합니다.
                sheet2.setDataSource(source);
            }
        }    
    }
</script>

<style scoped>
    .spread-host {
        width: 600px;
        height: 600px;
    }
</style>