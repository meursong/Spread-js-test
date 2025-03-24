<template>
  <div class="spreadsheet-container flex flex-col">
    <div class="relative">
      <gc-spread-sheets
          v-show="showSpreadsheet"
          class="spread-host"
          @workbookInitialized="initSpread">
      </gc-spread-sheets>
    </div>
    <div class="options-container">
      <div class="option-row">
        <div class="inputContainer">
          <p>엑셀 파일 가져오기 (.xlsx)</p>
          <input type="file" id="fileDemo" class="input" @change="changeFileDemo" />
          <input type="button" id="loadExcel" value="엑셀 파일 열기" class="button" @click="loadExcel" />
          <p>데이터 부여</p>
          <input type="button" id="saveExcel" value="데이터 넣기 테스트" class="button" @click="modifyExcel" />

          <p>엑셀 파일 내보내기 (.xlsx)</p>
          <input id="exportFileName" value="export.xlsx" class="input" @change="changeExportFileName" />
          <input type="button" id="saveExcel" value="저장" class="button" @click="saveExcel" />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue';
import "@mescius/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css";
import * as GC from "@mescius/spread-sheets";
import { GcSpreadSheets } from '@mescius/spread-sheets-vue'
import "@mescius/spread-sheets-io";
import { saveAs } from "file-saver";

// 반응형 상태 정의
const spread = ref(null);
const revenueCount = ref(8);
const newRowIndex = ref(11);
const importExcelFile = ref(null);
const exportFileName = ref('export.xlsx');
const showSpreadsheet = ref(false); // 스프레드시트 표시 여부 제어

// Initialize the Vue Spreadsheet Instance
const initSpread = (spreadInstance) => {
  spread.value = spreadInstance;
  // 스프레드 인스턴스가 설정되었는지 확인
  if (spread.value) {
    console.log("Spread instance initialized successfully");
  }
};

const changeFileDemo = (e) => {
  importExcelFile.value = e.target.files[0];
};

const changeExportFileName = (e) => {
  exportFileName.value = e.target.value;
};

const loadExcel = () => {
  if (!importExcelFile.value) {
    console.error("선택된 파일이 없습니다.");
    alert("먼저 Excel 파일을 선택해주세요.");
    return;
  }

  // 스프레드시트를 표시
  showSpreadsheet.value = true;

  // 약간의 지연을 두어 DOM이 업데이트된 후 스프레드 인스턴스가 생성되도록 함
  nextTick(() => {
    if (!spread.value) {
      console.error("스프레드 인스턴스가 초기화되지 않았습니다.");
      return;
    }

    const options = {
      fileType: GC.Spread.Sheets.FileType.excel,
    };

    // Import an existing Excel file to Vue spreadsheet
    spread.value.import(
        importExcelFile.value,
        () => {
          console.log("Import successful");
        },
        (e) => {
          console.error("Error during import:", e);
        },
        options
    );
  }, 300);
};

const modifyExcel = () => {
  if (!spread.value) {
    console.error("스프레드 인스턴스가 초기화되지 않았습니다.");
    return;
  }

  const sheet = spread.value.getActiveSheet();
  if (!sheet) {
    console.error("활성화된 시트가 없습니다.");
    return;
  }

  // Add a new row for the next revenue item
  sheet.addRows(newRowIndex.value, 1);
  // Copy styles from an existing row
  sheet.copyTo(10, 1, newRowIndex.value, 1, 1, 29, GC.Spread.Sheets.CopyToOptions.style);
  // Set the new row's first column with the revenue label
  var cellText = ("Revenue " + revenueCount.value++);
  sheet.setValue(newRowIndex.value, 1, cellText);
  // Fill the row with random revenue data
  for (var c = 3; c < 15; c++) {
    sheet.setValue(newRowIndex.value, c, Math.floor(Math.random() * 200) + 10);
  }
  // Add a sparkline chart for the revenue data
  var data = new GC.Spread.Sheets.Range(newRowIndex.value, 3, 1, 12);
  var setting = new GC.Spread.Sheets.Sparklines.SparklineSetting();
  setting.options.seriesColor = "Text 2";
  setting.options.lineWeight = 1;
  setting.options.showLow = true;
  setting.options.showHigh = true;
  setting.options.lowMarkerColor = "Text 2";
  setting.options.highMarkerColor = "Text 1";
  sheet.setSparkline(newRowIndex.value, 2, data, GC.Spread.Sheets.Sparklines.DataOrientation.horizontal, GC.Spread.Sheets.Sparklines.SparklineType.line, setting);
  // Add a formula to calculate the sum for the new row's year data
  sheet.setFormula(newRowIndex.value, 15, "=SUM([@[Jan]:[Dec]])");
  // Move to the next row for future modifications
  newRowIndex.value++;
  // Set a fixed value in the next column
  sheet.setValue(newRowIndex.value, 16, 0.15);
  // Copy formula settings to other parts of the row
  sheet.copyTo(10, 17, newRowIndex.value, 17, 1, 13, GC.Spread.Sheets.CopyToOptions.formula);
};

const saveExcel = () => {
  let spreadInstance = spread.value;
  var fileName = exportFileName.value || "엑셀_내보내기.xlsx";
  // Save Vue spreadsheet to local Excel XLSX file
  spreadInstance.export(
      function (blob) {
        saveAs(blob, fileName);
      },
      function (e) {
        console.log(e);
      },
      {
        fileType: GC.Spread.Sheets.FileType.excel
      }
  );
};
</script>