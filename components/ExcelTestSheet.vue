<template>
  <div class="spreadsheet-container">
    <gc-spread-sheets
        v-show="showSpreadsheet"
        class="spread-host"
        @workbookInitialized="initSpread">
    </gc-spread-sheets>
    <div class="options-container">
      <div class="option-row">
        <div class="inputContainer">
          <p>Open Excel File (.xlsx)</p>
          <input type="file" id="fileDemo" class="input" @change="changeFileDemo" />
          <input type="button" id="loadExcel" value="Open Excel" class="button" @click="loadExcel" />
          <p>Add Data</p>
          <input type="button" id="saveExcel" value="Add Revenue" class="button" @click="modifyExcel" />

          <p>Save Excel File (.xlsx)</p>
          <input id="exportFileName" value="export.xlsx" class="input" @change="changeExportFileName" />
          <input type="button" id="saveExcel" value="Save Excel" class="button" @click="saveExcel" />
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

  let spreadInstance = spread.value;
  let excelFile = importExcelFile.value;
  let options = {
    fileType: GC.Spread.Sheets.FileType.excel,
  };
  // Import an existing Excel file to Vue spreadsheet
  spreadInstance.import(
      excelFile,
      () => {
        console.log("Import successful");
      },
      (e) => {
        console.error("Error during import:", e);
      },
      options
  );
};

const modifyExcel = () => {
  let spreadInstance = spread.value;
  let sheet = spreadInstance.getActiveSheet();
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
  var fileName = exportFileName.value || "Excel_Export.xlsx";
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