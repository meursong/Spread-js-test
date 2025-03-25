<template>
  <div class="main-container">
    <!-- 스프레드시트 영역 (80%) -->
    <div class="spreadsheet-area">
      <gc-spread-sheets
          v-show="showSpreadsheet"
          class="spread-host"
          @workbookInitialized="initSpread">
      </gc-spread-sheets>
    </div>

    <!-- 버튼 및 컨트롤 영역 (20%) -->
    <div class="control-area">
      <div class="inputContainer">
        <p>엑셀 파일 가져오기 (.xlsx)</p>
        <input type="file" id="fileDemo" class="input" @change="changeFileDemo" />
        <input type="button" id="loadExcel" value="엑셀 파일 열기" class="button" @click="loadExcel" />

        <template v-if="showSpreadsheet">
          <p>엑셀 파일 내보내기 (.xlsx)</p>
          <input id="exportFileName" value="export.xlsx" class="input" @change="changeExportFileName" />
          <input type="button" id="saveExcel" value="저장" class="button" @click="saveExcel" />
        </template>
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


<style scoped>
/* 메인 컨테이너 전체 높이와 너비를 설정 */
.main-container {
  display: flex;
  width: 100%;
  height: 100vh;
}

/* 스프레드시트 영역 (80%) */
.spreadsheet-area {
  width: 80%;
  height: 80%;
}

/* 스프레드 호스트가 부모 컨테이너를 채우도록 설정 */
.spread-host {
  width: 100%;
  height: 100%;
}

/* 컨트롤 영역 (20%) */
.control-area {
  width: 20%;
  height: 100%;
  padding: 20px;
  box-sizing: border-box;
  overflow-y: auto;
  background-color: #f5f5f5;
  border-left: 1px solid #ddd;
}

/* 입력 컨테이너 */
.inputContainer {
  display: flex;
  flex-direction: column;
  gap: 10px;
}

/* 입력 필드 스타일 */
.input {
  padding: 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
}

/* 버튼 스타일 */
.button {
  padding: 10px;
  background-color: #4CAF50;
  color: white;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  font-weight: bold;
}

.button:hover {
  background-color: #45a049;
}

/* 문단 스타일 */
p {
  margin-bottom: 5px;
  font-weight: bold;
}
</style>
