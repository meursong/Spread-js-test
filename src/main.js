import { createApp } from "vue";
import App from "./App.vue";
import {
  GcSpreadSheets,
  GcWorksheet,
  GcColumn,
} from "@mescius/spread-sheets-vue";

let app = createApp(App);
app.component("gc-spread-sheets", GcSpreadSheets);
app.component("gc-worksheet", GcWorksheet);
app.component("gc-column", GcColumn);
app.mount("#app");
