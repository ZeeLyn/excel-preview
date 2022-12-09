import { createApp } from "vue";
import "./style.css";
import App from "./App.vue";
import ExcelPreview from "@zeelyn/excel-preview-vue3";
const app = createApp(App);
import "@zeelyn/excel-preview-vue3/dist/style.css";
app.use(ExcelPreview);
app.mount("#app");
