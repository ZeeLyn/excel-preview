# Excel Preview

```javascript
import { createApp } from "vue";
import App from "./App.vue";
import ExcelPreview from "@zeelyn/excel-preview-vue3";
const app = createApp(App);
import "@zeelyn/excel-preview-vue3/dist/style.css";
app.use(ExcelPreview);
app.mount("#app");

```


```html
<template>
    <div class="wrap">
        <input type="file" @change="selectedFile" />
        <div class="container">
            <excel-preview ref="excelPreview" url="./test.xlsx"></excel-preview>
        </div>
    </div>
</template>
```


```javascript
<script>
export default {
    name: "page",
    mounted() {},
    methods: {
        selectedFile(e) {
            var fileReader = new FileReader();
            var self = this;
            fileReader.onload = (e) => {
                self.$refs["excelPreview"].OpenExcel(fileReader.result);
            };
            fileReader.readAsArrayBuffer(e.target.files[0]);
        },
    },
};
</script>
```