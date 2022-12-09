<template>
    <div class="wrap">
        <input type="file" @change="selectedFile" />
        <div class="container">
            <excel-preview ref="excelPreview" url="./test.xlsx"></excel-preview>
        </div>
    </div>
</template>
<script>
// import ExcelPreview from "../packages/components/ExcelPreview/ExcelPreview.vue";
export default {
    name: "app",
    // components: { ExcelPreview },
    mounted() {},
    methods: {
        selectedFile(e) {
            var fileReader = new FileReader();
            var self = this;
            fileReader.onload = (e) => {
                // 读取 Excel 文件
                self.$refs["excelPreview"].OpenExcel(fileReader.result);
            };
            fileReader.readAsArrayBuffer(e.target.files[0]);
        },
    },
};
</script>

<style scoped>
.wrap {
    width: 100vw;
    height: 100vh;
    display: flex;
    flex-direction: column;
}
.container {
    flex: 1;
    margin: 15px 0;
    width: 100%;
    overflow: hidden;
}
</style>
