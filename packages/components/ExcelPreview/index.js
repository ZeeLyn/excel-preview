import ExcelPreview from "./ExcelPreview.vue";

ExcelPreview.install = (Vue) => {
    Vue.component(ExcelPreview.name, ExcelPreview);
};

export default ExcelPreview;
