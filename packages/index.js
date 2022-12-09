import ExcelPreview from "./components/ExcelPreview";
const components = [ExcelPreview];

const install = (Vue) => {
    // 判断是否安装
    if (install.installed) return;
    // 遍历注册全局组件
    components.map((component) => Vue.component(component.name, component));
};
if (typeof window !== "undefined" && window.Vue) {
    install(window.Vue);
}
export default {
    install,
    ...components,
};
