import { defineConfig } from "vite";
import vue from "@vitejs/plugin-vue";

// https://vitejs.dev/config/
export default defineConfig({
    plugins: [vue()],
    build: {
        copyPublicDir: false,
        rollupOptions: {
            external: ["vue", "numfmt", "@zeelyn/excel-to-table"],
            output: {
                globals: {
                    vue: "Vue",
                },
            },
        },
        lib: {
            name: "excel-preview",
            entry: "./packages/index.js",
            fileName: (format) => `excel-preview-vue3.${format}.js`,
        },
    },
});
