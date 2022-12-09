<template>
    <div class="excel-preview-wrap">
        <div class="msg" v-if="loading">loading...</div>
        <template v-if="!loading">
            <template v-if="success">
                <template v-for="sheet in sheets" :key="sheet.id">
                    <div v-if="currentSheet == sheet.id" class="sheet-container">
                        <table>
                            <template v-for="row in sheet.rows" :key="row">
                                <tr :style="row.hidden ? 'display:none' : ''">
                                    <td :height="row.height">{{ row.number }}</td>
                                    <td v-for="cell in row.cells" :key="cell" :rowspan="_GetCellRowspan(sheet, cell)" :colspan="_GetCellColspan(sheet, cell)" :valign="cell.style && cell.style.alignment && cell.style.alignment.vertical && cell.style.alignment.vertical != 'middle' ? cell.style.alignment.vertical : null" :align="cell.style && cell.style.alignment && cell.style.alignment.horizontal && cell.style.alignment.horizontal != 'left' ? cell.style.alignment.horizontal : null" :class="_GetCellClass(cell)" :style="_GetCellStyle(cell)">
                                        <a :href="_ConvertValue(cell)" v-if="cell.type == ValueType.Hyperlink" target="_blank"> {{ _ConvertValue(cell) }}</a> <template v-else> {{ _ConvertValue(cell) }}</template>
                                    </td>
                                </tr>
                            </template>
                        </table>
                    </div>
                </template>
                <div class="sheets">
                    <div v-for="sheet in sheets" :key="sheet.id" @click="currentSheet = sheet.id" :class="{ active: currentSheet == sheet.id }">{{ sheet.name }}</div>
                </div>
            </template>
            <template v-else>
                <div class="msg error">{{ error }}</div>
            </template>
        </template>
    </div>
</template>
<script>
import numfmt from "numfmt";
import ExcelToTable from "@zeelyn/excel-to-table";
export default {
    name: "ExcelPreview",
    props: {
        url: {
            type: String,
        },
        options: {
            type: Object,
            default: {},
        },
    },
    data() {
        return {
            currentSheet: 0,
            sheets: [],
            loading: true,
            success: false,
            error: "",
        };
    },
    setup() {
        return {
            ValueType: ExcelToTable.ValueType,
            ThemeColors: ExcelToTable.ThemeColors,
        };
    },
    watch: {
        url: function () {
            this._reset();
            this._loadExcelFromURL();
        }.bind(this),
    },
    mounted() {
        if (!this.url) return;
        this._loadExcelFromURL();
    },
    methods: {
        _reset() {
            this.currentSheet = 0;
            this.sheets = [];
            this.loading = true;
            this.success = false;
            this.error = "";
        },
        _loadExcelFromURL() {
            ExcelToTable.FromUrlAsync(this.url, this.options)
                .then((sheets) => {
                    this.sheets = sheets;
                    if (sheets && sheets.length > 0) {
                        this.currentSheet = sheets[0].id;
                    }
                    this.success = true;
                })
                .catch((err) => {
                    this.error = err;
                })
                .finally(() => {
                    this.loading = false;
                });
        },
        OpenExcel(data) {
            this._reset();
            ExcelToTable.FromDataAsync(data)
                .then((sheets) => {
                    this.sheets = sheets;
                    if (sheets && sheets.length > 0) {
                        this.currentSheet = sheets[0].id;
                    }
                    this.success = true;
                })
                .catch((err) => {
                    this.error = err;
                })
                .finally(() => {
                    this.loading = false;
                });
        },
        _GetCellClass(cell) {
            return { bold: cell.style && cell.style.font && cell.style.font.bold ? true : false, italic: cell.style && cell.style.font && cell.style.font.italic ? true : false, underline: cell.style && cell.style.font && cell.style.font.underline ? true : false };
        },
        _GetCellStyle(cell) {
            return {
                color: this._GetFontColor(cell),
                fontSize: this._GetFontSize(cell),
                background: this._GetBgColor(cell),
            };
        },
        _GetCellRowspan(sheet, cell) {
            if (cell.isMerged && sheet.merge && sheet.merge.length > 0) {
                var merge = sheet.merge.find((x) => x.address == cell.address);
                if (!merge) return null;
                return merge.end.row - merge.start.row + 1;
                //this.mergeCols = merge.end.col - merge.start.col + 1;
            }
            return null;
        },
        _GetCellColspan(sheet, cell) {
            if (cell.isMerged && sheet.merge && sheet.merge.length > 0) {
                var merge = sheet.merge.find((x) => x.address == cell.address);
                if (!merge) return null;
                return merge.end.col - merge.start.col + 1;
            }
            return null;
        },
        _ConvertValue(cell) {
            if (!cell.style || !cell.style.numFmt) return cell.text;
            return numfmt.format(cell.style.numFmt, cell.value);
        },
        _GetFontColor(cell) {
            if (!cell.style || !cell.style.font || !cell.style.font.color) return null;

            if (cell.style.font.color.argb) {
                // console.log("#" + this.cell.style.font.color.argb.slice(2, 8));
                return "#" + cell.style.font.color.argb.slice(2, 8);
            }
            if (cell.style.font.color.theme - 1) {
                var color = this.ThemeColors[cell.style.font.color.theme].toLowerCase();
                //默认黑色不用设置颜色
                if (color != "000000") return `#${color}`;
            }

            return null;
        },
        _GetFontSize(cell) {
            if (cell.style && cell.style.font && cell.style.font.size && cell.style.font.size != 11) return cell.style.font.size + "px";
            return null;
        },
        _GetBgColor(cell) {
            if (!cell.style || !cell.style.fill || !cell.style.fill.fgColor) return null;
            if (cell.style.fill.fgColor.argb) {
                return "#" + cell.style.fill.fgColor.argb.slice(2, 8);
            }
            if (cell.style.fill.fgColor.theme) {
                var color = this.ThemeColors[cell.style.fill.fgColor.theme].toLowerCase();
                //默认白色不用设置颜色
                if (color != "ffffff") return `#${color}`;
            }
            return null;
        },
    },
};
</script>
<style scoped>
.excel-preview-wrap {
    width: 100%;
    height: 100%;
    display: flex;
    flex-direction: column;
    justify-content: center;
}
.msg {
    text-align: center;
    width: 100%;
    font-size: 12px;
}
.error {
    color: red;
}
.sheet-container {
    flex: 1;
    overflow: auto;
}
table td {
    border-right: 1px #ccc solid;
    border-bottom: 1px #ccc solid;
    min-width: 60px;
    color: #000;
    white-space: pre-line;
}
table tr td:first-child {
    background: #eee;
    text-align: center;
    min-width: 30px !important;
}
table {
    border-left: 1px #ccc solid;
    border-top: 1px #ccc solid;
    border-collapse: collapse;
    font-size: 12px;
}
.bold {
    font-weight: bold;
}
.italic {
    font-style: italic;
}
.underline {
    text-decoration: underline;
}

.sheets {
    display: flex;
    background: #ccc;
    font-size: 13px;
    margin-top: 2px;
}
.sheets div {
    margin: 0 0 1px 1px;
    color: #666;
    padding: 5px 10px;
    cursor: pointer;
}
.sheets .active {
    background: #fff;
    color: #000;
}
</style>
