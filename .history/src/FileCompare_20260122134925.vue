<template>
  <div class="file-compare">
    <div class="upload-section">
      <div class="upload-box">
        <input
          type="file"
          @change="handleFileUpload($event, 'left')"
          accept=".xls,.xlsx"
          ref="leftFile"
          style="display: none"
        />
        <div class="upload-area" @click="$refs.leftFile.click()">
          <i class="icon-file"></i>
          <p v-if="!leftFile">点击上传文件1</p>
          <div v-else class="file-info">
            <span>{{ leftFile.name }}</span>
            <span class="file-size">{{ formatSize(leftFile.size) }}</span>
            <button @click.stop="removeFile('left')" class="remove-btn">
              ×
            </button>
          </div>
        </div>
      </div>
      <div class="upload-box">
        <input
          type="file"
          @change="handleFileUpload($event, 'right')"
          accept=".xls,.xlsx"
          ref="rightFile"
          style="display: none"
        />
        <div class="upload-area" @click="$refs.rightFile.click()">
          <i class="icon-file"></i>
          <p v-if="!rightFile">点击上传文件2</p>
          <div v-else class="file-info">
            <span>{{ rightFile.name }}</span>
            <span class="file-size">{{ formatSize(rightFile.size) }}</span>
            <button @click.stop="removeFile('right')" class="remove-btn">
              ×
            </button>
          </div>
        </div>
      </div>
    </div>

    <div v-if="comparing" class="loading">对比中...</div>

    <div v-if="comparisonResult" class="result-section">
      <div class="similarity-bar">
        <div class="similarity-label">文件相似度</div>
        <div class="progress-container">
          <div class="progress-bar">
            <div
              class="progress-fill"
              :style="{ width: similarity + '%' }"
            ></div>
          </div>
          <div class="similarity-value">{{ similarity }}%</div>
        </div>
        <div class="progress-labels">
          <span>0</span>
          <span>50%</span>
          <span>100%</span>
        </div>
      </div>

      <div class="compare-container">
        <div class="compare-panel">
          <div class="panel-header">文件 1</div>
          <div class="content-wrapper" v-html="leftContent"></div>
        </div>
        <div class="compare-panel">
          <div class="panel-header">文件 2</div>
          <div class="content-wrapper" v-html="rightContent"></div>
        </div>
      </div>
    </div>
  </div>
</template>

<script>
import * as XLSX from "xlsx";

export default {
  name: "FileCompare",
  data() {
    return {
      leftFile: null,
      rightFile: null,
      leftData: null,
      rightData: null,
      leftContent: "",
      rightContent: "",
      comparing: false,
      comparisonResult: null,
      similarity: 0,
    };
  },
  methods: {
    async handleFileUpload(event, side) {
      const file = event.target.files[0];
      if (!file) return;

      if (side === "left") {
        this.leftFile = file;
        this.leftData = await this.parseExcel(file);
      } else {
        this.rightFile = file;
        this.rightData = await this.parseExcel(file);
      }

      if (this.leftData && this.rightData) {
        this.compareFiles();
      }
    },

    parseExcel(file) {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });
          const result = [];

          workbook.SheetNames.forEach((sheetName) => {
            const sheet = workbook.Sheets[sheetName];
            const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
            const rows = [];

            for (let R = range.s.r; R <= range.e.r; R++) {
              const row = [];
              for (let C = range.s.c; C <= range.e.c; C++) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = sheet[cellAddress];
                row.push({
                  value: cell && cell.v !== undefined ? String(cell.v) : "",
                });
              }
              rows.push(row);
            }

            result.push({ name: sheetName, rows });
          });

          resolve({ type: "excel", sheets: result });
        };
        reader.readAsArrayBuffer(file);
      });
    },

    compareFiles() {
      this.comparing = true;
      setTimeout(() => {
        this.compareExcel();
        this.comparing = false;
      }, 100);
    },

    compareExcel() {
      const leftSheets = this.leftData.sheets;
      const rightSheets = this.rightData.sheets;

      let leftHtml = '<div class="excel-container">';
      let rightHtml = '<div class="excel-container">';
      let totalCells = 0;
      let matchedCells = 0;

      for (
        let i = 0;
        i < Math.max(leftSheets.length, rightSheets.length);
        i++
      ) {
        const leftSheet = leftSheets[i];
        const rightSheet = rightSheets[i];

        if (leftSheet) {
          leftHtml += `<div class="sheet-name">[工作表：${leftSheet.name}]</div><table class="excel-table">`;
        }
        if (rightSheet) {
          rightHtml += `<div class="sheet-name">[工作表：${rightSheet.name}]</div><table class="excel-table">`;
        }

        const leftRows = leftSheet ? leftSheet.rows : [];
        const rightRows = rightSheet ? rightSheet.rows : [];
        const maxRows = Math.max(leftRows.length, rightRows.length);

        for (let r = 0; r < maxRows; r++) {
          const leftRow = leftRows[r] || [];
          const rightRow = rightRows[r] || [];
          const maxCols = Math.max(leftRow.length, rightRow.length);

          leftHtml += "<tr>";
          for (let c = 0; c < maxCols; c++) {
            const cell = leftRow[c] || { value: "" };
            leftHtml += `<td>${this.escapeHtml(cell.value)}</td>`;
          }
          leftHtml += "</tr>";

          rightHtml += "<tr>";
          for (let c = 0; c < maxCols; c++) {
            const leftCell = leftRow[c] || { value: "" };
            const rightCell = rightRow[c] || { value: "" };
            const leftVal = String(leftCell.value || "").trim();
            const rightVal = String(rightCell.value || "").trim();

            totalCells++;

            if (leftVal === rightVal) {
              matchedCells++;
              rightHtml += `<td>${this.escapeHtml(rightCell.value)}</td>`;
            } else if (!leftVal && rightVal) {
              rightHtml += `<td style="background-color: #c8e6c9 !important;">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            } else if (leftVal && !rightVal) {
              rightHtml += `<td style="background-color: #ffcdd2 !important;">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            } else {
              rightHtml += `<td style="background-color: #ffe0b2 !important;">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            }
          }
          rightHtml += "</tr>";
        }

        if (leftSheet) leftHtml += "</table>";
        if (rightSheet) rightHtml += "</table>";
      }

      leftHtml += "</div>";
      rightHtml += "</div>";

      this.leftContent = leftHtml;
      this.rightContent = rightHtml;
      this.similarity =
        totalCells > 0 ? Math.round((matchedCells / totalCells) * 100) : 0;
      this.comparisonResult = true;
    },

    escapeHtml(text) {
      const div = document.createElement("div");
      div.textContent = text;
      return div.innerHTML;
    },

    formatSize(bytes) {
      if (bytes < 1024) return bytes + " B";
      if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + " KB";
      return (bytes / (1024 * 1024)).toFixed(2) + " MB";
    },

    removeFile(side) {
      if (side === "left") {
        this.leftFile = null;
        this.leftData = null;
        this.$refs.leftFile.value = "";
      } else {
        this.rightFile = null;
        this.rightData = null;
        this.$refs.rightFile.value = "";
      }
      this.comparisonResult = null;
    },
  },
};
</script>

<style>
.file-compare {
  padding: 20px;
  background: #f5f5f5;
  min-height: 100vh;
}
.upload-section {
  display: flex;
  gap: 20px;
  margin-bottom: 30px;
}
.upload-box {
  flex: 1;
}
.upload-area {
  border: 2px dashed #d9d9d9;
  border-radius: 8px;
  padding: 40px;
  text-align: center;
  background: #fff;
  cursor: pointer;
  transition: all 0.3s;
}
.upload-area:hover {
  border-color: #40a9ff;
  background: #f0f8ff;
}
.icon-file::before {
  content: "📄";
  font-size: 48px;
  display: block;
  margin-bottom: 10px;
}
.file-info {
  display: flex;
  flex-direction: column;
  gap: 8px;
  align-items: center;
}
.file-size {
  color: #999;
  font-size: 12px;
}
.remove-btn {
  width: 24px;
  height: 24px;
  border-radius: 50%;
  border: none;
  background: #ff4d4f;
  color: #fff;
  cursor: pointer;
  font-size: 18px;
  line-height: 1;
}
.loading {
  text-align: center;
  padding: 40px;
  font-size: 16px;
  color: #666;
}
.result-section {
  background: #fff;
  border-radius: 8px;
  padding: 20px;
}
.similarity-bar {
  margin-bottom: 30px;
  padding: 20px;
  background: #fafafa;
  border-radius: 8px;
}
.similarity-label {
  font-size: 14px;
  color: #666;
  margin-bottom: 15px;
}
.progress-container {
  display: flex;
  align-items: center;
  gap: 20px;
}
.progress-bar {
  flex: 1;
  height: 20px;
  background: #e8e8e8;
  border-radius: 10px;
  overflow: hidden;
}
.progress-fill {
  height: 100%;
  background: linear-gradient(90deg, #ff4d4f 0%, #ff7875 50%, #52c41a 100%);
  transition: width 0.5s;
}
.similarity-value {
  font-size: 32px;
  font-weight: bold;
  color: #ff4d4f;
  min-width: 100px;
  text-align: center;
}
.progress-labels {
  display: flex;
  justify-content: space-between;
  margin-top: 5px;
  font-size: 12px;
  color: #999;
}
.compare-container {
  display: flex;
  gap: 20px;
}
.compare-panel {
  flex: 1;
  border: 1px solid #e8e8e8;
  border-radius: 4px;
  overflow: hidden;
}
.panel-header {
  background: #fafafa;
  padding: 12px 16px;
  font-weight: 500;
  border-bottom: 1px solid #e8e8e8;
}
.content-wrapper {
  padding: 16px;
  max-height: 600px;
  overflow-y: auto;
  background: #fff;
}
.excel-container {
  font-family: "Courier New", monospace;
  font-size: 13px;
  line-height: 1.6;
}
.sheet-name {
  font-weight: bold;
  color: #1890ff;
  margin: 15px 0 10px 0;
  padding: 8px 0;
  border-bottom: 2px solid #1890ff;
  font-size: 14px;
}
.excel-table {
  width: 100%;
  border-collapse: collapse;
  border: 2px solid #000 !important;
  margin-bottom: 20px;
}
.excel-table tr {
  border: 1px solid #000 !important;
}
.excel-table td {
  border: 1px solid #000 !important;
  padding: 8px 12px;
  min-width: 100px;
  word-break: break-word;
  background-color: #fff;
  vertical-align: top;
}
</style>
