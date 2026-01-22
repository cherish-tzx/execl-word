<template>
  <div class="file-compare">
    <div class="upload-section">
      <div class="upload-box">
        <input
          type="file"
          @change="handleFileUpload($event, 'left')"
          accept=".xls,.xlsx,.doc,.docx"
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
          accept=".xls,.xlsx,.doc,.docx"
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
import mammoth from "mammoth";

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
        this.leftData = await this.parseFile(file);
      } else {
        this.rightFile = file;
        this.rightData = await this.parseFile(file);
      }

      if (this.leftData && this.rightData) {
        this.compareFiles();
      }
    },

    async parseFile(file) {
      const ext = file.name.split(".").pop().toLowerCase();

      if (ext === "xls" || ext === "xlsx") {
        return await this.parseExcel(file);
      } else if (ext === "doc" || ext === "docx") {
        return await this.parseWord(file);
      }
    },

    parseExcel(file) {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array", cellStyles: true });
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
                  value: cell
                    ? cell.v !== undefined
                      ? String(cell.v)
                      : ""
                    : "",
                  style: cell ? cell.s : null,
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

    parseWord(file) {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = async (e) => {
          const arrayBuffer = e.target.result;
          const result = await mammoth.extractRawText({ arrayBuffer });
          const lines = result.value
            .split("\n")
            .map((line) => line.replace(/\r/g, ""));
          resolve({ type: "word", lines });
        };
        reader.readAsArrayBuffer(file);
      });
    },

    compareFiles() {
      this.comparing = true;

      setTimeout(() => {
        if (this.leftData.type === "excel") {
          this.compareExcel();
        } else {
          this.compareWord();
        }
        this.comparing = false;
      }, 100);
    },

    compareExcel() {
      const leftSheets = this.leftData.sheets;
      const rightSheets = this.rightData.sheets;

      let leftHtml = '<div class="excel-container">';
      let rightHtml = '<div class="excel-container">';

      const maxSheets = Math.max(leftSheets.length, rightSheets.length);
      let totalCells = 0;
      let matchedCells = 0;

      for (let i = 0; i < maxSheets; i++) {
        const leftSheet = leftSheets[i];
        const rightSheet = rightSheets[i];

        if (leftSheet) {
          leftHtml += `<div class="sheet-name">[工作表：${leftSheet.name}]</div>`;
          leftHtml += '<table class="excel-table">';
        }
        if (rightSheet) {
          rightHtml += `<div class="sheet-name">[工作表：${rightSheet.name}]</div>`;
          rightHtml += '<table class="excel-table">';
        }

        const maxRows = Math.max(
          leftSheet ? leftSheet.rows.length : 0,
          rightSheet ? rightSheet.rows.length : 0
        );

        for (let r = 0; r < maxRows; r++) {
          const leftRow =
            leftSheet && leftSheet.rows[r] ? leftSheet.rows[r] : [];
          const rightRow =
            rightSheet && rightSheet.rows[r] ? rightSheet.rows[r] : [];
          const maxCols = Math.max(leftRow.length, rightRow.length);

          let leftRowHtml = "<tr>";
          let rightRowHtml = "<tr>";

          for (let c = 0; c < maxCols; c++) {
            const leftCell = leftRow[c] || { value: "", style: null };
            const rightCell = rightRow[c] || { value: "", style: null };

            totalCells++;
            const diff = this.compareCellValues(
              leftCell.value,
              rightCell.value
            );

            // 左侧：始终显示原始内容，不带颜色标记
            leftRowHtml += `<td class="excel-cell">${this.escapeHtml(
              leftCell.value
            )}</td>`;

            // 右侧：显示原始内容，但根据差异类型添加颜色标记
            if (diff.type === "equal") {
              matchedCells++;
              rightRowHtml += `<td class="excel-cell">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            } else if (diff.type === "delete") {
              // 右侧文件中删除的内容（左侧有，右侧无）
              rightRowHtml += `<td class="excel-cell deleted-cell"></td>`;
            } else if (diff.type === "add") {
              // 右侧文件中新增的内容（左侧无，右侧有）
              rightRowHtml += `<td class="excel-cell added-cell">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            } else {
              // 修改的内容
              rightRowHtml += `<td class="excel-cell modified-cell">${this.escapeHtml(
                rightCell.value
              )}</td>`;
            }
          }

          leftRowHtml += "</tr>";
          rightRowHtml += "</tr>";
          leftHtml += leftRowHtml;
          rightHtml += rightRowHtml;
        }

        if (leftSheet) {
          leftHtml += "</table>";
        }
        if (rightSheet) {
          rightHtml += "</table>";
        }
      }

      leftHtml += "</div>";
      rightHtml += "</div>";

      this.leftContent = leftHtml;
      this.rightContent = rightHtml;
      this.similarity =
        totalCells > 0 ? Math.round((matchedCells / totalCells) * 100) : 0;
      this.comparisonResult = true;
    },

    compareWord() {
      const leftLines = this.leftData.lines;
      const rightLines = this.rightData.lines;
      const maxLines = Math.max(leftLines.length, rightLines.length);

      let leftHtml = '<div class="word-container">';
      let rightHtml = '<div class="word-container">';
      let totalChars = 0;
      let matchedChars = 0;

      for (let i = 0; i < maxLines; i++) {
        const leftLine = leftLines[i] || "";
        const rightLine = rightLines[i] || "";

        totalChars += Math.max(leftLine.length, rightLine.length);
        const diff = this.compareCellValues(leftLine, rightLine);

        // 左侧：始终显示原始内容，不带颜色标记
        leftHtml += `<div class="word-line">${
          this.escapeHtml(leftLine) || "&nbsp;"
        }</div>`;

        // 右侧：显示原始内容，但根据差异类型添加颜色标记
        if (diff.type === "equal") {
          matchedChars += leftLine.length;
          rightHtml += `<div class="word-line">${
            this.escapeHtml(rightLine) || "&nbsp;"
          }</div>`;
        } else if (diff.type === "delete") {
          rightHtml += `<div class="word-line deleted-line">&nbsp;</div>`;
        } else if (diff.type === "add") {
          rightHtml += `<div class="word-line added-line">${this.escapeHtml(
            rightLine
          )}</div>`;
        } else {
          matchedChars += diff.matched;
          rightHtml += `<div class="word-line modified-line">${this.escapeHtml(
            rightLine
          )}</div>`;
        }
      }

      leftHtml += "</div>";
      rightHtml += "</div>";

      this.leftContent = leftHtml;
      this.rightContent = rightHtml;
      this.similarity =
        totalChars > 0 ? Math.round((matchedChars / totalChars) * 100) : 0;
      this.comparisonResult = true;
    },

    compareCellValues(left, right) {
      if (left === right) {
        return { type: "equal" };
      }
      if (!left && right) {
        return { type: "add" };
      }
      if (left && !right) {
        return { type: "delete" };
      }

      const leftChars = Array.from(left);
      const rightChars = Array.from(right);
      const dp = Array(leftChars.length + 1)
        .fill(null)
        .map(() => Array(rightChars.length + 1).fill(0));

      for (let i = 0; i <= leftChars.length; i++) {
        for (let j = 0; j <= rightChars.length; j++) {
          if (i === 0 || j === 0) {
            dp[i][j] = 0;
          } else if (leftChars[i - 1] === rightChars[j - 1]) {
            dp[i][j] = dp[i - 1][j - 1] + 1;
          } else {
            dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
          }
        }
      }

      const parts = [];
      let i = leftChars.length;
      let j = rightChars.length;
      let matched = 0;

      while (i > 0 && j > 0) {
        if (leftChars[i - 1] === rightChars[j - 1]) {
          parts.unshift({ type: "equal", char: rightChars[j - 1] });
          matched++;
          i--;
          j--;
        } else if (dp[i - 1][j] > dp[i][j - 1]) {
          i--;
        } else {
          parts.unshift({ type: "add", char: rightChars[j - 1] });
          j--;
        }
      }

      while (j > 0) {
        parts.unshift({ type: "add", char: rightChars[j - 1] });
        j--;
      }

      return { type: "modified", parts, matched };
    },

    renderDiff(parts) {
      return parts
        .map((part) => {
          const char = this.escapeHtml(part.char);
          if (part.type === "add") {
            return `<span class="char-added">${char}</span>`;
          }
          return char;
        })
        .join("");
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

<style scoped>
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
}

.excel-container,
.word-container {
  font-family: "Courier New", monospace;
  font-size: 13px;
  line-height: 1.6;
}

.excel-table {
  width: 100%;
  border-collapse: collapse;
  border: 1px solid #d9d9d9;
}

.excel-table td {
  border: 1px solid #d9d9d9;
  padding: 6px 10px;
  min-width: 80px;
  word-break: break-all;
  background: #fff;
}

.sheet-name {
  font-weight: bold;
  color: #1890ff;
  margin: 10px 0;
  padding: 5px 0;
  border-bottom: 1px solid #e8e8e8;
}

.excel-row {
  display: flex;
  border-bottom: 1px solid #f0f0f0;
}

.excel-cell {
  flex: 1;
  padding: 6px 10px;
  border-right: 1px solid #f0f0f0;
  min-width: 80px;
  word-break: break-all;
}

.word-line {
  padding: 4px 8px;
  min-height: 24px;
  border-bottom: 1px solid #f9f9f9;
}

/* 右侧表格的颜色标记 */
.deleted-cell {
  background: #ffcdd2 !important;
}

.added-cell {
  background: #c8e6c9 !important;
}

.modified-cell {
  background: #ffe0b2 !important;
}

.deleted-line {
  background: #ffcdd2;
}

.added-line {
  background: #c8e6c9;
}

.modified-line {
  background: #ffe0b2;
}

.deleted {
  background: #ffebee;
}

.added {
  background: #e8f5e9;
}

.modified {
  background: #fff3e0;
}

.empty {
  background: #fafafa;
}

.char-added {
  background: #a5d6a7;
  padding: 2px 0;
}
</style>
