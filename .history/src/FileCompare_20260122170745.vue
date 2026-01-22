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
        const rowAlignment = this.alignRowsWithLCS(leftRows, rightRows);

        // 全局分析列的对应关系（基于第一行或表头行）
        let globalColMapping = null;
        const firstEqualRow = rowAlignment.find(
          (item) => item.type === "equal"
        );
        if (firstEqualRow) {
          globalColMapping = this.getColumnMapping(
            firstEqualRow.leftRow,
            firstEqualRow.rightRow
          );
        }

        rowAlignment.forEach((item) => {
          const { type, leftRow, rightRow } = item;

          if (leftRow) {
            leftHtml += "<tr>";
            leftRow.forEach((cell) => {
              leftHtml += `<td>${this.escapeHtml(cell.value)}</td>`;
            });
            leftHtml += "</tr>";
          } else {
            leftHtml += "<tr>";
            if (rightRow) {
              rightRow.forEach(() => {
                leftHtml += '<td style="background-color: #f5f5f5;"></td>';
              });
            }
            leftHtml += "</tr>";
          }

          if (type === "equal") {
            // 使用列对齐算法找到列的对应关系
            const colMapping = this.getColumnMapping(leftRow, rightRow);

            // 右侧显示原始内容，但根据映射关系标记颜色
            rightHtml += "<tr>";
            rightRow.forEach((rightCell, rightIdx) => {
              totalCells++;
              const leftIdx = colMapping.rightToLeft[rightIdx];

              if (leftIdx !== undefined) {
                // 右侧列在左侧有对应
                const leftCell = leftRow[leftIdx];
                const leftVal = String(leftCell.value || "").trim();
                const rightVal = String(rightCell.value || "").trim();

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
              } else {
                // 右侧新增的列
                rightHtml += `<td style="background-color: #c8e6c9 !important;">${this.escapeHtml(
                  rightCell.value
                )}</td>`;
              }
            });
            rightHtml += "</tr>";
          } else if (type === "insert") {
            rightHtml += "<tr>";
            rightRow.forEach((cell) => {
              totalCells++;
              rightHtml += `<td style="background-color: #c8e6c9 !important;">${this.escapeHtml(
                cell.value
              )}</td>`;
            });
            rightHtml += "</tr>";
          } else if (type === "delete") {
            rightHtml += "<tr>";
            leftRow.forEach(() => {
              totalCells++;
              rightHtml += `<td style="background-color: #ffcdd2 !important;"></td>`;
            });
            rightHtml += "</tr>";
          }
        });

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

    alignRowsWithLCS(leftRows, rightRows) {
      const m = leftRows.length;
      const n = rightRows.length;
      const dp = Array(m + 1)
        .fill(null)
        .map(() => Array(n + 1).fill(0));

      for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
          if (this.rowsAreSimilar(leftRows[i - 1], rightRows[j - 1])) {
            dp[i][j] = dp[i - 1][j - 1] + 1;
          } else {
            dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
          }
        }
      }

      const alignment = [];
      let i = m,
        j = n;
      while (i > 0 || j > 0) {
        if (
          i > 0 &&
          j > 0 &&
          this.rowsAreSimilar(leftRows[i - 1], rightRows[j - 1])
        ) {
          alignment.unshift({
            type: "equal",
            leftRow: leftRows[i - 1],
            rightRow: rightRows[j - 1],
            leftIndex: i - 1,
            rightIndex: j - 1,
          });
          i--;
          j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
          alignment.unshift({
            type: "insert",
            leftRow: null,
            rightRow: rightRows[j - 1],
            leftIndex: -1,
            rightIndex: j - 1,
          });
          j--;
        } else if (i > 0) {
          alignment.unshift({
            type: "delete",
            leftRow: leftRows[i - 1],
            rightRow: null,
            leftIndex: i - 1,
            rightIndex: -1,
          });
          i--;
        }
      }
      return alignment;
    },

    getColumnMapping(leftRow, rightRow) {
      const m = leftRow.length;
      const n = rightRow.length;
      const leftToRight = {};
      const rightToLeft = {};

      // 第一步：使用LCS找到非空值的强匹配
      const dp = Array(m + 1)
        .fill(null)
        .map(() => Array(n + 1).fill(0));

      for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
          const leftVal = String(leftRow[i - 1].value || "").trim();
          const rightVal = String(rightRow[j - 1].value || "").trim();

          // 只有当两个值都非空且相等时才认为是强匹配
          if (leftVal && rightVal && leftVal === rightVal) {
            // 位置越接近，权重越高
            const positionBonus = 1 - (Math.abs(i - j) / Math.max(m, n)) * 0.3;
            dp[i][j] = dp[i - 1][j - 1] + 1 + positionBonus;
          } else {
            dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
          }
        }
      }

      // 回溯找到强匹配的映射关系
      let i = m,
        j = n;

      while (i > 0 && j > 0) {
        const leftVal = String(leftRow[i - 1].value || "").trim();
        const rightVal = String(rightRow[j - 1].value || "").trim();

        if (
          leftVal &&
          rightVal &&
          leftVal === rightVal &&
          dp[i][j] > Math.max(dp[i - 1][j], dp[i][j - 1])
        ) {
          leftToRight[i - 1] = j - 1;
          rightToLeft[j - 1] = i - 1;
          i--;
          j--;
        } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
          j--;
        } else {
          i--;
        }
      }

      // 第二步：对未匹配的列，按位置距离最近原则进行兜底匹配
      // 但只匹配都为空的列，避免把空列和有内容的列错误匹配
      const unmatchedLeft = [];
      const unmatchedRight = [];

      for (let i = 0; i < m; i++) {
        if (leftToRight[i] === undefined) {
          unmatchedLeft.push(i);
        }
      }

      for (let j = 0; j < n; j++) {
        if (rightToLeft[j] === undefined) {
          unmatchedRight.push(j);
        }
      }

      // 第一轮：只匹配都为空的列
      const usedRight = new Set();
      for (const leftIdx of unmatchedLeft) {
        const leftVal = String(leftRow[leftIdx].value || "").trim();
        let bestRightIdx = -1;
        let minDistance = Infinity;

        for (const rightIdx of unmatchedRight) {
          if (!usedRight.has(rightIdx)) {
            const rightVal = String(rightRow[rightIdx].value || "").trim();
            // 只有当两列都为空时才考虑匹配
            if (!leftVal && !rightVal) {
              const distance = Math.abs(leftIdx - rightIdx);
              if (distance < minDistance) {
                minDistance = distance;
                bestRightIdx = rightIdx;
              }
            }
          }
        }

        if (bestRightIdx !== -1) {
          leftToRight[leftIdx] = bestRightIdx;
          rightToLeft[bestRightIdx] = leftIdx;
          usedRight.add(bestRightIdx);
        }
      }

      // 第二轮：对于位置非常接近的列（距离<=2），即使内容不同也匹配为修改
      const stillUnmatchedLeft = unmatchedLeft.filter(
        (idx) => leftToRight[idx] === undefined
      );
      const stillUnmatchedRight = unmatchedRight.filter(
        (idx) => !usedRight.has(idx)
      );

      for (const leftIdx of stillUnmatchedLeft) {
        let bestRightIdx = -1;
        let minDistance = Infinity;

        for (const rightIdx of stillUnmatchedRight) {
          if (!usedRight.has(rightIdx)) {
            const distance = Math.abs(leftIdx - rightIdx);
            // 只匹配位置非常接近的列（距离<=2）
            if (distance <= 2 && distance < minDistance) {
              minDistance = distance;
              bestRightIdx = rightIdx;
            }
          }
        }

        if (bestRightIdx !== -1) {
          leftToRight[leftIdx] = bestRightIdx;
          rightToLeft[bestRightIdx] = leftIdx;
          usedRight.add(bestRightIdx);
        }
      }

      return { leftToRight, rightToLeft };
    },

    alignColumnsWithLCS(leftRow, rightRow) {
      const m = leftRow.length;
      const n = rightRow.length;
      const dp = Array(m + 1)
        .fill(null)
        .map(() => Array(n + 1).fill(0));

      for (let i = 1; i <= m; i++) {
        for (let j = 1; j <= n; j++) {
          const leftVal = String(leftRow[i - 1].value || "").trim();
          const rightVal = String(rightRow[j - 1].value || "").trim();
          if (leftVal && rightVal && leftVal === rightVal) {
            dp[i][j] = dp[i - 1][j - 1] + 1;
          } else {
            dp[i][j] = Math.max(dp[i - 1][j], dp[i][j - 1]);
          }
        }
      }

      const alignment = [];
      let i = m,
        j = n;
      while (i > 0 || j > 0) {
        if (i > 0 && j > 0) {
          const leftVal = String(leftRow[i - 1].value || "").trim();
          const rightVal = String(rightRow[j - 1].value || "").trim();
          if (leftVal && rightVal && leftVal === rightVal) {
            alignment.unshift({
              type: "equal",
              leftCell: leftRow[i - 1],
              rightCell: rightRow[j - 1],
            });
            i--;
            j--;
            continue;
          }
        }
        if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1][j])) {
          alignment.unshift({
            type: "insert",
            leftCell: null,
            rightCell: rightRow[j - 1],
          });
          j--;
        } else if (i > 0) {
          alignment.unshift({
            type: "delete",
            leftCell: leftRow[i - 1],
            rightCell: null,
          });
          i--;
        }
      }
      return alignment;
    },

    rowsAreSimilar(row1, row2) {
      if (!row1 || !row2) return false;
      const getFingerprint = (row) => {
        return row
          .map((cell) => String(cell.value || "").trim())
          .filter((v) => v !== "")
          .slice(0, 3)
          .join("|");
      };
      const fp1 = getFingerprint(row1);
      const fp2 = getFingerprint(row2);
      if (fp1 && fp2 && fp1 === fp2) return true;

      const minLen = Math.min(row1.length, row2.length);
      if (minLen === 0) return false;
      let matchCount = 0;
      let totalNonEmpty = 0;
      for (let i = 0; i < minLen; i++) {
        const val1 = String(row1[i].value || "").trim();
        const val2 = String(row2[i].value || "").trim();
        if (val1 || val2) {
          totalNonEmpty++;
          if (val1 === val2) matchCount++;
        }
      }
      return totalNonEmpty > 0 && matchCount / totalNonEmpty >= 0.7;
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
