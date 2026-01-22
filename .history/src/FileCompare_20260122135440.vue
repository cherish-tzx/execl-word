<template>
  <div class="file-compare">
    <div class="upload-section">
      <div class="upload-box">
        <input type="file" @change="handleFileUpload($event, 'left')" accept=".xls,.xlsx" ref="leftFile" style="display: none" />
        <div class="upload-area" @click="$refs.leftFile.click()">
          <i class="icon-file"></i>
          <p v-if="!leftFile">点击上传文件1</p>
          <div v-else class="file-info">
            <span>{{ leftFile.name }}</span>
            <span class="file-size">{{ formatSize(leftFile.size) }}</span>
            <button @click.stop="removeFile('left')" class="remove-btn">×</button>
          </div>
        </div>
      </div>
      <div class="upload-box">
        <input type="file" @change="handleFileUpload($event, 'right')" accept=".xls,.xlsx" ref="rightFile" style="display: none" />
        <div class="upload-area" @click="$refs.rightFile.click()">
          <i class="icon-file"></i>
          <p v-if="!rightFile">点击上传文件2</p>
          <div v-else class="file-info">
            <span>{{ rightFile.name }}</span>
            <span class="file-size">{{ formatSize(rightFile.size) }}</span>
            <button @click.stop="removeFile('right')" class="remove-btn">×</button>
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
            <div class="progress-fill" :style="{ width: similarity + '%' }"></div>
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
          c