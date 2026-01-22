<template>
  <div class="file-compare">
    <div class="upload-section">
      <div class="upload-box">
        <input type="file" @change="handleFileUpload($event, 'left')" accept=".xls,.xlsx,.doc,.docx" ref="leftFile" style="display:none">
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
        <input type="file" @change="handleFileUpload($event, 'right')" accept=".xls,.xlsx,.doc,.docx" ref="rightFile" style="display:none">
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
        <div class="progress-