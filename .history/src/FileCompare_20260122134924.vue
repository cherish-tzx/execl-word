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
        <input type="file" @change="handleFil