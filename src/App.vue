<template>
  <div id="app">
    <FileCompare />
  </div>
</template>

<script>
import FileCompare from "./FileCompare.vue";

export default {
  name: "App",
  components: {
    FileCompare,
  },
};
</script>

<style>
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC",
    "Hiragino Sans GB", "Microsoft YaHei", "Helvetica Neue", Helvetica, Arial,
    sans-serif;
}

#app {
  width: 100%;
  min-height: 100vh;
}
</style>
