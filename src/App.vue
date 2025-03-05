<template>
  <div>
    <button @click="handleClick" :disabled="loading">
      {{ loading ? "生成中..." : "生成 DOC" }}
    </button>
    <p v-if="error" style="color: red">{{ error }}</p>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import { createDoc } from "./utils/generateDocx";

const loading = ref(false);
const error = ref("");

async function handleClick() {
  loading.value = true;
  error.value = "";

  try {
    await createDoc({
      template: "/template.docx",
      fileName: "生成文档",
      imageSize: [100, 100],
      images: {
        image: "https://img.yzcdn.cn/vant/logo.png",
        images: [
          {
            url: "https://img.yzcdn.cn/vant/logo.png",
            description: "异常图片1",
          },
          {
            url: "https://img.yzcdn.cn/vant/cat.jpeg",
            description: "异常图片2",
          },
          {
            url: "https://img.yzcdn.cn/vant/apple-1.jpg",
            description: "异常图片3",
          },
        ],
      },
      form: {
        name: "哈哈哈",
      },
    });
  } catch (e) {
    error.value = e instanceof Error ? e.message : "生成文档失败";
  } finally {
    loading.value = false;
  }
}
</script>

<style scoped></style>
