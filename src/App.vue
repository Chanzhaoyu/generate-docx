<template>
  <div>
    <button @click="handleClick">生成 DOC</button>
  </div>
</template>

<script setup lang="ts">
import { createDoc } from "./utils/generateDocx";

function handleClick() {
  createDoc({
    template: "/template.docx",
    fileName: "生成文档",
    imageSize: [100, 100],
    images: {
      image: "https://img.yzcdn.cn/vant/logo.png",
      images: [
        { url: "https://img.yzcdn.cn/vant/logo.png", description: "异常图片1" },
        { url: "https://img.yzcdn.cn/vant/cat.jpeg", description: "异常图片2" },
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
}
</script>

<style scoped></style>
