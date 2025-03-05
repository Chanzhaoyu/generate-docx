// @ts-ignore
import ImageModule from "docxtemplater-image-module-free";

import PizZip from "pizzip";
import DocxTemplater from "docxtemplater";

const BASE64_REGEX = /^data:image\/(png|jpg|jpeg|svg|svg\+xml);base64,/;

function getFileBinaryString(
  templateFile: string | File | Blob
): Promise<string> {
  return new Promise((resolve, reject) => {
    // 如果传的是个字符串表示就是二进制字符串
    if (typeof templateFile === "string") {
      resolve(templateFile);
      return;
    }
    // 处理File或Blob对象
    const reader = new FileReader();
    reader.onload = (e: ProgressEvent<FileReader>) => {
      if (e.target?.result) {
        const result = e.target.result as string;
        resolve(result);
      } else {
        reject(new Error("读取文件失败"));
      }
    };
    reader.onerror = (error) => {
      reject(error);
    };
    reader.readAsBinaryString(templateFile);
  });
}

function base64DataURLToArrayBuffer(dataURL: string) {
  if (!BASE64_REGEX.test(dataURL)) {
    return false;
  }
  const stringBase64 = dataURL.replace(BASE64_REGEX, "");
  const binaryString = window.atob(stringBase64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < bytes.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}

async function processImage(tagValue: string): Promise<ArrayBuffer | false> {
  // 处理 base64 格式的图片
  if (typeof tagValue === "string" && tagValue.startsWith("data:image/")) {
    return base64DataURLToArrayBuffer(tagValue);
  }

  // 处理 URL 格式的图片
  if (
    typeof tagValue === "string" &&
    (tagValue.startsWith("http://") || tagValue.startsWith("https://"))
  ) {
    return fetchImageAsArrayBuffer(tagValue);
  }

  return false;
}

async function fetchImageAsArrayBuffer(url: string): Promise<ArrayBuffer> {
  try {
    const response = await fetch(url, { mode: "cors" });
    if (!response.ok) {
      throw new Error(
        `获取图片失败: ${response.status} ${response.statusText}`
      );
    }
    const blob = await response.blob();
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as ArrayBuffer);
      reader.onerror = reject;
      reader.readAsArrayBuffer(blob);
    });
  } catch (error) {
    console.error("获取图片出错:", error);
    throw error;
  }
}

function createImageOpts(imageSize: [number, number] = [200, 200]) {
  return {
    centered: false,
    getImage: async (tagValue: string) => {
      try {
        return await processImage(tagValue);
      } catch (error) {
        console.error("处理图片失败:", error);
        throw error;
      }
    },
    getSize: () => {
      return imageSize;
    },
  };
}

export async function generateDocxFile(
  template: Blob,
  fileData: Record<string, unknown>,
  imageSize: [number, number] = [200, 200]
): Promise<Blob> {
  try {
    const templateData = await getFileBinaryString(template);
    const zip = new PizZip(templateData);
    const imageOpts = createImageOpts(imageSize);

    const doc = new DocxTemplater()
      .loadZip(zip)
      .attachModule(new ImageModule(imageOpts))
      .compile();

    await doc.resolveData(fileData);
    doc.render();

    return doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  } catch (error) {
    console.error("生成Word文档失败:", error);
    throw error;
  }
}
