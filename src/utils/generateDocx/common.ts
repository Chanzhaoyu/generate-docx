// @ts-ignore
import ImageModule from "docxtemplater-image-module-free";
import PizZip from "pizzip";
import DocxTemplater from "docxtemplater";

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
  const base64Regex = /^data:image\/(png|jpg|jpeg|svg|svg\+xml);base64,/;
  if (!base64Regex.test(dataURL)) {
    return false;
  }
  const stringBase64 = dataURL.replace(base64Regex, "");
  const binaryString = window.atob(stringBase64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    const ascii = binaryString.charCodeAt(i);
    bytes[i] = ascii;
  }
  return bytes.buffer;
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
      // 处理 base64 格式的图片
      if (typeof tagValue === "string" && tagValue.startsWith("data:image/")) {
        const value = base64DataURLToArrayBuffer(tagValue);
        if (value) return value;
      }

      // 处理 URL 格式的图片
      if (
        typeof tagValue === "string" &&
        (tagValue.startsWith("http://") || tagValue.startsWith("https://"))
      ) {
        try {
          return await fetchImageAsArrayBuffer(tagValue);
        } catch (error) {
          console.error("处理图片URL出错:", error);
          throw error;
        }
      }

      return false;
    },
    getSize: () => {
      return imageSize;
    },
  };
}

export async function generateDocxFile(
  template: Blob,
  fileData: any,
  imageSize: [number, number] = [200, 200]
): Promise<Blob> {
  return new Promise((resolve, reject) => {
    getFileBinaryString(template)
      .then((templateData) => {
        const zip = new PizZip(templateData);

        // 使用创建的图片选项配置
        const imageOpts = createImageOpts(imageSize);

        const doc = new DocxTemplater()
          .loadZip(zip)
          .attachModule(new ImageModule(imageOpts))
          .compile();

        doc.resolveData(fileData).then(() => {
          doc.render();
          const out = doc.getZip().generate({
            type: "blob",
            mimeType:
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          resolve(out);
        });
      })
      .catch(reject);
  });
}
