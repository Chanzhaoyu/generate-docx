/**
 * 生成doc文件
 */
import { saveAs } from "file-saver";
import { generateDocxFile } from "./common";

interface FormData {
  [key: string]: string | number | boolean | null | undefined;
}

interface ImageData {
  url: string;
  description?: string;
}

interface CreateDocOptions {
  /** 表单数据 */
  form: FormData;
  /** 图片数据 */
  images?: Record<string, string | ImageData[]>;
  /** 文件名（不含扩展名） */
  fileName?: string;
  /** 图片尺寸 [宽度, 高度] */
  imageSize?: [number, number];
  /** 模板文件URL */
  template: string;
}

export async function createDoc(option: CreateDocOptions): Promise<void> {
  const { form, images, template, imageSize = [200, 200], fileName } = option;

  if (!template) {
    console.error("模板文件URL不能为空");
    return;
  }

  try {
    const templateBlob = await fetch(template).then((res) => {
      if (!res.ok)
        throw new Error(`获取模板文件失败: ${res.status} ${res.statusText}`);
      return res.blob();
    });

    const data = { ...form, ...images };

    let docBlob = await generateDocxFile(templateBlob, data, imageSize);

    if (!docBlob) return;

    saveAs(docBlob, `${fileName || new Date().getTime()}.docx`);
  } catch (error) {
    console.error("生成文档时发生错误：", error);
  }
}
