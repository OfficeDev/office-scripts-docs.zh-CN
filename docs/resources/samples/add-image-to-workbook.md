---
title: 向工作簿添加图像
description: 了解如何使用 Office 脚本将映像添加到工作簿并将其复制到工作表中。
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 78c7779cf4d524ed62bf8d419135863228b23d33
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572603"
---
# <a name="add-images-to-a-workbook"></a>向工作簿添加图像

此示例演示如何在 Excel 中使用 Office 脚本处理图像。

## <a name="scenario"></a>应用场景

图像有助于品牌打造、视觉对象标识和模板。 他们帮助制作工作簿不仅仅是一张巨大的桌子。

第一个示例将图像从一个工作表复制到另一个工作表。 这可用于在每个工作表上将公司徽标放在同一位置。

第二个示例从 URL 复制图像。 这可用于将同事存储在共享文件夹中的照片复制到相关工作簿。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的add-images.xlsx](add-images.xlsx) 。 添加以下脚本并自行试用示例！

## <a name="sample-code-copy-an-image-across-worksheets"></a>示例代码：跨工作表复制图像

```TypeScript
/**
 * This script transfers an image from one worksheet to another.
 */
function main(workbook: ExcelScript.Workbook)
{
  // Get the worksheet with the image on it.
  let firstWorksheet = workbook.getWorksheet("FirstSheet");

  // Get the first image from the worksheet.
  // If a script added the image, you could add a name to make it easier to find.
  let image: ExcelScript.Image;
  firstWorksheet.getShapes().forEach((shape, index) => {
    if (shape.getType() === ExcelScript.ShapeType.image) {
      image = shape.getImage();
      return;
    }
  });

  // Copy the image to another worksheet.
  image.getShape().copyTo("SecondSheet");
}
```

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a>示例代码：将图像从 URL 添加到工作簿

```TypeScript
async function main(workbook: ExcelScript.Workbook) {
  // Fetch the image from a URL.
  const link = "https://raw.githubusercontent.com/OfficeDev/office-scripts-docs/master/docs/images/git-octocat.png";
  const response = await fetch(link);

  // Store the response as an ArrayBuffer, since it is a raw image file.
  const data = await response.arrayBuffer();

  // Convert the image data into a base64-encoded string.
  const image = convertToBase64(data);

  // Add the image to a worksheet.
  workbook.getWorksheet("WebSheet").addImage(image);
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) as string[];
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
