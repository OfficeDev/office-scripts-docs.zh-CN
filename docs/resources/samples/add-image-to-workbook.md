---
title: 向工作簿添加图像
description: 了解如何使用 Office 脚本将图像添加到工作簿中，以及如何跨工作表复制该图像。
ms.date: 07/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 0c4b3446df8de280b6cb557e291504ceed5ee7f7
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59326854"
---
# <a name="add-images-to-a-workbook"></a>向工作簿添加图像

此示例演示如何使用 Excel 中的 Office 脚本处理Excel。

## <a name="scenario"></a>应用场景

图像有助于打造品牌、视觉标识和模板。 它们帮助使工作簿不仅仅是一个大型表。

第一个示例将图像从一个工作表复制到另一个工作表。 这可用于将公司的徽标置于每个工作表上的相同位置。

第二个示例从 URL 复制图像。 这可用于将同事存储在共享文件夹中的照片复制到相关工作簿。

## <a name="sample-excel-file"></a>示例Excel文件

下载 <a href="add-images.xlsx">add-images.xlsx</a> 工作簿的工作簿。 添加以下脚本，然后自己尝试示例！

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
