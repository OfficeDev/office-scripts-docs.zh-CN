---
title: 向工作簿添加图像
description: 了解如何使用 Office 脚本将图像添加到工作簿中，以及如何跨工作表复制该图像。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 64c356b2d76a561276b2955263555b16de27b3ba
ms.sourcegitcommit: a2b85168d2b5e2c4e6951f808368f7d726400df0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592752"
---
# <a name="add-images-to-a-workbook"></a><span data-ttu-id="bb972-103">向工作簿添加图像</span><span class="sxs-lookup"><span data-stu-id="bb972-103">Add images to a workbook</span></span>

<span data-ttu-id="bb972-104">此示例演示如何使用 Excel 中的 Office 脚本处理Excel。</span><span class="sxs-lookup"><span data-stu-id="bb972-104">This sample shows how to work with images using an Office Script in Excel.</span></span>

## <a name="scenario"></a><span data-ttu-id="bb972-105">应用场景</span><span class="sxs-lookup"><span data-stu-id="bb972-105">Scenario</span></span>

<span data-ttu-id="bb972-106">图像有助于打造品牌、视觉标识和模板。</span><span class="sxs-lookup"><span data-stu-id="bb972-106">Images help with branding, visual identity, and templates.</span></span> <span data-ttu-id="bb972-107">它们帮助使工作簿不仅仅是一个大型表。</span><span class="sxs-lookup"><span data-stu-id="bb972-107">They help make a workbook more than just a giant table.</span></span>

<span data-ttu-id="bb972-108">第一个示例将图像从一个工作表复制到另一个工作表。</span><span class="sxs-lookup"><span data-stu-id="bb972-108">The first sample copies an image from one worksheet to another.</span></span> <span data-ttu-id="bb972-109">这可用于将公司的徽标置于每个工作表上的相同位置。</span><span class="sxs-lookup"><span data-stu-id="bb972-109">This could be used to put your company's logo in the same position on every sheet.</span></span>

<span data-ttu-id="bb972-110">第二个示例从 URL 复制图像。</span><span class="sxs-lookup"><span data-stu-id="bb972-110">The second sample copies an image from a URL.</span></span> <span data-ttu-id="bb972-111">这可用于将同事存储在共享文件夹中的照片复制到相关工作簿。</span><span class="sxs-lookup"><span data-stu-id="bb972-111">This could be used to copy photos that a colleague stored in a shared folder to a related workbook.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="bb972-112">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="bb972-112">Sample Excel file</span></span>

<span data-ttu-id="bb972-113">下载这些 <a href="add-images.xlsx">add-images.xlsx</a> 中使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="bb972-113">Download the file <a href="add-images.xlsx">add-images.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-copy-an-image-across-worksheets"></a><span data-ttu-id="bb972-114">示例代码：跨工作表复制图像</span><span class="sxs-lookup"><span data-stu-id="bb972-114">Sample code: Copy an image across worksheets</span></span>

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

## <a name="sample-code-add-an-image-from-a-url-to-a-workbook"></a><span data-ttu-id="bb972-115">示例代码：将图像从 URL 添加到工作簿</span><span class="sxs-lookup"><span data-stu-id="bb972-115">Sample code: Add an image from a URL to a workbook</span></span>

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
  workbook.getWorksheet("WebSheet").addImage(image)
}

/**
 * Converts an ArrayBuffer containing a .png image into a base64-encoded string.
 */
function convertToBase64(input: ArrayBuffer) {
  const uInt8Array = new Uint8Array(input);
  const count = uInt8Array.length;

  // Allocate the necessary space up front.
  const charCodeArray = new Array(count) 
  
  // Convert every entry in the array to a character.
  for (let i = count; i >= 0; i--) { 
    charCodeArray[i] = String.fromCharCode(uInt8Array[i]);
  }

  // Convert the characters to base64.
  const base64 = btoa(charCodeArray.join(''));
  return base64;
}
```
