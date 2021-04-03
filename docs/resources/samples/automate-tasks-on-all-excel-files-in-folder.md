---
title: 对文件夹中的所有 Excel 文件运行脚本
description: 了解如何对 OneDrive for Business 上文件夹中的所有 Excel 文件运行脚本。
ms.date: 03/31/2021
localization_priority: Normal
ms.openlocfilehash: a11876e8241a069a7c640bbcf2c36b4842d3bd90
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571255"
---
# <a name="run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="09e9a-103">对文件夹中的所有 Excel 文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="09e9a-103">Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="09e9a-104">此项目对位于 OneDrive for Business 上的文件夹中的所有文件执行一组自动化任务。</span><span class="sxs-lookup"><span data-stu-id="09e9a-104">This project performs a set of automation tasks on all files situated in a folder on OneDrive for Business.</span></span> <span data-ttu-id="09e9a-105">还可以在 SharePoint 文件夹上使用。</span><span class="sxs-lookup"><span data-stu-id="09e9a-105">It could also be used on a SharePoint folder.</span></span>
<span data-ttu-id="09e9a-106">该代码对 Excel 文件执行计算，添加格式，并插入一个[@mentions注释。](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7)</span><span class="sxs-lookup"><span data-stu-id="09e9a-106">It performs calculations on the Excel files, adds formatting, and inserts a comment that [@mentions](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) a colleague.</span></span>

## <a name="sample-code-add-formatting-and-insert-comment"></a><span data-ttu-id="09e9a-107">示例代码：添加格式并插入注释</span><span class="sxs-lookup"><span data-stu-id="09e9a-107">Sample code: Add formatting and insert comment</span></span>

<span data-ttu-id="09e9a-108">下载文件 <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip，</a>将文件解压缩到本示例中使用的名为 **Sales** 的文件夹，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="09e9a-108">Download the file <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/highlight-alert-excel-files.zip?raw=true">highlight-alert-excel-files.zip</a>, extract the files to a folder titled **Sales** used in this sample, and try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  let table1 = workbook.getTable("Table1");
  const rowCount = table1.getRowCount();
  if (rowCount === 0) {
    return;
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const amountDueCol = table1.getColumnByName('Amount Due');
  const amountDueValues = amountDueCol.getRangeBetweenHeaderAndTotal().getValues();

  let highestValue = amountDueValues[0][0];
  let row = 0;
  for (let i = 1; i < amountDueValues.length; i++) {
    if (amountDueValues[i][0] > highestValue) {
      highestValue = amountDueValues[i][0];
      row = i;
    }
  }
  // Set fill color to FFFF00 for range in table Table1 cell in row 0 on column "Amount due".
  table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row)
    .getFormat()
    .getFill()
    .setColor("FFFF00");
  let selectedSheet = workbook.getActiveWorksheet();
  // Insert comment at cell InvoiceAmounts!F2.
  workbook.addComment(table1.getColumn("Amount due")
    .getRangeBetweenHeaderAndTotal()
    .getRow(row), {
    mentions: [{
      email: "AdeleV@M365x904181.OnMicrosoft.com",
      id: 0,
      name: "Adele Vance"
    }],
    richContent: "<at id=\"0\">Adele Vance</at> Please review this amount"
  }, ExcelScript.ContentType.mention);
}
```

## <a name="training-video-run-a-script-on-all-excel-files-in-a-folder"></a><span data-ttu-id="09e9a-109">培训视频：对文件夹中的所有 Excel 文件运行脚本</span><span class="sxs-lookup"><span data-stu-id="09e9a-109">Training video: Run a script on all Excel files in a folder</span></span>

<span data-ttu-id="09e9a-110">[观看分步视频，](https://youtu.be/xMg711o7k6w) 了解如何对 OneDrive for Business 或 SharePoint 文件夹中的所有 Excel 文件运行脚本。</span><span class="sxs-lookup"><span data-stu-id="09e9a-110">[Watch step-by-step video](https://youtu.be/xMg711o7k6w) on how to run a script on all Excel files in a OneDrive for Business or SharePoint folder.</span></span>
