---
title: 将多个 Excel 表中的数据合并到单个表中
description: 了解如何使用 Office 脚本将多个 Excel 表中的数据合并到一个表中。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 2f3f7232216f686946861d8c2cdec44013333ec7
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571207"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="ca4f3-103">将多个 Excel 表中的数据合并到单个表中</span><span class="sxs-lookup"><span data-stu-id="ca4f3-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="ca4f3-104">此示例将多个 Excel 表中的数据合并到一个包含所有行的表中。</span><span class="sxs-lookup"><span data-stu-id="ca4f3-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="ca4f3-105">它假定使用的所有表都具有相同的结构。</span><span class="sxs-lookup"><span data-stu-id="ca4f3-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="ca4f3-106">此脚本有两种变体：</span><span class="sxs-lookup"><span data-stu-id="ca4f3-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="ca4f3-107">第 [一个](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) 脚本将 Excel 文件的所有表组合在一起。</span><span class="sxs-lookup"><span data-stu-id="ca4f3-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="ca4f3-108">第 [二个](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) 脚本有选择地获取一组工作表中的表。</span><span class="sxs-lookup"><span data-stu-id="ca4f3-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="ca4f3-109">示例代码：将多个 Excel 表中的数据合并到一个表中</span><span class="sxs-lookup"><span data-stu-id="ca4f3-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="ca4f3-110">下载示例文件 <a href="tables-copy.xlsx">tables-copy.xlsx</a> 并使用以下脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="ca4f3-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    
    const tables = workbook.getTables();    
    const headerValues = tables[0].getHeaderRowRange().getTexts();
    console.log(headerValues);
    const targetRange = updateRange(newSheet, headerValues);
    const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
    for (let table of tables) {      
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();
      if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
      }
    }
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="ca4f3-111">示例代码：将选定工作表中多个 Excel 表中的数据合并到单个表中</span><span class="sxs-lookup"><span data-stu-id="ca4f3-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="ca4f3-112">下载示例文件 <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> 并使用以下脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="ca4f3-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
    workbook.getWorksheet('Combined')?.delete();
    const newSheet = workbook.addWorksheet('Combined');
    let targetTableCreated = false;
    let combinedTable;
    sheetNames.forEach((sheet) => {
      const tables = workbook.getWorksheet(sheet).getTables();
      if (!targetTableCreated) {
        const headerValues = tables[0].getHeaderRowRange().getTexts();
        const targetRange = updateRange(newSheet, headerValues);
        combinedTable = newSheet.addTable(targetRange.getAddress(), true);
        targetTableCreated = true;
      }      
      for (let table of tables) {
        let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
        let rowCount = table.getRowCount();
        if (rowCount > 0) {
        combinedTable.addRows(-1, dataValues);
        }
      }
    })
}

function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
  const targetRange = sheet.getRange('A1').getResizedRange(data.length-1, data[0].length-1);
  targetRange.setValues(data);
  return targetRange;
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="ca4f3-113">培训视频：将多个 Excel 表中的数据合并到一个表中</span><span class="sxs-lookup"><span data-stu-id="ca4f3-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="ca4f3-114">[![观看分步视频，了解如何将多个 Excel 表中的数据合并到单个表中](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "如何将多个 Excel 表中的数据合并到单个表中的分步视频")</span><span class="sxs-lookup"><span data-stu-id="ca4f3-114">[![Watch step-by-step video on how to combine data from multiple Excel tables into a single table](../../images/merge-tables-vid.jpg)](https://youtu.be/di-8JukK3Lc "Step-by-step video on how to combine data from multiple Excel tables into a single table")</span></span>
