---
title: 将多个数据表中的Excel组合到一个表中
description: 了解如何使用 Office 脚本将多个Excel表中的数据合并到一个表中。
ms.date: 05/03/2021
localization_priority: Normal
ms.openlocfilehash: 2b9bb4d0db2ddd67e1cba10dbff707c59ea27501
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285918"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="70c06-103">将多个数据表中的Excel组合到一个表中</span><span class="sxs-lookup"><span data-stu-id="70c06-103">Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="70c06-104">此示例将来自多个Excel的数据组合到一个包含所有行的表中。</span><span class="sxs-lookup"><span data-stu-id="70c06-104">This sample combines data from multiple Excel tables into a single table that includes all the rows.</span></span> <span data-ttu-id="70c06-105">它假定使用的所有表都具有相同的结构。</span><span class="sxs-lookup"><span data-stu-id="70c06-105">It assumes that all tables being used have the same structure.</span></span>

<span data-ttu-id="70c06-106">此脚本有两种变体：</span><span class="sxs-lookup"><span data-stu-id="70c06-106">There are two variations of this script:</span></span>

1. <span data-ttu-id="70c06-107">第[一个](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)脚本将合并该脚本文件Excel表。</span><span class="sxs-lookup"><span data-stu-id="70c06-107">The [first script](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table) combines all tables in the Excel file.</span></span>
1. <span data-ttu-id="70c06-108">第 [二个](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) 脚本有选择地获取一组工作表中的表。</span><span class="sxs-lookup"><span data-stu-id="70c06-108">The [second script](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) selectively gets tables within a set of worksheets.</span></span>

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="70c06-109">示例代码：将数据从多个Excel组合到一个表中</span><span class="sxs-lookup"><span data-stu-id="70c06-109">Sample code: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="70c06-110">下载示例文件 <a href="tables-copy.xlsx">tables-copy.xlsx</a> 并使用以下脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="70c06-110">Download the sample file <a href="tables-copy.xlsx">tables-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');
  
  // Get the header values for the first table in the workbook.
  // This also saves the table list before we add the new, combined table.
  const tables = workbook.getTables();    
  const headerValues = tables[0].getHeaderRowRange().getTexts();
  console.log(headerValues);

  // Copy the headers on a new worksheet to an equal-sized range.
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);

  // Add the data from each table in the workbook to the new table.
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);
  for (let table of tables) {      
    let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
    let rowCount = table.getRowCount();

    // If the table is not empty, add its rows to the combined table.
    if (rowCount > 0) {
      combinedTable.addRows(-1, dataValues);
    }
  }
}
```

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a><span data-ttu-id="70c06-111">示例代码：将选定工作表Excel多个数据表的数据合并到一个表中</span><span class="sxs-lookup"><span data-stu-id="70c06-111">Sample code: Combine data from multiple Excel tables in select worksheets into a single table</span></span>

<span data-ttu-id="70c06-112">下载示例文件 <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> 并使用以下脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="70c06-112">Download the sample file <a href="tables-select-copy.xlsx">tables-select-copy.xlsx</a> and use it with the following script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Set the worksheet names to get tables from.
  const sheetNames = ['Sheet1', 'Sheet2', 'Sheet3'];
    
  // Delete the "Combined" worksheet, if it's present.
  workbook.getWorksheet('Combined')?.delete();

  // Create a new worksheet named "Combined" for the combined table.
  const newSheet = workbook.addWorksheet('Combined');

  // Create a new table with the same headers as the other tables.
  const headerValues = workbook.getWorksheet(sheetNames[0]).getTables()[0].getHeaderRowRange().getTexts();
  const targetRange = newSheet.getRange('A1').getResizedRange(headerValues.length-1, headerValues[0].length-1);
  targetRange.setValues(headerValues);
  const combinedTable = newSheet.addTable(targetRange.getAddress(), true);

  // Go through each listed worksheet and get their tables.
  sheetNames.forEach((sheet) => {
    const tables = workbook.getWorksheet(sheet).getTables();     
    for (let table of tables) {
      // Get the rows from the tables.
      let dataValues = table.getRangeBetweenHeaderAndTotal().getTexts();
      let rowCount = table.getRowCount();

      // If there's data in the table, add it to the combined table.
      if (rowCount > 0) {
          combinedTable.addRows(-1, dataValues);
      }
    }
  });
}
```

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a><span data-ttu-id="70c06-113">培训视频：将数据从多个Excel表组合到一个表中</span><span class="sxs-lookup"><span data-stu-id="70c06-113">Training video: Combine data from multiple Excel tables into a single table</span></span>

<span data-ttu-id="70c06-114">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/di-8JukK3Lc)。</span><span class="sxs-lookup"><span data-stu-id="70c06-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/di-8JukK3Lc).</span></span>
