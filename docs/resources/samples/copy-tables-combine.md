---
title: 将多个 Excel 表中的数据合并到单个表中
description: 了解如何使用 Office 脚本将多个 Excel 表中的数据合并到单个表中。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3db510514c676b9012fd47abc2a7e92492a9cf87
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572449"
---
# <a name="combine-data-from-multiple-excel-tables-into-a-single-table"></a>将多个 Excel 表中的数据合并到单个表中

此示例将多个 Excel 表中的数据合并到包含所有行的单个表中。 它假定正在使用的所有表都具有相同的结构。

此脚本有两种变体：

1. [第一个脚本](#sample-code-combine-data-from-multiple-excel-tables-into-a-single-table)合并了 Excel 文件中的所有表。
1. 第 [二个脚本](#sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table) 有选择地获取一组工作表中的表。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的tables-copy.xlsx](tables-copy.xlsx) 。 添加以下脚本以自行尝试示例！

## <a name="sample-code-combine-data-from-multiple-excel-tables-into-a-single-table"></a>示例代码：将多个 Excel 表中的数据合并到单个表中

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

## <a name="sample-code-combine-data-from-multiple-excel-tables-in-select-worksheets-into-a-single-table"></a>示例代码：将选定工作表中多个 Excel 表中的数据合并到单个表中

下载示例文件 [tables-select-copy.xlsx](tables-select-copy.xlsx) 并将其与以下脚本一起使用，以便自己试用！

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

## <a name="training-video-combine-data-from-multiple-excel-tables-into-a-single-table"></a>培训视频：将多个 Excel 表中的数据合并到单个表中

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/di-8JukK3Lc)。
