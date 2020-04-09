---
title: Web 上的 Excel 中 Office 脚本的示例脚本
description: 要用于 web 上 Excel 中的 Office 脚本的一组代码示例。
ms.date: 04/06/2020
localization_priority: Normal
ms.openlocfilehash: abf6b87b63ad027cca8ee5c947b687f54815409c
ms.sourcegitcommit: 0b2232c4c228b14d501edb8bb489fe0e84748b42
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/08/2020
ms.locfileid: "43191008"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a>Excel 网页版中 Office 脚本的示例脚本（预览）

下面的示例是您在自己的工作簿中尝试的简单脚本。 若要在 Excel 网页上使用它们，请执行以下操作：

1. 打开“**自动**”选项卡。
2. 按**代码编辑器**。
3. 在代码编辑器的任务窗格中，按 "**新建脚本**"。
4. 将整个脚本替换为您选择的示例。
5. 在代码编辑器的任务窗格中按 "**运行**"。

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a>脚本基础

这些示例演示 Office 脚本的基本构建基块。 将这些应用程序添加到脚本以扩展解决方案并解决常见问题。

### <a name="read-and-log-one-cell"></a>读取和记录一个单元格

此示例读取**A1**的值并将其打印到控制台。

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a>使用日期

本节中的示例演示如何使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)对象。

下面的示例获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

下一个示例读取存储在 Excel 中的日期，并将其转换为 JavaScript Date 对象。 它使用[日期的数字序列号](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46)作为 JavaScript 日期的输入。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Read a date at cell A1 from Excel.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  dateRange.load("values");
  await context.sync();

  // Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.values[0][0];
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## <a name="display-data"></a>显示数据

这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。

### <a name="apply-conditional-formatting"></a>应用条件格式

此示例向工作表中当前使用的区域应用条件格式。 条件格式是前10% 的数值的绿色填充。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a>创建已排序的表

本示例从当前工作表的已用区域创建一个表格，然后基于第一列对其进行排序。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a>协作

这些示例演示如何使用 Excel 的与协作相关的功能，如注释。

### <a name="delete-resolved-comments"></a>删除已解决的注释

此示例从当前工作表中删除所有已解析的注释。

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a>方案示例

有关 showcasing 大型的真实解决方案的示例，请访问[Office 脚本的示例方案](scenarios/sample-scenario-overview.md)。

## <a name="suggest-new-samples"></a>建议新示例

我们欢迎您提出新示例建议。 如果有一个可帮助其他脚本开发人员的常见方案，请在下面的 "反馈" 部分告诉我们。
