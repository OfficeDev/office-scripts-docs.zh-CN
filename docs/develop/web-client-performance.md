---
title: 提高 Office 脚本的性能
description: 通过了解 Excel 工作簿和脚本之间的通信来创建更快的脚本。
ms.date: 06/15/2020
localization_priority: Normal
ms.openlocfilehash: 4d5b7c70f14e3fc598b95a6226e3ef8caf89f651
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878749"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>提高 Office 脚本的性能

Office 脚本的用途是自动化通常执行的一系列任务以节省时间。 较慢的脚本可能感觉不会加快工作流的速度。 大多数情况下，您的脚本完全正常，并按预期运行。 但是，有几个可能会影响性能的 avoidable 方案。

速度较慢的脚本的最常见原因是与工作簿之间的通信过多。 您的脚本在本地计算机上运行，而该工作簿存在于云中。 在某些情况下，您的脚本会将其本地数据与工作簿的同步。 这意味着， `workbook.addWorksheet()` 在这种幕后同步发生时，任何写操作（如）都仅适用于工作簿。 同样，任何读取操作（例如 `myRange.getValues()` ）仅在这些时间从脚本的工作簿中获取数据。 在这两种情况下，脚本都会在对数据进行操作之前提取信息。 例如，以下代码将准确记录所用区域中的行数。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office 脚本 Api 可确保工作簿或脚本中的任何数据在必要时都是准确且最新的。 您无需担心这些同步即可正确运行脚本。 但是，对此脚本到云通信的感知可帮助您避免不需要的网络调用。

## <a name="performance-optimizations"></a>性能优化

您可以应用简单的技术来帮助减少与云的通信。 下面的模式可帮助您提高脚本速度。

- 读取一次工作簿数据，而不是重复执行循环。
- 删除不必要 `console.log` 的语句。
- 避免使用 try/catch 块。

### <a name="read-workbook-data-outside-of-a-loop"></a>读取循环外部的工作簿数据

从工作簿中获取数据的任何方法都可以触发网络调用。 应尽可能在本地保存数据，而不是反复进行相同的调用。 在处理循环时尤其如此。

考虑使用脚本来获取工作表的所用区域中的负数计数。 脚本需要循环访问所使用区域中的每个单元格。 若要执行此操作，它需要范围、行数和列数。 在开始循环之前，应将这些变量存储为局部变量。 否则，循环的每个迭代都将强制返回工作簿。

```TypeScript
/**
 * This script provides the count of negative numbers that are present
 * in the used range of the current worksheet.
 */
function main(workbook: ExcelScript.Workbook) {
  // Get the working range.
  let usedRange = workbook.getActiveWorksheet().getUsedRange();

  // Save the values locally to avoid repeatedly asking the workbook.
  let usedRangeValues = usedRange.getValues();

  // Start the negative number counter.
  let negativeCount = 0;

  // Iterate over the entire range looking for negative numbers.
  for (let i = 0; i < usedRangeValues.length; i++) {
    for (let j = 0; j < usedRangeValues[i].length; j++) {
      if (usedRangeValues[i][j] < 0) {
        negativeCount++;
      }
    }
  }

  // Log the negative number count to the console.
  console.log(negativeCount);
}
```

> [!NOTE]
> 作为实验，请尝试 `usedRangeValues` 将循环中的替换为 `usedRange.getValues()` 。 您可能会注意到，在处理大型区域时脚本运行时间要长得多。

### <a name="remove-unnecessary-consolelog-statements"></a>删除不必要的 `console.log` 语句

控制台日志记录是[调试脚本](../testing/troubleshooting.md)的重要工具。 但是，它确实强制脚本与工作簿同步，以确保记录的信息是最新的。 在共享脚本之前，请考虑删除不必要的日志记录语句（如用于测试的日志记录语句）。 除非语句在循环中，否则通常不会引起显著的性能问题 `console.log()` 。

### <a name="avoid-using-trycatch-blocks"></a>避免使用 try/catch 块

建议不要将[ `try` / `catch` 块](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch)用作脚本的预期控制流的一部分。 通过检查从工作簿返回的对象，可以避免大多数错误。 例如，下面的脚本在尝试添加行之前检查工作簿返回的表是否存在。

```TypeScript
/**
 * This script adds a row to "MyTable", if that table is present.
 */
function main(workbook: ExcelScript.Workbook) {
  let table = workbook.getTable("MyTable");

  // Check if the table exists.
  if (table) {
    // Add the row.
    table.addRow(-1, ["2012", "Yes", "Maybe"]);
  } else {
    // Report the missing table.
    console.log("MyTable not found.");
  }
}
```

## <a name="case-by-case-help"></a>按大小写帮助

随着 Office 脚本平台扩展以配合使用[电源自动化](https://flow.microsoft.com/)、[自适应卡](https://docs.microsoft.com/adaptive-cards)和其他跨产品功能，脚本工作簿通信的详细信息会变得更加复杂。 如果需要帮助使脚本运行得更快，请通过[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)。 请务必使用 "office-scripts" 标记你的问题，以便专家可以找到它和帮助。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中 Office 脚本的脚本基础](scripting-fundamentals.md)
- [MDN web 文档：循环和迭代](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
