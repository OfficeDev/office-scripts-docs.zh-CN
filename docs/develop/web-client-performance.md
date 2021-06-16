---
title: 提高脚本Office性能
description: 通过了解工作簿和脚本之间的通信Excel创建更快的脚本。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: a5bd879625b9c3bac0caa621dde312f7c961dd5c
ms.sourcegitcommit: 2aaf7dc527cb6c9f1206550b2c5745280503b2a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/16/2021
ms.locfileid: "52957698"
---
# <a name="improve-the-performance-of-your-office-scripts"></a>提高脚本Office性能

使用脚本Office自动执行一系列任务，以节省时间。 较慢的脚本可能会感觉它无法加快工作流的速度。 大多数情况下，脚本将完全正常并如期运行。 但是，有一些可避免的场景可能会影响性能。

脚本运行缓慢的最常见原因是与工作簿的通信过多。 当工作簿存在于云中时，脚本将在本地计算机上运行。 在某些时候，脚本会将其本地数据与工作簿的本地数据同步。 这意味着，当 (同步时，) 写入操作（如) ） `workbook.addWorksheet()` 都只应用于工作簿。 同样，任何读取操作 (，) 这些时间仅从脚本的 `myRange.getValues()` 工作簿获取数据。 在任一情况下，脚本都先提取信息，然后再处理数据。 例如，以下代码将准确记录已用区域中的行数。

```TypeScript
let usedRange = workbook.getActiveWorksheet().getUsedRange();
let rowCount = usedRange.getRowCount();
// The script will read the range and row count from
// the workbook before logging the information.
console.log(rowCount);
```

Office脚本 API 确保工作簿或脚本中任何数据都准确且在必要时是最新的。 无需担心这些同步，脚本就能够正常运行。 但是，了解此脚本到云通信可以帮助您避免不必要的网络调用。

## <a name="performance-optimizations"></a>性能优化

你可以应用简单的技术来帮助减少与云的通信。 以下模式有助于加快脚本速度。

- 读取工作簿数据一次，而不是在循环中重复读取。
- 删除不必要的 `console.log` 语句。
- 避免使用 try/catch 块。

### <a name="read-workbook-data-outside-of-a-loop"></a>在循环之外读取工作簿数据

从工作簿获取数据的任何方法都可以触发网络调用。 应尽量在本地保存数据，而不是重复进行同一调用。 在处理循环时尤其如此。

请考虑使用脚本获取工作表的已用区域中的负数计数。 脚本需要对已用区域内每个单元格进行访问。 为此，它需要范围、行数和列数。 您应该在启动循环之前，将那些变量存储为本地变量。 否则，循环的每个迭代将强制返回到工作簿。

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
> 作为实验，请尝试在 `usedRangeValues` 循环中用 替换 `usedRange.getValues()` 。 您可能会注意到，在处理较大范围时，脚本需要更长的时间运行。

### <a name="avoid-using-trycatch-blocks-in-or-surrounding-loops"></a>避免 `try...catch` 在循环或周围循环中使用块

我们不建议在循环 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 或周围循环中使用语句。 这是您避免在循环中读取数据的相同原因：每次迭代都强制脚本与工作簿同步，以确保未引发任何错误。 通过检查从工作簿返回的对象，可以避免大多数错误。 例如，以下脚本在尝试添加行之前检查工作簿返回的表是否存在。

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

### <a name="remove-unnecessary-consolelog-statements"></a>删除不必要的 `console.log` 语句

控制台日志记录是调试 [脚本的重要工具](../testing/troubleshooting.md)。 但是，它会强制脚本与工作簿同步，以确保记录的信息是最新的。 请考虑在共享脚本 (不必要的日志记录语句，例如) 测试日志的日志记录语句。 这通常不会导致明显的性能问题，除非语句 `console.log()` 位于循环中。

## <a name="case-by-case-help"></a>按案例帮助

随着 Office 脚本平台的扩展以使用[Power Automate、](https://flow.microsoft.com/)自适应卡片和其他跨产品[](/adaptive-cards)功能，脚本工作簿通信的细节变得更加复杂。 如果您需要有关加快脚本运行速度的帮助，请通过 [Microsoft Q&A 联系](/answers/topics/office-scripts-excel-dev.html)。 请务必使用"office-scripts-dev"标记你的问题，以便专家可以找到它并提供帮助。

## <a name="see-also"></a>另请参阅

- [Excel 网页版中 Office 脚本的脚本基础知识](scripting-fundamentals.md)
- [MDN Web 文档：循环和迭代](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Loops_and_iteration)
