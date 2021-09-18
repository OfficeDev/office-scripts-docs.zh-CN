---
title: Office 脚本中的最佳实践
description: 如何防止常见问题，并编写可Office输入或数据的稳固脚本。
ms.date: 05/10/2021
ms.localizationpriority: medium
ms.openlocfilehash: c37559c978a04bd99fff044674b2f64b7758438b
ms.sourcegitcommit: 5ec904cbb1f2cc00a301a5ba7ccb8ae303341267
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/18/2021
ms.locfileid: "59447456"
---
# <a name="best-practices-in-office-scripts"></a>Office 脚本中的最佳实践

这些模式和做法旨在帮助脚本每次成功运行。 使用它们可以避免开始自动执行工作流时出现Excel错误。

## <a name="verify-an-object-is-present"></a>验证对象是否存在

脚本通常依赖于工作簿中呈现的某个工作表或表。 但是，在脚本运行之间，它们可能会重命名或删除。 通过先检查这些表或工作表是否存在，然后再对它们调用方法，您可以确保脚本不会突然结束。

以下示例代码检查工作簿中是否包含"索引"工作表。 如果工作表存在，脚本将获取一个范围并继续。 如果不存在，脚本将记录自定义错误消息。

```TypeScript
// Make sure the "Index" worksheet exists before using it.
let indexSheet = workbook.getWorksheet('Index');
if (indexSheet) {
  let range = indexSheet.getRange("A1");
  // Continue using the range...
} else {
  console.log("Index sheet not found.");
}
```

TypeScript `?` 运算符在调用方法之前检查对象是否存在。 如果无需在对象不存在时执行任何特殊操作，这可以使代码更加简化。

```TypeScript
// The ? ensures that the delete() API is only called if the object exists.
workbook.getWorksheet('Index')?.delete();
```

## <a name="validate-data-and-workbook-state-first"></a>首先验证数据和工作簿状态

处理数据之前，请确保存在所有工作表、表、形状和其他对象。 使用以前的模式，检查所有内容是否都位于工作簿中并符合您的预期。 在写入任何数据之前执行此操作可确保脚本不会使工作簿保持部分状态。

以下脚本要求存在两个名为"Table1"和"Table2"的表。 该脚本首先检查表是否存在，然后以 语句和相应的消息结尾（如果没有 `return` ）。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return;
  }

  // Continue...
}
```

如果验证发生在单独的函数中，您仍必须通过从 函数发出 语句 `return` 来结束 `main` 脚本。 从子函数返回不会结束脚本。

以下脚本具有与上一脚本相同的行为。 区别在于函数 `main` 调用 函数 `inputPresent` 来验证所有内容。 `inputPresent` 返回一个 boolean (`true` 或 `false`) 以指示是否存在所有必需的输入。 函数 `main` 使用该布尔值决定继续或结束脚本。

```TypeScript
function main(workbook: ExcelScript.Workbook) {

  // Get the table objects.
  if (!inputPresent(workbook)) {
    return;
  }

  // Continue...
}

function inputPresent(workbook: ExcelScript.Workbook): boolean {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    console.log(`Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`);
    return false;
  }

  return true;
}
```

## <a name="when-to-use-a-throw-statement"></a>何时使用 `throw` 语句

语句 [`throw`](https://developer.mozilla.org/docs/web/javascript/reference/statements/throw) 指示发生了意外错误。 它立即结束代码。 大多数情况下，不需要从 `throw` 脚本执行。 通常，脚本会自动通知用户脚本由于问题无法运行。 在大多数情况下，用一条错误消息和函数中的语句结束脚本 `return` `main` 就足够了。

但是，如果您的脚本作为流Power Automate运行，您可能需要阻止该流继续。 `throw`语句将停止脚本，并指示流也停止。

以下脚本显示如何使用表 `throw` 检查示例中的 语句。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // These tables must be in the workbook for the script.
  const TargetTableName = 'Table1';
  const SourceTableName = 'Table2';

  // Get the table objects.
  let targetTable = workbook.getTable(TargetTableName);
  let sourceTable = workbook.getTable(SourceTableName);

  // Check if the tables are there.
  if (!targetTable || !sourceTable) {
    // Immediately end the script with an error.
    throw `Required tables missing - Check that both the source (${TargetTableName}) and target (${SourceTableName}) tables are present before running the script.`;
  }
  
```

## <a name="when-to-use-a-trycatch-statement"></a>何时使用 `try...catch` 语句

语句 [`try...catch`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Statements/try...catch) 是检测 API 调用是否失败并继续运行脚本的方法。

请考虑以下对区域执行大型数据更新的代码段。

```TypeScript
range.setValues(someLargeValues);
```

如果 `someLargeValues` 大于Excel 网页版，调用 `setValues()` 将失败。 脚本随后也会失败，出现 [运行时错误](../testing/troubleshooting.md#runtime-errors)。 语句 `try...catch` 使脚本能够识别此情况，而不会立即结束脚本并显示默认错误。

为脚本用户提供更好的体验的一个方法是向用户显示自定义错误消息。 以下代码段显示了 `try...catch` 一个语句，它记录更多的错误信息，以更好地帮助读者。

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Please inspect and run again.`);
    console.log(error);
    return; // End the script (assuming this is in the main function).
}
```

处理错误的另一个方法是具有处理错误案例的回退行为。 以下代码段使用 `catch` 块尝试备用方法将更新分解为较小的部分，并避免错误。

> [!TIP]
> 有关如何更新较大区域的完整示例，请参阅编写 [大型数据集](../resources/samples/write-large-dataset.md)。

```TypeScript
try {
    range.setValues(someLargeValues);
} catch (error) {
    console.log(`The script failed to update the values at location ${range.getAddress()}. Trying a different approach.`);
    handleUpdatesInSmallerBatches(someLargeValues);
}

// Continue...
}
```

> [!NOTE]
> 在 `try...catch` 循环内部或周围使用会降低脚本的速度。 有关更多性能信息，请参阅 [避免使用 `try...catch` 块](web-client-performance.md#avoid-using-trycatch-blocks-in-or-surrounding-loops)。

## <a name="see-also"></a>另请参阅

- [Office 脚本疑难解答](../testing/troubleshooting.md)
- [有关使用脚本Power Automate疑Office信息](../testing/power-automate-troubleshooting.md)
- [Office 脚本的平台限制](../testing/platform-limits.md)
- [提高脚本Office性能](web-client-performance.md)
