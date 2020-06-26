---
title: 使用 Office 脚本异步 Api 支持旧版脚本
description: Office 脚本异步 Api 的入门知识，以及如何使用旧脚本的加载/同步模式。
ms.date: 06/22/2020
localization_priority: Normal
ms.openlocfilehash: c7b3c1401ecc2b4d0371590e71f61ae6e9ad8a9d
ms.sourcegitcommit: aec3c971c6640429f89b6bb99d2c95ea06725599
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2020
ms.locfileid: "44878734"
---
# <a name="using-the-office-scripts-async-apis-to-support-legacy-scripts"></a>使用 Office 脚本异步 Api 支持旧版脚本

本文将教您如何使用旧版、异步、Api 编写脚本。 这些 Api 与标准的同步 Office 脚本 Api 具有相同的核心功能，但它们要求您的脚本控制脚本和工作簿之间的数据同步。

> [!IMPORTANT]
> 异步模型仅可用于在实现当前[API 模型](scripting-fundamentals.md?view=office-scripts)之前创建的脚本。 脚本将被永久锁定为它们创建时所拥有的 API 模型。 这也意味着，如果您想要将旧脚本转换为新模型，则必须使用全新的脚本。 我们建议您在进行更改时将旧脚本更新到新模型，因为当前模型更易于使用。 将[旧的异步脚本转换为 "当前模型"](#converting-legacy-async-scripts-to-the-current-model)部分包含有关如何进行此转换的建议。

## <a name="main-function"></a>`main` 函数

使用异步 Api 的脚本具有不同的 `main` 函数。 它是一个 `async` 具有 `Excel.RequestContext` 作为第一个参数的函数。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>上下文

`main` 函数接受名为 `context` 的 `Excel.RequestContext` 参数。 将 `context` 视作脚本和工作簿之间的桥梁。 脚本使用 `context` 对象访问工作簿，并使用该 `context` 来回发送数据。

`context` 对象是必需的，因为脚本和 Excel 在不同的进程和位置中运行。 该脚本将需要对云中的工作簿进行更改或从中查询数据。 `context` 对象管理以下事务。

## <a name="sync-and-load"></a>同步和加载

因为脚本和工作簿在不同的位置运行，所以两者之间的任何数据传输都需要时间。 在异步 API 中，命令将一直排队，直到脚本显式调用 `sync` 操作以同步脚本和工作簿。 脚本可以独立运行，直到需要执行以下任一操作：

- 从工作簿中读取数据（遵循返回 [ClientResult](/javascript/api/office-scripts/excel/excel.clientresult?view=office-scripts-async) 的 `load` 操作或方法）。
- 将数据写入工作簿（通常是因为脚本已完成）。

下图显示了脚本和工作簿之间的示例控制流：

![该图显示了从脚本转到工作簿的读取和写入操作。](../images/load-sync.png)

### <a name="sync"></a>同步

只要异步脚本需要从工作簿中读取数据或将数据写入工作簿，请调用如下 `RequestContext.sync` 所示的方法：

```TypeScript
await context.sync();
```

> [!NOTE]
> 脚本结束时将隐式调用 `context.sync()`。

`sync` 操作完成后，工作簿将更新以反映脚本已指定的任何写入操作。 写入操作在 Excel 对象上设置任何属性（例如 `range.format.fill.color = "red"`），或调用更改属性的方法（例如 `range.format.autoFitColumns()`）。 `sync` 操作还从脚本请求的工作簿中读取任何值，方式是通过使用能返回 `ClientResult` 的 `load` 操作或方法（如下一节所述）。

将脚本与工作簿同步可能需要一些时间，具体取决于网络。 最大程度地减少 `sync` 用于帮助脚本运行速度的调用次数。 否则，异步 Api 不会更快地成为标准的同步 Api。

### <a name="load"></a>加载

异步脚本必须先从工作簿加载数据，然后再读取。 但是，从整个工作簿中加载数据将极大地降低脚本速度。 此 `load` 方法使您的脚本明确声明应从工作簿中检索哪些数据。

`load` 方法可用于每个 Excel 对象。 脚本必须先加载对象的属性，然后才能读取它们。 如果不这样做，则会导致错误。

下面的示例使用 `Range` 对象显示 `load` 方法可用于加载数据的三种方式。

|意图 |示例命令 | 效果 |
|:--|:--|:--|
|加载一个属性 |`myRange.load("values");` | 加载单个属性，此例中为此范围内的二维值数组。 |
|加载多个属性 |`myRange.load("values, rowCount, columnCount");`| 从逗号分隔的列表中加载所有属性，此例中为值、行数和列数。 |
|加载所有内容 | `myRange.load();`|加载范围内的所有属性。 这不是建议的解决方案，因为它会通过获取不必要的数据来降低脚本的速度。 仅在测试脚本时使用此属性，或者如果需要从对象中的每个属性。 |

脚本必须先调用 `context.sync()`，然后才能读取任何加载的值。

```TypeScript
/**
 * This script uses the async API to get the row count for a range.
 * It shows how to load a property in the async model.
 */
async function main(context: Excel.RequestContext) {
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let range = selectedSheet.getRange("A1:B3");

    // Load the property.
    range.load("rowCount");

    // Synchronize with the workbook to get the property.
    await context.sync();

    // Read and log the property value (3).
    console.log(range.rowCount);
}
```

还可以在整个集合中加载属性。 异步 API 中的每个集合对象都具有一个 `items` 属性，该属性是包含该集合中的对象的数组。 使用 `items` 作为对 `load` 的层次调用 (`items\myProperty`) 的开始，将在其中的每个项目上加载指定的属性。 下面的示例在工作表的 `CommentCollection` 对象中的每个 `Comment` 对象上加载 `resolved` 属性。

```TypeScript
/**
 * This script uses the async API to get resolved property on every comment in the worksheet.
 * It shows how to load a property from every object in a collection.
 */
async function main(context: Excel.RequestContext){
    let selectedSheet = context.workbook.worksheets.getActiveWorksheet();
    let comments = selectedSheet.comments;

    // Load the `resolved` property from every comment in this collection.
    comments.load("items/resolved");

    // Synchronize with the workbook to get the properties.
    await context.sync();
}
```

### <a name="clientresult"></a>ClientResult

异步 API 中从工作簿返回信息的方法的模式与范例中的方法类似 `load` / `sync` 。 举个例子，`TableCollection.getCount`获取集合中的表的数量。 `getCount` 返回 `ClientResult<number>`，这意味着返回 `ClientResult` 中的 `value` 属性为 "数字"。 在调用 `context.sync()` 之前，脚本无法访问此值。 与加载属性很相似，直到 `sync` 调用，`value` 是本地 "空" 值。

以下脚本获取工作簿中的表的总数，并将该数目记录到控制台。

```TypeScript
/**
 * This script uses the async API to get the table count of the workbook.
 * It shows how ClientResult objects return workbook information.
 */
async function main(context: Excel.RequestContext) {
    let tableCount = context.workbook.tables.getCount();

    // This sync call implicitly loads tableCount.value.
    // Any other ClientResult values are loaded too.
    await context.sync();

    // Trying to log the value before calling sync would throw an error.
    console.log(tableCount.value);
}
```

## <a name="converting-legacy-async-scripts-to-the-current-model"></a>将旧的异步脚本转换为当前模型

当前 API 模型不使用 `load` 、 `sync` 或 `RequestContext` 。 这使脚本更易于编写和维护。 转换旧脚本的最佳资源是[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)。 在这里，你可以向社区寻求有关特定方案的帮助。 以下指南应帮助概述你需要执行的常规步骤。

1. 创建一个新脚本，并将旧的异步代码复制到该脚本中。 `main`请务必改用 current，而不要包含旧的方法签名 `function main(workbook: ExcelScript.Workbook)` 。

2. 删除所有 `load` 和 `sync` 调用。 不再需要它们。

3. 已删除所有属性。 现在，您可以通过和方法访问这些对象 `get` `set` ，因此您需要将这些属性引用切换到方法调用。 例如， `mySheet.getRange("A2:C2").format.fill.color = "blue";` 您现在可以使用如下所示的方法，而不是通过属性访问来设置单元格的填充颜色：`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. 集合类已被数组替换。 `add` `get` 这些集合类的和方法被移至拥有集合的对象，因此必须相应地更新引用。 例如，若要从工作簿中的第一个工作表中获取一个名为 "MyChart" 的图表，请使用以下代码： `workbook.getWorksheets()[0].getChart("MyChart");` 。 请注意， `[0]` 若要访问返回的返回的的第一个值 `Worksheet[]` `getWorksheets()` 。

5. 为清楚起见，一些方法已重命名，添加为方便。 有关更多详细信息，请参阅[Office 脚本 API 参考](/javascript/api/office-scripts/overview?view=office-scripts)。

## <a name="office-scripts-async-api-reference-documentation"></a>Office 脚本异步 API 参考文档

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
