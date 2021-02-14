---
title: 支持使用异步 API 的较旧 Office 脚本
description: 有关 Office 脚本异步 API 以及如何对较旧的脚本使用加载/同步模式的一本本。
ms.date: 02/08/2021
localization_priority: Normal
ms.openlocfilehash: be7847efe59dc6026875b8a8e3b3c93e0eb82e4d
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242023"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a>支持使用异步 API 的较旧 Office 脚本

本文将指导你如何维护和更新使用旧模型的异步 API 的脚本。 这些 API 具有与现在标准的同步 Office 脚本 API 相同的核心功能，但它们要求脚本控制脚本和工作簿之间的数据同步。

> [!IMPORTANT]
> 异步模型只能用于实现当前 API 模型之前创建的 [脚本](scripting-fundamentals.md?view=office-scripts&preserve-view=true)。 脚本被永久锁定到它们创建时具有的 API 模型。 这也意味着如果要将旧脚本转换为新模型，则必须创建全新的脚本。 我们建议你在进行更改时将旧脚本更新到新模型，因为当前模型更易于使用。 " [将异步脚本转换为当前模型](#converting-async-scripts-to-the-current-model) "部分提供了如何进行此转换的建议。

## <a name="main-function"></a>`main` 函数

使用异步 API 的脚本具有不同的 `main` 函数。 它是一 `async` 个作为第一 `Excel.RequestContext` 个参数的函数。

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a>上下文

`main` 函数接受名为 `context` 的 `Excel.RequestContext` 参数。 将 `context` 视作脚本和工作簿之间的桥梁。 脚本使用 `context` 对象访问工作簿，并使用该 `context` 来回发送数据。

`context` 对象是必需的，因为脚本和 Excel 在不同的进程和位置中运行。 该脚本将需要对云中的工作簿进行更改或从中查询数据。 `context` 对象管理以下事务。

## <a name="sync-and-load"></a>同步和加载

因为脚本和工作簿在不同的位置运行，所以两者之间的任何数据传输都需要时间。 在异步 API 中，命令将排入队列，直到脚本显式调用操作以 `sync` 同步脚本和工作簿。 脚本可以独立运行，直到需要执行以下任一操作：

- 从工作簿中读取数据（遵循返回 [ClientResult](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) 的 `load` 操作或方法）。
- 将数据写入工作簿（通常是因为脚本已完成）。

下图显示了脚本和工作簿之间的示例控制流：

![该图显示了从脚本转到工作簿的读取和写入操作。](../images/load-sync.png)

### <a name="sync"></a>同步

每当异步脚本需要读取工作簿数据或将数据写入工作簿时，调用 `RequestContext.sync` 方法，如下所示：

```TypeScript
await context.sync();
```

> [!NOTE]
> 脚本结束时将隐式调用 `context.sync()`。

`sync` 操作完成后，工作簿将更新以反映脚本已指定的任何写入操作。 写入操作在 Excel 对象上设置任何属性 (例如，) 或调用更改属性的方法 (例如 `range.format.fill.color = "red"` `range.format.autoFitColumns()`) 。 `sync` 操作还从脚本请求的工作簿中读取任何值，方式是通过使用能返回 `ClientResult` 的 `load` 操作或方法（如下一节所述）。

将脚本与工作簿同步可能需要一些时间，具体取决于网络。 尽量减少调用 `sync` 次数以帮助脚本快速运行。 否则，异步 API 不是标准同步 API 的速度更快。

### <a name="load"></a>加载

异步脚本必须先从工作簿加载数据，然后才能读取数据。 但是，从整个工作簿加载数据会大大降低脚本的速度。 `load`此方法使脚本可以专门指出应从工作簿检索哪些数据。

`load` 方法可用于每个 Excel 对象。 脚本必须先加载对象的属性，然后才能读取它们。 否则会导致错误。

下面的示例使用 `Range` 对象显示 `load` 方法可用于加载数据的三种方式。

|意图 |示例命令 | 效果 |
|:--|:--|:--|
|加载一个属性 |`myRange.load("values");` | 加载单个属性，此例中为此范围内的二维值数组。 |
|加载多个属性 |`myRange.load("values, rowCount, columnCount");`| 从逗号分隔的列表中加载所有属性，此例中为值、行数和列数。 |
|加载所有内容 | `myRange.load();`|加载范围内的所有属性。 这不是建议的解决方案，因为它会通过获取不必要的数据来减慢脚本的速度。 仅在测试脚本或需要对象中的每个属性时使用此参数。 |

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

还可以在整个集合中加载属性。 异步 API 中的每个集合对象都有一个属性，该属性是包含该集合 `items` 中的对象的数组。 使用 `items` 作为对 `load` 的层次调用 (`items\myProperty`) 的开始，将在其中的每个项目上加载指定的属性。 下面的示例在工作表的 `CommentCollection` 对象中的每个 `Comment` 对象上加载 `resolved` 属性。

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

从工作簿返回信息的异步 API 中的方法具有与范例类似的 `load` / `sync` 模式。 举个例子，`TableCollection.getCount`获取集合中的表的数量。 `getCount` 返回 `ClientResult<number>` 一个 ，表示 `value` 返回的属性 [`ClientResult`](/javascript/api/office/officeextension.clientresult?view=excel-js-online&preserve-view=true) 是一个数字。 在调用 `context.sync()` 之前，脚本无法访问此值。 与加载属性很相似，直到 `sync` 调用，`value` 是本地 "空" 值。

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

## <a name="converting-async-scripts-to-the-current-model"></a>将异步脚本转换为当前模型

当前 API 模型不使用 `load` 、 `sync` 或 `RequestContext` . 这使脚本更易于编写和维护。 转换旧脚本的最佳资源是 Stack [Overflow。](https://stackoverflow.com/questions/tagged/office-scripts) 你可以向社区请求特定方案的帮助。 以下指南应有助于概述需要执行的一般步骤。

1. 创建新脚本，将旧异步代码复制到该脚本中。 请务必不要包含旧的 `main` 方法签名，而应改为使用 `function main(workbook: ExcelScript.Workbook)` 当前方法签名。

2. 删除所有 `load` 和 `sync` 呼叫。 不再需要它们。

3. 已删除所有属性。 现在，您通过和方法访问这些对象，因此您需要将那些属性引用切换到 `get` `set` 方法调用。 例如，现在将使用以下方法，而不是通过属性访问设置单元格的填充 `mySheet.getRange("A2:C2").format.fill.color = "blue";` 颜色： `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`

4. 集合类已被数组取代。 这些 `add` 集合类的方法和方法已移动到拥有集合的对象，因此必须相应地更新 `get` 引用。 例如，若要从工作簿的第一个工作表获取名为"MyChart"的图表，请使用以下 `workbook.getWorksheets()[0].getChart("MyChart");` 代码： 请注意 `[0]` ，要访问由 返回 `Worksheet[]` 的第一个值 `getWorksheets()` 。

5. 为了清楚起见，一些方法已重命名并添加为方便使用。 有关详细信息，请参阅 [Office 脚本 API](/javascript/api/office-scripts/overview?view=office-scripts&preserve-view=true) 参考。

## <a name="office-scripts-async-api-reference-documentation"></a>Office 脚本异步 API 参考文档

异步 API 与 Office 外接程序中使用的 API 等效。参考文档位于 Office 加载项 [JavaScript API](/javascript/api/excel?view=excel-js-online&preserve-view=true)参考的 Excel 部分。
