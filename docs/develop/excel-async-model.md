---
title: 支持使用异步 Api 的较旧的 Office 脚本
description: Office 脚本异步 Api 的入门知识，以及如何对旧版脚本使用 load/sync 模式。
ms.date: 07/08/2020
localization_priority: Normal
ms.openlocfilehash: e7ca5b276cff0e3a38bffc2af1541c0051cf5490
ms.sourcegitcommit: ebd1079c7e2695ac0e7e4c616f2439975e196875
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45160458"
---
# <a name="support-older-office-scripts-that-use-the-async-apis"></a><span data-ttu-id="09c67-103">支持使用异步 Api 的较旧的 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="09c67-103">Support older Office Scripts that use the async APIs</span></span>

<span data-ttu-id="09c67-104">本文将教您如何维护和更新使用较旧模型的异步 Api 的脚本。</span><span class="sxs-lookup"><span data-stu-id="09c67-104">This article will teach you how to maintain and update scripts that use the older model's async APIs.</span></span> <span data-ttu-id="09c67-105">这些 Api 与 now-standard 同步 Office 脚本 Api 具有相同的核心功能，但它们要求您的脚本控制脚本和工作簿之间的数据同步。</span><span class="sxs-lookup"><span data-stu-id="09c67-105">These APIs have the same core functionality as the now-standard, synchronous Office Scripts APIs, but they require your script to control the data synchronization between the script and the workbook.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="09c67-106">异步模型仅可用于在实现当前[API 模型](scripting-fundamentals.md?view=office-scripts)之前创建的脚本。</span><span class="sxs-lookup"><span data-stu-id="09c67-106">The async model can only be used with scripts created before the implementation of the current [API model](scripting-fundamentals.md?view=office-scripts).</span></span> <span data-ttu-id="09c67-107">脚本将被永久锁定为它们创建时所拥有的 API 模型。</span><span class="sxs-lookup"><span data-stu-id="09c67-107">Scripts are permanently locked to the API model they have upon creation.</span></span> <span data-ttu-id="09c67-108">这也意味着，如果您想要将旧脚本转换为新模型，则必须创建全新的脚本。</span><span class="sxs-lookup"><span data-stu-id="09c67-108">This also means that if you want to convert an old script to the new model, you must create a brand new script.</span></span> <span data-ttu-id="09c67-109">我们建议您在进行更改时将旧脚本更新到新模型，因为当前模型更易于使用。</span><span class="sxs-lookup"><span data-stu-id="09c67-109">We recommend you update your old scripts to the new model when making changes, since the current model is easier to use.</span></span> <span data-ttu-id="09c67-110">将[异步脚本转换为 "当前模型"](#converting-async-scripts-to-the-current-model)部分包含有关如何进行此转换的建议。</span><span class="sxs-lookup"><span data-stu-id="09c67-110">The [Converting async scripts to the current model](#converting-async-scripts-to-the-current-model) section has advice on how to make this transition.</span></span>

## <a name="main-function"></a><span data-ttu-id="09c67-111">`main` 函数</span><span class="sxs-lookup"><span data-stu-id="09c67-111">`main` function</span></span>

<span data-ttu-id="09c67-112">使用异步 Api 的脚本具有不同的 `main` 函数。</span><span class="sxs-lookup"><span data-stu-id="09c67-112">Scripts that use the async APIs have a different `main` function.</span></span> <span data-ttu-id="09c67-113">它是一个 `async` 具有 `Excel.RequestContext` 作为第一个参数的函数。</span><span class="sxs-lookup"><span data-stu-id="09c67-113">It's an `async` function that has an `Excel.RequestContext` as the first parameter.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
    // Your async Office Script
}
```

## <a name="context"></a><span data-ttu-id="09c67-114">上下文</span><span class="sxs-lookup"><span data-stu-id="09c67-114">Context</span></span>

<span data-ttu-id="09c67-115">`main` 函数接受名为 `context` 的 `Excel.RequestContext` 参数。</span><span class="sxs-lookup"><span data-stu-id="09c67-115">The `main` function accepts an `Excel.RequestContext` parameter, named `context`.</span></span> <span data-ttu-id="09c67-116">将 `context` 视作脚本和工作簿之间的桥梁。</span><span class="sxs-lookup"><span data-stu-id="09c67-116">Think of `context` as the bridge between your script and the workbook.</span></span> <span data-ttu-id="09c67-117">脚本使用 `context` 对象访问工作簿，并使用该 `context` 来回发送数据。</span><span class="sxs-lookup"><span data-stu-id="09c67-117">Your script accesses the workbook with the `context` object and uses that `context` to send data back and forth.</span></span>

<span data-ttu-id="09c67-118">`context` 对象是必需的，因为脚本和 Excel 在不同的进程和位置中运行。</span><span class="sxs-lookup"><span data-stu-id="09c67-118">The `context` object is necessary because the script and Excel are running in different processes and locations.</span></span> <span data-ttu-id="09c67-119">该脚本将需要对云中的工作簿进行更改或从中查询数据。</span><span class="sxs-lookup"><span data-stu-id="09c67-119">The script will need to make changes to or query data from the workbook in the cloud.</span></span> <span data-ttu-id="09c67-120">`context` 对象管理以下事务。</span><span class="sxs-lookup"><span data-stu-id="09c67-120">The `context` object manages those transactions.</span></span>

## <a name="sync-and-load"></a><span data-ttu-id="09c67-121">同步和加载</span><span class="sxs-lookup"><span data-stu-id="09c67-121">Sync and Load</span></span>

<span data-ttu-id="09c67-122">因为脚本和工作簿在不同的位置运行，所以两者之间的任何数据传输都需要时间。</span><span class="sxs-lookup"><span data-stu-id="09c67-122">Because your script and workbook run in different locations, any data transfer between the two takes time.</span></span> <span data-ttu-id="09c67-123">在异步 API 中，命令将一直排队，直到脚本显式调用 `sync` 操作以同步脚本和工作簿。</span><span class="sxs-lookup"><span data-stu-id="09c67-123">In the async API, commands are queued up until the script explicitly calls the `sync` operation to synchronize the script and workbook.</span></span> <span data-ttu-id="09c67-124">脚本可以独立运行，直到需要执行以下任一操作：</span><span class="sxs-lookup"><span data-stu-id="09c67-124">Your script can work independently until it needs to do either of the following:</span></span>

- <span data-ttu-id="09c67-125">从工作簿中读取数据（遵循返回 [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) 的 `load` 操作或方法）。</span><span class="sxs-lookup"><span data-stu-id="09c67-125">Read data from the workbook (following a `load` operation or method that returns a [ClientResult](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async)).</span></span>
- <span data-ttu-id="09c67-126">将数据写入工作簿（通常是因为脚本已完成）。</span><span class="sxs-lookup"><span data-stu-id="09c67-126">Write data to the workbook (usually because the script has finished).</span></span>

<span data-ttu-id="09c67-127">下图显示了脚本和工作簿之间的示例控制流：</span><span class="sxs-lookup"><span data-stu-id="09c67-127">The following image shows an example control flow between the script and workbook:</span></span>

![该图显示了从脚本转到工作簿的读取和写入操作。](../images/load-sync.png)

### <a name="sync"></a><span data-ttu-id="09c67-129">同步</span><span class="sxs-lookup"><span data-stu-id="09c67-129">Sync</span></span>

<span data-ttu-id="09c67-130">只要异步脚本需要从工作簿中读取数据或将数据写入工作簿，请调用如下 `RequestContext.sync` 所示的方法：</span><span class="sxs-lookup"><span data-stu-id="09c67-130">Whenever your async script needs to read data from or write data to the workbook, call the `RequestContext.sync` method as shown here:</span></span>

```TypeScript
await context.sync();
```

> [!NOTE]
> <span data-ttu-id="09c67-131">脚本结束时将隐式调用 `context.sync()`。</span><span class="sxs-lookup"><span data-stu-id="09c67-131">`context.sync()` is implicitly called when a script ends.</span></span>

<span data-ttu-id="09c67-132">`sync` 操作完成后，工作簿将更新以反映脚本已指定的任何写入操作。</span><span class="sxs-lookup"><span data-stu-id="09c67-132">After the `sync` operation completes, the workbook updates to reflect any write operations that script has specified.</span></span> <span data-ttu-id="09c67-133">写操作是设置 Excel 对象（例如， `range.format.fill.color = "red"` ）或调用更改属性（如）的方法的任何属性 `range.format.autoFitColumns()` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-133">A write operation is setting any property on a Excel object (e.g., `range.format.fill.color = "red"`) or calling a method that changes a property (e.g., `range.format.autoFitColumns()`).</span></span> <span data-ttu-id="09c67-134">`sync` 操作还从脚本请求的工作簿中读取任何值，方式是通过使用能返回 `ClientResult` 的 `load` 操作或方法（如下一节所述）。</span><span class="sxs-lookup"><span data-stu-id="09c67-134">The `sync` operation also reads any values from the workbook that the script requested by using a `load` operation or a method that returns a `ClientResult` (as discussed in the next sections).</span></span>

<span data-ttu-id="09c67-135">将脚本与工作簿同步可能需要一些时间，具体取决于网络。</span><span class="sxs-lookup"><span data-stu-id="09c67-135">Synchronizing your script with the workbook can take time, depending on your network.</span></span> <span data-ttu-id="09c67-136">最大程度地减少 `sync` 用于帮助脚本运行速度的调用次数。</span><span class="sxs-lookup"><span data-stu-id="09c67-136">Minimize the number of `sync` calls to help your script run fast.</span></span> <span data-ttu-id="09c67-137">否则，异步 Api 不会更快地成为标准的同步 Api。</span><span class="sxs-lookup"><span data-stu-id="09c67-137">Otherwise, the async APIs are not faster the standard, synchronous APIs.</span></span>

### <a name="load"></a><span data-ttu-id="09c67-138">加载</span><span class="sxs-lookup"><span data-stu-id="09c67-138">Load</span></span>

<span data-ttu-id="09c67-139">异步脚本必须先从工作簿加载数据，然后再读取。</span><span class="sxs-lookup"><span data-stu-id="09c67-139">An async script must load data from the workbook before reading it.</span></span> <span data-ttu-id="09c67-140">但是，从整个工作簿中加载数据将极大地降低脚本速度。</span><span class="sxs-lookup"><span data-stu-id="09c67-140">However, loading data from the entire workbook would greatly reduce the script's speed.</span></span> <span data-ttu-id="09c67-141">此 `load` 方法使您的脚本明确声明应从工作簿中检索哪些数据。</span><span class="sxs-lookup"><span data-stu-id="09c67-141">The `load` method lets your script specifically state what data should be retrieved from the workbook.</span></span>

<span data-ttu-id="09c67-142">`load` 方法可用于每个 Excel 对象。</span><span class="sxs-lookup"><span data-stu-id="09c67-142">The `load` method is available on every Excel object.</span></span> <span data-ttu-id="09c67-143">脚本必须先加载对象的属性，然后才能读取它们。</span><span class="sxs-lookup"><span data-stu-id="09c67-143">Your script must load an object's properties before it can read them.</span></span> <span data-ttu-id="09c67-144">如果不这样做，则会导致错误。</span><span class="sxs-lookup"><span data-stu-id="09c67-144">Not doing so results in an error.</span></span>

<span data-ttu-id="09c67-145">下面的示例使用 `Range` 对象显示 `load` 方法可用于加载数据的三种方式。</span><span class="sxs-lookup"><span data-stu-id="09c67-145">The following examples use a `Range` object to show the three ways the `load` method can be used to load data.</span></span>

|<span data-ttu-id="09c67-146">意图</span><span class="sxs-lookup"><span data-stu-id="09c67-146">Intent</span></span> |<span data-ttu-id="09c67-147">示例命令</span><span class="sxs-lookup"><span data-stu-id="09c67-147">Example Command</span></span> | <span data-ttu-id="09c67-148">效果</span><span class="sxs-lookup"><span data-stu-id="09c67-148">Effect</span></span> |
|:--|:--|:--|
|<span data-ttu-id="09c67-149">加载一个属性</span><span class="sxs-lookup"><span data-stu-id="09c67-149">Load one property</span></span> |`myRange.load("values");` | <span data-ttu-id="09c67-150">加载单个属性，此例中为此范围内的二维值数组。</span><span class="sxs-lookup"><span data-stu-id="09c67-150">Loads a single property, in this case the two-dimensional array of values in this range.</span></span> |
|<span data-ttu-id="09c67-151">加载多个属性</span><span class="sxs-lookup"><span data-stu-id="09c67-151">Load multiple properties</span></span> |`myRange.load("values, rowCount, columnCount");`| <span data-ttu-id="09c67-152">从逗号分隔的列表中加载所有属性，此例中为值、行数和列数。</span><span class="sxs-lookup"><span data-stu-id="09c67-152">Loads all the properties from a comma-delimited list, in this example the values, row count, and column count.</span></span> |
|<span data-ttu-id="09c67-153">加载所有内容</span><span class="sxs-lookup"><span data-stu-id="09c67-153">Load everything</span></span> | `myRange.load();`|<span data-ttu-id="09c67-154">加载范围内的所有属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-154">Loads all the properties on the range.</span></span> <span data-ttu-id="09c67-155">这不是建议的解决方案，因为它会通过获取不必要的数据来降低脚本的速度。</span><span class="sxs-lookup"><span data-stu-id="09c67-155">This isn't a recommended solution, since it will slow down your script by getting unnecessary data.</span></span> <span data-ttu-id="09c67-156">仅在测试脚本时使用此属性，或者如果需要从对象中的每个属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-156">Only use this while testing your script or if you need every property from the object.</span></span> |

<span data-ttu-id="09c67-157">脚本必须先调用 `context.sync()`，然后才能读取任何加载的值。</span><span class="sxs-lookup"><span data-stu-id="09c67-157">Your script must call `context.sync()` before reading any loaded values.</span></span>

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

<span data-ttu-id="09c67-158">还可以在整个集合中加载属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-158">You can also load properties across an entire collection.</span></span> <span data-ttu-id="09c67-159">异步 API 中的每个集合对象都具有一个 `items` 属性，该属性是包含该集合中的对象的数组。</span><span class="sxs-lookup"><span data-stu-id="09c67-159">Every collection object in the async API has an `items` property that is an array containing the objects in that collection.</span></span> <span data-ttu-id="09c67-160">使用 `items` 作为对 `load` 的层次调用 (`items\myProperty`) 的开始，将在其中的每个项目上加载指定的属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-160">Using `items` as the start of a hierarchical call (`items\myProperty`) to `load` loads the specified properties on each of those items.</span></span> <span data-ttu-id="09c67-161">下面的示例在工作表的 `CommentCollection` 对象中的每个 `Comment` 对象上加载 `resolved` 属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-161">The following example loads the `resolved` property on every `Comment` object in the `CommentCollection` object of a worksheet.</span></span>

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

### <a name="clientresult"></a><span data-ttu-id="09c67-162">ClientResult</span><span class="sxs-lookup"><span data-stu-id="09c67-162">ClientResult</span></span>

<span data-ttu-id="09c67-163">异步 API 中从工作簿返回信息的方法的模式与范例中的方法类似 `load` / `sync` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-163">Methods in the async API that return information from the workbook have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="09c67-164">举个例子，`TableCollection.getCount`获取集合中的表的数量。</span><span class="sxs-lookup"><span data-stu-id="09c67-164">As an example, `TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="09c67-165">`getCount`返回 a `ClientResult<number>` ，表示 `value` 返回的属性 [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) 为数字。</span><span class="sxs-lookup"><span data-stu-id="09c67-165">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office-scripts/excelscript/excelscript.clientresult?view=office-scripts-async) is a number.</span></span> <span data-ttu-id="09c67-166">在调用 `context.sync()` 之前，脚本无法访问此值。</span><span class="sxs-lookup"><span data-stu-id="09c67-166">Your script can't access that value until `context.sync()` is called.</span></span> <span data-ttu-id="09c67-167">与加载属性很相似，直到 `sync` 调用，`value` 是本地 "空" 值。</span><span class="sxs-lookup"><span data-stu-id="09c67-167">Much like loading a property, the `value` is a local "empty" value until that `sync` call.</span></span>

<span data-ttu-id="09c67-168">以下脚本获取工作簿中的表的总数，并将该数目记录到控制台。</span><span class="sxs-lookup"><span data-stu-id="09c67-168">The following script gets the total number of tables in the workbook and logs that number to the console.</span></span>

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

## <a name="converting-async-scripts-to-the-current-model"></a><span data-ttu-id="09c67-169">将异步脚本转换为当前模型</span><span class="sxs-lookup"><span data-stu-id="09c67-169">Converting async scripts to the current model</span></span>

<span data-ttu-id="09c67-170">当前 API 模型不使用 `load` 、 `sync` 或 `RequestContext` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-170">The current API model doesn't use `load`, `sync`, or a `RequestContext`.</span></span> <span data-ttu-id="09c67-171">这使脚本更易于编写和维护。</span><span class="sxs-lookup"><span data-stu-id="09c67-171">This makes the scripts much easier to write and maintain.</span></span> <span data-ttu-id="09c67-172">转换旧脚本的最佳资源是[堆栈溢出](https://stackoverflow.com/questions/tagged/office-scripts)。</span><span class="sxs-lookup"><span data-stu-id="09c67-172">Your best resource for converting old scripts is [Stack Overflow](https://stackoverflow.com/questions/tagged/office-scripts).</span></span> <span data-ttu-id="09c67-173">在这里，你可以向社区寻求有关特定方案的帮助。</span><span class="sxs-lookup"><span data-stu-id="09c67-173">There, you can ask the community for help with specific scenarios.</span></span> <span data-ttu-id="09c67-174">以下指南应帮助概述你需要执行的常规步骤。</span><span class="sxs-lookup"><span data-stu-id="09c67-174">The following guidance should help outline the general steps you'll need to take.</span></span>

1. <span data-ttu-id="09c67-175">创建一个新脚本，并将旧的异步代码复制到该脚本中。</span><span class="sxs-lookup"><span data-stu-id="09c67-175">Create a new script and copy the old async code into it.</span></span> <span data-ttu-id="09c67-176">`main`请务必改用 current，而不要包含旧的方法签名 `function main(workbook: ExcelScript.Workbook)` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-176">Be sure not to include the old `main` method signature, using the current `function main(workbook: ExcelScript.Workbook)` instead.</span></span>

2. <span data-ttu-id="09c67-177">删除所有 `load` 和 `sync` 调用。</span><span class="sxs-lookup"><span data-stu-id="09c67-177">Remove all the `load` and `sync` calls.</span></span> <span data-ttu-id="09c67-178">不再需要它们。</span><span class="sxs-lookup"><span data-stu-id="09c67-178">They are no longer necessary.</span></span>

3. <span data-ttu-id="09c67-179">已删除所有属性。</span><span class="sxs-lookup"><span data-stu-id="09c67-179">All properties have been removed.</span></span> <span data-ttu-id="09c67-180">现在，您可以通过和方法访问这些对象 `get` `set` ，因此您需要将这些属性引用切换到方法调用。</span><span class="sxs-lookup"><span data-stu-id="09c67-180">You now access those objects through `get` and `set` methods, so you'll need to switch those property references to method calls.</span></span> <span data-ttu-id="09c67-181">例如， `mySheet.getRange("A2:C2").format.fill.color = "blue";` 您现在可以使用如下所示的方法，而不是通过属性访问来设置单元格的填充颜色：`mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span><span class="sxs-lookup"><span data-stu-id="09c67-181">For example, instead of setting a cell's fill color through property access like this: `mySheet.getRange("A2:C2").format.fill.color = "blue";`, you'll now use methods like this: `mySheet.getRange("A2:C2").getFormat().getFill().setColor("blue");`</span></span>

4. <span data-ttu-id="09c67-182">集合类已被数组替换。</span><span class="sxs-lookup"><span data-stu-id="09c67-182">Collection classes have been replaced by arrays.</span></span> <span data-ttu-id="09c67-183">`add` `get` 这些集合类的和方法被移至拥有集合的对象，因此必须相应地更新引用。</span><span class="sxs-lookup"><span data-stu-id="09c67-183">The `add` and `get` methods of those collection classes were moved to the object that owned the collection, so your references must be updated accordingly.</span></span> <span data-ttu-id="09c67-184">例如，若要从工作簿中的第一个工作表中获取一个名为 "MyChart" 的图表，请使用以下代码： `workbook.getWorksheets()[0].getChart("MyChart");` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-184">For example, to get a chart named "MyChart" from the first worksheet in the workbook, use the following code: `workbook.getWorksheets()[0].getChart("MyChart");`.</span></span> <span data-ttu-id="09c67-185">请注意， `[0]` 若要访问返回的返回的的第一个值 `Worksheet[]` `getWorksheets()` 。</span><span class="sxs-lookup"><span data-stu-id="09c67-185">Note the `[0]` to access the first value of the `Worksheet[]` returned by `getWorksheets()`.</span></span>

5. <span data-ttu-id="09c67-186">为清楚起见，一些方法已重命名，添加为方便。</span><span class="sxs-lookup"><span data-stu-id="09c67-186">Some methods have been renamed for clarity and added for convenience.</span></span> <span data-ttu-id="09c67-187">有关更多详细信息，请参阅[Office 脚本 API 参考](/javascript/api/office-scripts/overview?view=office-scripts)。</span><span class="sxs-lookup"><span data-stu-id="09c67-187">Please consult the [Office Scripts API reference](/javascript/api/office-scripts/overview?view=office-scripts) for more details.</span></span>

## <a name="office-scripts-async-api-reference-documentation"></a><span data-ttu-id="09c67-188">Office 脚本异步 API 参考文档</span><span class="sxs-lookup"><span data-stu-id="09c67-188">Office Scripts Async API reference documentation</span></span>

[!INCLUDE [Async reference documentation](../includes/async-reference-documentation-link.md)]
