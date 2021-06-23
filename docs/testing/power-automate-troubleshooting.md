---
title: 对Office中运行的脚本进行Power Automate
description: 使用技巧脚本和脚本之间的集成时，Office、平台信息和Power Automate。
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 7ba128314c0d632a3e77792b7ee545bfb7dca71d
ms.sourcegitcommit: 4693c8f79428ec74695328275703af0ba1bfea8f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53074632"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a><span data-ttu-id="6318e-103">对Office中运行的脚本进行Power Automate</span><span class="sxs-lookup"><span data-stu-id="6318e-103">Troubleshoot Office Scripts running in Power Automate</span></span>

<span data-ttu-id="6318e-104">Power Automate，你可以将Office脚本自动化上一个级别。</span><span class="sxs-lookup"><span data-stu-id="6318e-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="6318e-105">但是，Power Automate在独立会话中代表您Excel脚本，因此有一些重要的注意事项。</span><span class="sxs-lookup"><span data-stu-id="6318e-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="6318e-106">如果你刚开始将 Office 脚本与 Power Automate 一起Power Automate运行 Office [Scripts with Power Automate](../develop/power-automate-integration.md)了解平台。</span><span class="sxs-lookup"><span data-stu-id="6318e-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-relative-references"></a><span data-ttu-id="6318e-107">避免相对引用</span><span class="sxs-lookup"><span data-stu-id="6318e-107">Avoid relative references</span></span>

<span data-ttu-id="6318e-108">Power Automate代表您Excel所选工作簿中运行脚本。</span><span class="sxs-lookup"><span data-stu-id="6318e-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="6318e-109">发生这种情况时，工作簿可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="6318e-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="6318e-110">任何依赖用户当前状态（如 ）的 API 在用户 `Workbook.getActiveWorksheet` Power Automate。</span><span class="sxs-lookup"><span data-stu-id="6318e-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="6318e-111">这是因为 API 基于用户视图或游标的相对位置，并且该引用不存在于Power Automate流中。</span><span class="sxs-lookup"><span data-stu-id="6318e-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="6318e-112">某些相对引用 API 在Power Automate。</span><span class="sxs-lookup"><span data-stu-id="6318e-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="6318e-113">其他人有一个默认行为，表示用户的状态。</span><span class="sxs-lookup"><span data-stu-id="6318e-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="6318e-114">在设计脚本时，请确保对工作表和范围使用绝对引用。</span><span class="sxs-lookup"><span data-stu-id="6318e-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="6318e-115">这样，即使Power Automate重新排列，也使工作表流保持一致。</span><span class="sxs-lookup"><span data-stu-id="6318e-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a><span data-ttu-id="6318e-116">在流中运行时失败的Power Automate方法</span><span class="sxs-lookup"><span data-stu-id="6318e-116">Script methods that fail when run in Power Automate flows</span></span>

<span data-ttu-id="6318e-117">以下方法引发错误，在从脚本流中的脚本调用时Power Automate失败。</span><span class="sxs-lookup"><span data-stu-id="6318e-117">The following methods throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="6318e-118">类</span><span class="sxs-lookup"><span data-stu-id="6318e-118">Class</span></span> | <span data-ttu-id="6318e-119">方法</span><span class="sxs-lookup"><span data-stu-id="6318e-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="6318e-120">Chart</span><span class="sxs-lookup"><span data-stu-id="6318e-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="6318e-121">区域</span><span class="sxs-lookup"><span data-stu-id="6318e-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="6318e-122">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="6318e-123">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="6318e-124">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="6318e-125">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="6318e-126">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="6318e-127">脚本方法，其默认行为在Power Automate流</span><span class="sxs-lookup"><span data-stu-id="6318e-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="6318e-128">以下方法使用默认行为代替任何用户的当前状态。</span><span class="sxs-lookup"><span data-stu-id="6318e-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="6318e-129">类</span><span class="sxs-lookup"><span data-stu-id="6318e-129">Class</span></span> | <span data-ttu-id="6318e-130">方法</span><span class="sxs-lookup"><span data-stu-id="6318e-130">Method</span></span> | <span data-ttu-id="6318e-131">Power Automate行为</span><span class="sxs-lookup"><span data-stu-id="6318e-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="6318e-132">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="6318e-133">返回工作簿中的第一个工作表或该方法当前激活的 `Worksheet.activate` 工作表。</span><span class="sxs-lookup"><span data-stu-id="6318e-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="6318e-134">Worksheet</span><span class="sxs-lookup"><span data-stu-id="6318e-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="6318e-135">出于目的，将工作表标记为活动工作表 `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="6318e-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="data-refresh-not-supported-in-power-automate"></a><span data-ttu-id="6318e-136">数据刷新不受支持Power Automate</span><span class="sxs-lookup"><span data-stu-id="6318e-136">Data refresh not supported in Power Automate</span></span>

<span data-ttu-id="6318e-137">Office脚本在脚本中运行时无法刷新Power Automate。</span><span class="sxs-lookup"><span data-stu-id="6318e-137">Office Scripts can't refresh data when run in Power Automate.</span></span> <span data-ttu-id="6318e-138">在流 `PivotTable.refresh` 中调用此类方法时不执行任何操作。</span><span class="sxs-lookup"><span data-stu-id="6318e-138">Methods such as `PivotTable.refresh` do nothing when called in a flow.</span></span> <span data-ttu-id="6318e-139">此外，Power Automate不触发使用工作簿链接的公式的数据刷新。</span><span class="sxs-lookup"><span data-stu-id="6318e-139">Additionally, Power Automate doesn't trigger a data refresh for formulas that use workbook links.</span></span>

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a><span data-ttu-id="6318e-140">在流中运行时不执行任何操作的Power Automate方法</span><span class="sxs-lookup"><span data-stu-id="6318e-140">Script methods that do nothing when run in Power Automate flows</span></span>

<span data-ttu-id="6318e-141">通过脚本调用时，以下方法在脚本中Power Automate。</span><span class="sxs-lookup"><span data-stu-id="6318e-141">The following methods do nothing in a script when called through Power Automate.</span></span> <span data-ttu-id="6318e-142">它们仍然成功返回，并且不会引发任何错误。</span><span class="sxs-lookup"><span data-stu-id="6318e-142">They still return successfully and don't throw any errors.</span></span>

| <span data-ttu-id="6318e-143">类</span><span class="sxs-lookup"><span data-stu-id="6318e-143">Class</span></span> | <span data-ttu-id="6318e-144">方法</span><span class="sxs-lookup"><span data-stu-id="6318e-144">Method</span></span> |
|--|--|
| [<span data-ttu-id="6318e-145">PivotTable</span><span class="sxs-lookup"><span data-stu-id="6318e-145">PivotTable</span></span>](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [<span data-ttu-id="6318e-146">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-146">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [<span data-ttu-id="6318e-147">Workbook</span><span class="sxs-lookup"><span data-stu-id="6318e-147">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [<span data-ttu-id="6318e-148">Worksheet</span><span class="sxs-lookup"><span data-stu-id="6318e-148">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="6318e-149">使用文件浏览器控件选择工作簿</span><span class="sxs-lookup"><span data-stu-id="6318e-149">Select workbooks with the file browser control</span></span>

<span data-ttu-id="6318e-150">构建流 **中的"运行**"Power Automate步骤时，需要选择哪个工作簿是流的一部分。</span><span class="sxs-lookup"><span data-stu-id="6318e-150">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="6318e-151">使用文件浏览器选择工作簿，而不是手动键入工作簿的名称。</span><span class="sxs-lookup"><span data-stu-id="6318e-151">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="显示Power Automate文件浏览器选项的运行脚本操作。":::

<span data-ttu-id="6318e-153">有关工作簿动态Power Automate可能的解决方法的更多上下文，请参阅 Microsoft Power Automate Community 中的[此线程](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="6318e-153">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="6318e-154">时区差异</span><span class="sxs-lookup"><span data-stu-id="6318e-154">Time zone differences</span></span>

<span data-ttu-id="6318e-155">Excel文件没有固有位置或时区。</span><span class="sxs-lookup"><span data-stu-id="6318e-155">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="6318e-156">用户每次打开工作簿时，其会话都会使用该用户的本地时区进行日期计算。</span><span class="sxs-lookup"><span data-stu-id="6318e-156">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="6318e-157">Power Automate始终使用 UTC。</span><span class="sxs-lookup"><span data-stu-id="6318e-157">Power Automate always uses UTC.</span></span>

<span data-ttu-id="6318e-158">如果您的脚本使用日期或时间，则在本地测试脚本时与在脚本运行期间的行为Power Automate。</span><span class="sxs-lookup"><span data-stu-id="6318e-158">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="6318e-159">Power Automate允许你转换、设置格式和调整时间。</span><span class="sxs-lookup"><span data-stu-id="6318e-159">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="6318e-160">有关如何[在](https://flow.microsoft.com/blog/working-with-dates-and-times/)Power Automate 和[ `main` Parameters： Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)中使用这些函数的说明，请参阅在流内使用日期和时间，以了解如何为脚本提供该时间信息。</span><span class="sxs-lookup"><span data-stu-id="6318e-160">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="6318e-161">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6318e-161">See also</span></span>

- [<span data-ttu-id="6318e-162">脚本Office疑难解答</span><span class="sxs-lookup"><span data-stu-id="6318e-162">Troubleshoot Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="6318e-163">使用Office运行 Power Automate</span><span class="sxs-lookup"><span data-stu-id="6318e-163">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="6318e-164">ExcelOnline (Business) 连接器参考文档</span><span class="sxs-lookup"><span data-stu-id="6318e-164">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
