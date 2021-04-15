---
title: Power Automate with Office Scripts 疑难解答信息
description: 有关 Office 脚本和 Power Automate 之间集成的提示、平台信息和已知问题。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: 59f4cd8b3476c2ee2a1a862f136173a543ba8a15
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51755005"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a><span data-ttu-id="c6887-103">Power Automate with Office Scripts 疑难解答信息</span><span class="sxs-lookup"><span data-stu-id="c6887-103">Troubleshooting information for Power Automate with Office Scripts</span></span>

<span data-ttu-id="c6887-104">借助 Power Automate，你可以将 Office 脚本自动化上一个级别。</span><span class="sxs-lookup"><span data-stu-id="c6887-104">Power Automate lets you take your Office Script automation to the next level.</span></span> <span data-ttu-id="c6887-105">但是，由于 Power Automate 在独立的 Excel 会话中代表您运行脚本，因此有一些重要的注意事项。</span><span class="sxs-lookup"><span data-stu-id="c6887-105">However, because Power Automate runs scripts on your behalf in independent Excel sessions, there are a few important things to note.</span></span>

> [!TIP]
> <span data-ttu-id="c6887-106">如果刚开始将 Office 脚本与 Power Automate 一同使用，请从使用 [Power Automate 运行 Office 脚本](../develop/power-automate-integration.md) 开始了解平台。</span><span class="sxs-lookup"><span data-stu-id="c6887-106">If you're just starting to use Office Scripts with Power Automate, please start with [Run Office Scripts with Power Automate](../develop/power-automate-integration.md) to learn about the platforms.</span></span>

## <a name="avoid-using-relative-references"></a><span data-ttu-id="c6887-107">避免使用相对引用</span><span class="sxs-lookup"><span data-stu-id="c6887-107">Avoid using relative references</span></span>

<span data-ttu-id="c6887-108">Power Automate 代表你运行所选 Excel 工作簿中的脚本。</span><span class="sxs-lookup"><span data-stu-id="c6887-108">Power Automate runs your script in the chosen Excel workbook on your behalf.</span></span> <span data-ttu-id="c6887-109">发生这种情况时，工作簿可能会关闭。</span><span class="sxs-lookup"><span data-stu-id="c6887-109">The workbook might be closed when this happens.</span></span> <span data-ttu-id="c6887-110">依赖于用户当前状态的任何 API（如 ）在 Power Automate 中的行为 `Workbook.getActiveWorksheet` 可能有所不同。</span><span class="sxs-lookup"><span data-stu-id="c6887-110">Any API that relies on the user's current state, such as `Workbook.getActiveWorksheet`, may behave differently in Power Automate.</span></span> <span data-ttu-id="c6887-111">这是因为 API 基于用户视图或游标的相对位置，并且 Power Automate 流中不存在该引用。</span><span class="sxs-lookup"><span data-stu-id="c6887-111">This is because the APIs are based on a relative position of the user's view or cursor and that reference doesn't exist in a Power Automate flow.</span></span>

<span data-ttu-id="c6887-112">某些相对引用 API 在 Power Automate 中引发错误。</span><span class="sxs-lookup"><span data-stu-id="c6887-112">Some relative reference APIs throw errors in Power Automate.</span></span> <span data-ttu-id="c6887-113">其他人有一个默认行为，表示用户的状态。</span><span class="sxs-lookup"><span data-stu-id="c6887-113">Others have a default behavior that implies a user's state.</span></span> <span data-ttu-id="c6887-114">在设计脚本时，请确保对工作表和范围使用绝对引用。</span><span class="sxs-lookup"><span data-stu-id="c6887-114">When designing your scripts, be sure to use absolute references for worksheets and ranges.</span></span> <span data-ttu-id="c6887-115">这使 Power Automate 流程保持一致，即使工作表已重新排列。</span><span class="sxs-lookup"><span data-stu-id="c6887-115">This makes your Power Automate flow consistent, even if worksheets are rearranged.</span></span>

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a><span data-ttu-id="c6887-116">运行 Power Automate 流时失败的脚本方法</span><span class="sxs-lookup"><span data-stu-id="c6887-116">Script methods that fail when run Power Automate flows</span></span>

<span data-ttu-id="c6887-117">从 Power Automate 流中的脚本调用时，以下方法将引发错误并失败。</span><span class="sxs-lookup"><span data-stu-id="c6887-117">The following methods will throw an error and fail when called from a script in a Power Automate flow.</span></span>

| <span data-ttu-id="c6887-118">类</span><span class="sxs-lookup"><span data-stu-id="c6887-118">Class</span></span> | <span data-ttu-id="c6887-119">Method</span><span class="sxs-lookup"><span data-stu-id="c6887-119">Method</span></span> |
|--|--|
| [<span data-ttu-id="c6887-120">Chart</span><span class="sxs-lookup"><span data-stu-id="c6887-120">Chart</span></span>](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [<span data-ttu-id="c6887-121">Range</span><span class="sxs-lookup"><span data-stu-id="c6887-121">Range</span></span>](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [<span data-ttu-id="c6887-122">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-122">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [<span data-ttu-id="c6887-123">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-123">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [<span data-ttu-id="c6887-124">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-124">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [<span data-ttu-id="c6887-125">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-125">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [<span data-ttu-id="c6887-126">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-126">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a><span data-ttu-id="c6887-127">Power Automate 流中具有默认行为的脚本方法</span><span class="sxs-lookup"><span data-stu-id="c6887-127">Script methods with a default behavior in Power Automate flows</span></span>

<span data-ttu-id="c6887-128">以下方法使用默认行为代替任何用户的当前状态。</span><span class="sxs-lookup"><span data-stu-id="c6887-128">The following methods use a default behavior, in lieu of any user's current state.</span></span>

| <span data-ttu-id="c6887-129">类</span><span class="sxs-lookup"><span data-stu-id="c6887-129">Class</span></span> | <span data-ttu-id="c6887-130">Method</span><span class="sxs-lookup"><span data-stu-id="c6887-130">Method</span></span> | <span data-ttu-id="c6887-131">Power Automate 行为</span><span class="sxs-lookup"><span data-stu-id="c6887-131">Power Automate behavior</span></span> |
|--|--|--|
| [<span data-ttu-id="c6887-132">Workbook</span><span class="sxs-lookup"><span data-stu-id="c6887-132">Workbook</span></span>](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | <span data-ttu-id="c6887-133">返回工作簿中的第一个工作表或该方法当前激活的 `Worksheet.activate` 工作表。</span><span class="sxs-lookup"><span data-stu-id="c6887-133">Returns either the first worksheet in the workbook or the worksheet currently activated by the `Worksheet.activate` method.</span></span> |
| [<span data-ttu-id="c6887-134">Worksheet</span><span class="sxs-lookup"><span data-stu-id="c6887-134">Worksheet</span></span>](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | <span data-ttu-id="c6887-135">出于目的，将工作表标记为活动工作表 `Workbook.getActiveWorksheet` 。</span><span class="sxs-lookup"><span data-stu-id="c6887-135">Marks the worksheet as the active worksheet for purposes of `Workbook.getActiveWorksheet`.</span></span> |

## <a name="select-workbooks-with-the-file-browser-control"></a><span data-ttu-id="c6887-136">使用文件浏览器控件选择工作簿</span><span class="sxs-lookup"><span data-stu-id="c6887-136">Select workbooks with the file browser control</span></span>

<span data-ttu-id="c6887-137">生成 Power **Automate 流的 Run 脚本** 步骤时，需要选择哪个工作簿是流的一部分。</span><span class="sxs-lookup"><span data-stu-id="c6887-137">When building the **Run script** step of a Power Automate flow, you need to select which workbook is part of the flow.</span></span> <span data-ttu-id="c6887-138">使用文件浏览器选择工作簿，而不是手动键入工作簿的名称。</span><span class="sxs-lookup"><span data-stu-id="c6887-138">Use the file browser to select your workbook, instead of manually typing the workbook's name.</span></span>

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="显示显示选取器文件浏览器选项的 Power Automate Run 脚本操作。":::

<span data-ttu-id="c6887-140">有关 Power Automate 限制的更多上下文和有关动态选择工作簿的潜在解决方法的讨论，请参阅 Microsoft Power Automate Community 中的 [此线程](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。</span><span class="sxs-lookup"><span data-stu-id="c6887-140">For more context on the Power Automate limitation and a discussion of potential workarounds for the dynamic selection of workbooks, see [this thread in the Microsoft Power Automate Community](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#).</span></span>

## <a name="time-zone-differences"></a><span data-ttu-id="c6887-141">时区差异</span><span class="sxs-lookup"><span data-stu-id="c6887-141">Time zone differences</span></span>

<span data-ttu-id="c6887-142">Excel 文件没有固有位置或时区。</span><span class="sxs-lookup"><span data-stu-id="c6887-142">Excel files don't have an inherent location or timezone.</span></span> <span data-ttu-id="c6887-143">用户每次打开工作簿时，其会话都会使用该用户的本地时区进行日期计算。</span><span class="sxs-lookup"><span data-stu-id="c6887-143">Every time a user opens the workbook, their session uses that user's local timezone for date calculations.</span></span> <span data-ttu-id="c6887-144">Power Automate 始终使用 UTC。</span><span class="sxs-lookup"><span data-stu-id="c6887-144">Power Automate always uses UTC.</span></span>

<span data-ttu-id="c6887-145">如果脚本使用日期或时间，则在本地测试脚本时与通过 Power Automate 运行脚本时可能有行为差异。</span><span class="sxs-lookup"><span data-stu-id="c6887-145">If your script uses dates or times, there may be behavioral differences when the script is tested locally versus when it is run through Power Automate.</span></span> <span data-ttu-id="c6887-146">Power Automate 允许你转换、格式化和调整时间。</span><span class="sxs-lookup"><span data-stu-id="c6887-146">Power Automate allows you to convert, format, and adjust times.</span></span> <span data-ttu-id="c6887-147">有关如何[在](https://flow.microsoft.com/blog/working-with-dates-and-times/)Power Automate 和[ `main` Parameters： Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script)中使用这些函数的说明，请参阅在流内使用日期和时间，以了解如何为脚本提供该时间信息。</span><span class="sxs-lookup"><span data-stu-id="c6887-147">See [Working with Dates and Times inside of your flows](https://flow.microsoft.com/blog/working-with-dates-and-times/) for instructions on how to use those functions in Power Automate and [`main` Parameters: Passing data to a script](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script) to learn how to provide that time information for the script.</span></span>

## <a name="see-also"></a><span data-ttu-id="c6887-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="c6887-148">See also</span></span>

- [<span data-ttu-id="c6887-149">Office 脚本疑难解答</span><span class="sxs-lookup"><span data-stu-id="c6887-149">Troubleshooting Office Scripts</span></span>](troubleshooting.md)
- [<span data-ttu-id="c6887-150">使用 Power Automate 运行 Office 脚本</span><span class="sxs-lookup"><span data-stu-id="c6887-150">Run Office Scripts with Power Automate</span></span>](../develop/power-automate-integration.md)
- [<span data-ttu-id="c6887-151">Excel Online (Business) 连接器参考文档</span><span class="sxs-lookup"><span data-stu-id="c6887-151">Excel Online (Business) connector reference documentation</span></span>](/connectors/excelonlinebusiness/)
