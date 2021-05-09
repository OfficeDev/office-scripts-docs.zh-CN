---
title: 在计算模式下管理Excel
description: 了解如何在 Office 脚本中管理计算Excel web 版。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: a60fddc91b3a8f124a44722d0d75e6e9f239351d
ms.sourcegitcommit: 763d341857bcb209b2f2c278a82fdb63d0e18f0a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/08/2021
ms.locfileid: "52285911"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="336b0-103">在计算模式下管理Excel</span><span class="sxs-lookup"><span data-stu-id="336b0-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="336b0-104">此示例演示如何在脚本中使用计算[模式](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)和Excel web 版Office方法。</span><span class="sxs-lookup"><span data-stu-id="336b0-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="336b0-105">您可以尝试对任意文件Excel脚本。</span><span class="sxs-lookup"><span data-stu-id="336b0-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="336b0-106">应用场景</span><span class="sxs-lookup"><span data-stu-id="336b0-106">Scenario</span></span>

<span data-ttu-id="336b0-107">包含大量公式的工作簿可能需要一段时间才能重新计算。</span><span class="sxs-lookup"><span data-stu-id="336b0-107">Workbooks with large numbers of formulas can take a while to recalculate.</span></span> <span data-ttu-id="336b0-108">与其让Excel控制何时进行计算，不如将它们作为脚本的一部分进行管理。</span><span class="sxs-lookup"><span data-stu-id="336b0-108">Rather than letting Excel control when calculations happen, you can manage them as part of your script.</span></span> <span data-ttu-id="336b0-109">这将在某些情况下帮助提高性能。</span><span class="sxs-lookup"><span data-stu-id="336b0-109">This will help with performance in certain scenarios.</span></span>

<span data-ttu-id="336b0-110">示例脚本将计算模式设置为手动。</span><span class="sxs-lookup"><span data-stu-id="336b0-110">The sample script sets the calculation mode to manual.</span></span> <span data-ttu-id="336b0-111">这意味着，当脚本指示工作簿执行自定义操作或通过 UI (手动计算时，工作簿将仅[重新计算) 。](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)</span><span class="sxs-lookup"><span data-stu-id="336b0-111">This means that the workbook will only recalculate formulas when the script tells it to (or you [manually calculate through the UI](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4)).</span></span> <span data-ttu-id="336b0-112">然后，该脚本显示当前计算模式并完全重新计算整个工作簿。</span><span class="sxs-lookup"><span data-stu-id="336b0-112">The script then displays the current calculation mode and fully recalculates the entire workbook.</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="336b0-113">示例代码：控制计算模式</span><span class="sxs-lookup"><span data-stu-id="336b0-113">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set the calculation mode to manual.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get and log the calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Manually calculate the file.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="336b0-114">培训视频：管理计算模式</span><span class="sxs-lookup"><span data-stu-id="336b0-114">Training video: Manage calculation mode</span></span>

<span data-ttu-id="336b0-115">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/iw6O8QH01CI)。</span><span class="sxs-lookup"><span data-stu-id="336b0-115">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
