---
title: 在计算模式下管理Excel
description: 了解如何在 Office 脚本中管理计算Excel web 版。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: 34a14874197ffda8487df5e450e3dcab980f7ed5
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232450"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="24b05-103">在计算模式下管理Excel</span><span class="sxs-lookup"><span data-stu-id="24b05-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="24b05-104">此示例演示如何在脚本中使用计算[模式](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)和Excel web 版Office方法。</span><span class="sxs-lookup"><span data-stu-id="24b05-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="24b05-105">您可以尝试对任意文件Excel脚本。</span><span class="sxs-lookup"><span data-stu-id="24b05-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="24b05-106">应用场景</span><span class="sxs-lookup"><span data-stu-id="24b05-106">Scenario</span></span>

<span data-ttu-id="24b05-107">在Excel web 版中，可以使用 API 以编程方式控制文件的计算模式。</span><span class="sxs-lookup"><span data-stu-id="24b05-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="24b05-108">使用脚本可以执行Office操作。</span><span class="sxs-lookup"><span data-stu-id="24b05-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="24b05-109">获取计算模式。</span><span class="sxs-lookup"><span data-stu-id="24b05-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="24b05-110">设置计算模式。</span><span class="sxs-lookup"><span data-stu-id="24b05-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="24b05-111">计算Excel设置为手动模式的文件的公式 (也称为重新计算) 。</span><span class="sxs-lookup"><span data-stu-id="24b05-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="24b05-112">示例代码：控制计算模式</span><span class="sxs-lookup"><span data-stu-id="24b05-112">Sample code: Control calculation mode</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook) {
    // Set calculation mode.
    workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.manual);
    // Get calculation mode.
    const calcMode = workbook.getApplication().getCalculationMode();    
    console.log(calcMode);
    // Calculate (for manual mode files).
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
}
```

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="24b05-113">培训视频：管理计算模式</span><span class="sxs-lookup"><span data-stu-id="24b05-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="24b05-114">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/iw6O8QH01CI)。</span><span class="sxs-lookup"><span data-stu-id="24b05-114">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/iw6O8QH01CI).</span></span>
