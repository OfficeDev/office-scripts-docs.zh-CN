---
title: 在 Excel 中管理计算模式
description: 了解如何使用 Office 脚本管理 Excel 网页中的计算模式。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: 0239437c7b52dca1fd8d1a4fc66bab7965cbd91a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571282"
---
# <a name="manage-calculation-mode-in-excel"></a><span data-ttu-id="eeabb-103">在 Excel 中管理计算模式</span><span class="sxs-lookup"><span data-stu-id="eeabb-103">Manage calculation mode in Excel</span></span>

<span data-ttu-id="eeabb-104">此示例演示如何使用 Office 脚本 [在](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) Excel 网页中使用计算模式和计算方法。</span><span class="sxs-lookup"><span data-stu-id="eeabb-104">This sample shows how to use the [calculation mode](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) and calculate methods in Excel on the web using Office Scripts.</span></span> <span data-ttu-id="eeabb-105">您可以尝试在任何 Excel 文件上的脚本。</span><span class="sxs-lookup"><span data-stu-id="eeabb-105">You can try the script on any Excel file.</span></span>

## <a name="scenario"></a><span data-ttu-id="eeabb-106">方案</span><span class="sxs-lookup"><span data-stu-id="eeabb-106">Scenario</span></span>

<span data-ttu-id="eeabb-107">在 Excel 网页 Excel 中，可以使用 API 以编程方式控制文件的计算模式。</span><span class="sxs-lookup"><span data-stu-id="eeabb-107">In Excel on the web, a file's calculation mode can be controlled programmatically using APIs.</span></span> <span data-ttu-id="eeabb-108">使用 Office 脚本可以执行下列操作。</span><span class="sxs-lookup"><span data-stu-id="eeabb-108">The following actions are possible using Office Scripts.</span></span>

1. <span data-ttu-id="eeabb-109">获取计算模式。</span><span class="sxs-lookup"><span data-stu-id="eeabb-109">Get the calculation mode.</span></span>
1. <span data-ttu-id="eeabb-110">设置计算模式。</span><span class="sxs-lookup"><span data-stu-id="eeabb-110">Set the calculation mode.</span></span>
1. <span data-ttu-id="eeabb-111">计算设置为手动模式的文件的 Excel 公式 (也称为重新计算) 。</span><span class="sxs-lookup"><span data-stu-id="eeabb-111">Calculate Excel formulas for files that are set to the manual mode (also referred to as recalculate).</span></span>

## <a name="sample-code-control-calculation-mode"></a><span data-ttu-id="eeabb-112">示例代码：控制计算模式</span><span class="sxs-lookup"><span data-stu-id="eeabb-112">Sample code: Control calculation mode</span></span>

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

## <a name="training-video-manage-calculation-mode"></a><span data-ttu-id="eeabb-113">培训视频：管理计算模式</span><span class="sxs-lookup"><span data-stu-id="eeabb-113">Training video: Manage calculation mode</span></span>

<span data-ttu-id="eeabb-114">[![观看有关如何在 Excel 网页中管理计算模式的分步视频](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "有关如何在 Excel 网页中管理计算模式的分步视频")</span><span class="sxs-lookup"><span data-stu-id="eeabb-114">[![Watch step-by-step video on how to manage calculation mode in Excel on the web](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "Step-by-step video on how to manage calculation mode in Excel on the web")</span></span>
