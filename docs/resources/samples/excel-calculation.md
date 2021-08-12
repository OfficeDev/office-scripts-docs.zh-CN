---
title: 在计算模式下管理Excel
description: 了解如何在 Office 脚本中管理计算Excel web 版。
ms.date: 05/06/2021
localization_priority: Normal
ms.openlocfilehash: d33c4f21b21333ccefe26effc3df70235978b480a999364793e9a45d21dfba7f
ms.sourcegitcommit: 75f7ed8c2d23a104acc293f8ce29ea580b4fcdc5
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/11/2021
ms.locfileid: "57846706"
---
# <a name="manage-calculation-mode-in-excel"></a>在计算模式下管理Excel

此示例演示如何在脚本中使用计算[模式](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)和计算Excel web 版脚本Office方法。 您可以尝试对任意文件Excel脚本。

## <a name="scenario"></a>方案

包含大量公式的工作簿可能需要一段时间才能重新计算。 与其让Excel控制何时进行计算，不如将其作为脚本的一部分进行管理。 这将在某些情况下帮助提高性能。

示例脚本将计算模式设置为手动。 这意味着，当脚本指示工作簿执行自定义操作或通过 UI (手动计算时，工作簿将仅重新计算[) 。](https://support.microsoft.com/office/change-formula-recalculation-iteration-or-precision-in-excel-73fc7dac-91cf-4d36-86e8-67124f6bcce4) 然后，该脚本显示当前计算模式并完全重新计算整个工作簿。

## <a name="sample-code-control-calculation-mode"></a>示例代码：控制计算模式

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

## <a name="training-video-manage-calculation-mode"></a>培训视频：管理计算模式

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/iw6O8QH01CI)。
