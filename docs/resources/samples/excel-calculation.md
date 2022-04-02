---
title: 在计算模式下管理Excel
description: 了解如何在 Office 脚本中管理计算Excel web 版。
ms.date: 05/06/2021
ms.localizationpriority: medium
ms.openlocfilehash: fec88c904d95bfdab1514d44921f7fb1c6e9dd35
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585511"
---
# <a name="manage-calculation-mode-in-excel"></a>在计算模式下管理Excel

此示例演示如何在脚本中使用计算[模式](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)Excel web 版计算Office方法。 您可以尝试对任意文件Excel脚本。

## <a name="scenario"></a>应用场景

包含大量公式的工作簿可能需要一段时间才能重新计算。 与其让Excel控制何时进行计算，不如将它们作为脚本的一部分进行管理。 这将在某些情况下帮助提高性能。

示例脚本将计算模式设置为手动。 这意味着，当脚本指示工作簿执行自定义操作或你通过 [UI (手动](https://support.microsoft.com/office/73fc7dac-91cf-4d36-86e8-67124f6bcce4) 计算时，工作簿将仅重新计算) 。 然后，该脚本显示当前计算模式并完全重新计算整个工作簿。

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
