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
# <a name="manage-calculation-mode-in-excel"></a>在计算模式下管理Excel

此示例演示如何在脚本中使用计算[模式](/javascript/api/office-scripts/excelscript/excelscript.calculationmode)和Excel web 版Office方法。 您可以尝试对任意文件Excel脚本。

## <a name="scenario"></a>应用场景

在Excel web 版中，可以使用 API 以编程方式控制文件的计算模式。 使用脚本可以执行Office操作。

1. 获取计算模式。
1. 设置计算模式。
1. 计算Excel设置为手动模式的文件的公式 (也称为重新计算) 。

## <a name="sample-code-control-calculation-mode"></a>示例代码：控制计算模式

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

## <a name="training-video-manage-calculation-mode"></a>培训视频：管理计算模式

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/iw6O8QH01CI)。
