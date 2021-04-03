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
# <a name="manage-calculation-mode-in-excel"></a>在 Excel 中管理计算模式

此示例演示如何使用 Office 脚本 [在](/javascript/api/office-scripts/excelscript/excelscript.calculationmode) Excel 网页中使用计算模式和计算方法。 您可以尝试在任何 Excel 文件上的脚本。

## <a name="scenario"></a>方案

在 Excel 网页 Excel 中，可以使用 API 以编程方式控制文件的计算模式。 使用 Office 脚本可以执行下列操作。

1. 获取计算模式。
1. 设置计算模式。
1. 计算设置为手动模式的文件的 Excel 公式 (也称为重新计算) 。

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

[![观看有关如何在 Excel 网页中管理计算模式的分步视频](../../images/calc-mode-vid.jpg)](https://youtu.be/iw6O8QH01CI "有关如何在 Excel 网页中管理计算模式的分步视频")
