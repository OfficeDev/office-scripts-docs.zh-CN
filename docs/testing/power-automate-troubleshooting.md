---
title: Power Automate with Office Scripts 疑难解答信息
description: 有关 Office 脚本和 Power Automate 之间集成的提示、平台信息和已知问题。
ms.date: 01/14/2021
localization_priority: Normal
ms.openlocfilehash: b0f5b2f542216789f0d96f309cb7d799d201ba0f
ms.sourcegitcommit: e7e019ba36c2f49451ec08c71a1679eb6dba4268
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/22/2021
ms.locfileid: "49933264"
---
# <a name="troubleshooting-information-for-power-automate-with-office-scripts"></a>Power Automate with Office Scripts 疑难解答信息

借助 Power Automate，你可以将 Office 脚本自动化提高至下一级别。 但是，由于 Power Automate 在独立的 Excel 会话中代表您运行脚本，因此有一些重要的注意事项。

> [!TIP]
> 如果刚开始将 Office 脚本与 Power Automate 一同使用，请从运行 Office 脚本和 [Power Automate](../develop/power-automate-integration.md) 开始了解平台。

## <a name="avoid-using-relative-references"></a>避免使用相对引用

Power Automate 代表你运行所选 Excel 工作簿中的脚本。 发生这种情况时，工作簿可能会关闭。 依赖于用户当前状态的任何 API（如 Power `Workbook.getActiveWorksheet` Automate）的行为可能有所不同。 这是因为 API 基于用户视图或游标的相对位置，并且该引用在 Power Automate 流中不存在。

某些相对引用 API 在 Power Automate 中引发错误。 其他人有一个表示用户状态的默认行为。 在设计脚本时，请确保对工作表和范围使用绝对引用。 这使 Power Automate 流保持一致，即使工作表已重新排列。

### <a name="script-methods-that-fail-when-run-power-automate-flows"></a>运行 Power Automate 流时失败的脚本方法

从 Power Automate 流的脚本调用时，以下方法将引发错误并失败。

| 类 | Method |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>Power Automate 流中具有默认行为的脚本方法

以下方法使用默认行为代替任何用户的当前状态。

| 类 | Method | Power Automate 行为 |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | 返回工作簿的第一个工作表或该方法当前激活的 `Worksheet.activate` 工作表。 |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | 出于目的将工作表标记为活动工作表 `Workbook.getActiveWorksheet` 。 |

## <a name="select-workbooks-with-the-file-browser-control"></a>使用文件浏览器控件选择工作簿

生成 Power Automate 流的 **Run** 脚本步骤时，需要选择哪个工作簿是流的一部分。 使用文件浏览器选择工作簿，而不是手动键入工作簿的名称。

![在 Power Automate 中创建"运行脚本"操作时的文件浏览器选项](../images/power-automate-file-browser.png)

有关 Power Automate 限制的更多上下文和有关动态选择工作簿的潜在解决方法的讨论，请参阅 [Microsoft Power Automate 社区中的此线程](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#)。

## <a name="time-zone-differences"></a>时区差异

Excel 文件没有固有位置或时区。 用户每次打开工作簿时，其会话都会使用该用户的本地时区进行日期计算。 Power Automate 始终使用 UTC。

如果您的脚本使用日期或时间，则在本地测试脚本时与通过 Power Automate 运行脚本时可能有行为差异。 Power Automate 允许你转换、设置格式和调整时间。 有关如何[在](https://flow.microsoft.com/blog/working-with-dates-and-times/)Power Automate 和[ `main` Parameters](../develop/power-automate-integration.md#main-parameters-passing-data-to-a-script)中使用这些函数的说明，请参阅在流内使用日期和时间：将数据传递到脚本，了解如何为脚本提供该时间信息。

## <a name="see-also"></a>另请参阅

- [Office 脚本疑难解答](troubleshooting.md)
- [使用 Power Automate 运行 Office 脚本](../develop/power-automate-integration.md)
- [Excel Online (Business) 连接器参考文档](/connectors/excelonlinebusiness/)
