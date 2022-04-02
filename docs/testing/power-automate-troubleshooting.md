---
title: 对Office中运行的脚本进行Power Automate
description: 使用技巧脚本和脚本之间的集成时，Office、平台信息和Power Automate。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2c256c2ddc64fcfc510f24e27662234f44b65ac0
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64586029"
---
# <a name="troubleshoot-office-scripts-running-in-power-automate"></a>对Office中运行的脚本进行Power Automate

Power Automate使脚本自动化Office一个级别。 但是，Power Automate在独立会话中代表您Excel脚本，因此有一些重要的注意事项。

> [!TIP]
> 如果你刚刚开始将 Office 脚本与 Power Automate 一起Power Automate[运行 Office 脚本](../develop/power-automate-integration.md)，了解这些平台。

## <a name="avoid-relative-references"></a>避免相对引用

Power Automate代表您Excel所选工作簿中运行脚本。 发生这种情况时，工作簿可能会关闭。 任何依赖用户`Workbook.getActiveWorksheet`当前状态（如 ）的 API 在用户Power Automate。 这是因为 API 基于用户视图或游标的相对位置，并且该引用在用户流中不存在Power Automate。

某些相对引用 API 在Power Automate。 其他人有一个默认行为，表示用户的状态。 在设计脚本时，请确保对工作表和范围使用绝对引用。 这样即使重新Power Automate工作表，也使数据流保持一致。

### <a name="script-methods-that-fail-when-run-in-power-automate-flows"></a>在流中运行时失败的Power Automate方法

以下方法引发错误，在从脚本流中的脚本调用时Power Automate失败。

| 类 | 方法 |
|--|--|
| [Chart](/javascript/api/office-scripts/excelscript/excelscript.chart) | `activate` |
| [Range](/javascript/api/office-scripts/excelscript/excelscript.range) | `select` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveCell` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveChart` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveSlicer` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRange` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getSelectedRanges` |

### <a name="script-methods-with-a-default-behavior-in-power-automate-flows"></a>脚本方法，其默认行为在Power Automate流

以下方法使用默认行为代替任何用户的当前状态。

| 类 | 方法 | Power Automate行为 |
|--|--|--|
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `getActiveWorksheet` | 返回工作簿中的第一个工作表或该方法当前激活的 `Worksheet.activate` 工作表。 |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `activate` | 出于目的，将工作表标记为活动工作表 `Workbook.getActiveWorksheet`。 |

## <a name="data-refresh-not-supported-in-power-automate"></a>数据刷新不受支持Power Automate

Office脚本在脚本中运行时无法刷新Power Automate。 在流中 `PivotTable.refresh` 调用此类方法时不执行任何操作。 此外，Power Automate不触发使用工作簿链接的公式的数据刷新。

### <a name="script-methods-that-do-nothing-when-run-in-power-automate-flows"></a>在流中运行时不执行任何操作的Power Automate方法

通过脚本调用时，以下方法在脚本中Power Automate。 它们仍然成功返回，并且不会引发任何错误。

| 类 | 方法 |
|--|--|
| [PivotTable](/javascript/api/office-scripts/excelscript/excelscript.pivottable) | `refresh` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllDataConnections` |
| [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) | `refreshAllPivotTables` |
| [Worksheet](/javascript/api/office-scripts/excelscript/excelscript.worksheet) | `refreshAllPivotTables` |

## <a name="select-workbooks-with-the-file-browser-control"></a>使用文件浏览器控件选择工作簿

构建流 **中的"运行"** Power Automate步骤时，需要选择哪个工作簿是流的一部分。 使用文件浏览器选择工作簿，而不是手动键入工作簿的名称。

:::image type="content" source="../images/power-automate-file-browser.png" alt-text="显示Power Automate文件浏览器选项的&quot;运行脚本&quot;操作。":::

有关工作簿动态Power Automate可能的解决方法的更多上下文，请参阅 [Microsoft](https://powerusers.microsoft.com/t5/Power-Automate-Ideas/Allow-for-dynamic-quot-file-quot-value-for-excel-quot-get-a-row/idi-p/103091#) Power Automate Community。

## <a name="pass-entire-arrays-as-script-parameters"></a>将整个数组作为脚本参数传递

Power Automate允许用户将数组作为变量或数组中的单个元素传递给连接器。 默认值是传递单个元素，这将在流中生成数组。 对于将整个数组作为参数的脚本或其他连接器，需要选择"切换到 **输入** 整个数组"按钮以将数组作为一个完整的对象传递。 此按钮位于每个数组参数输入字段的右上角。

:::image type="content" source="../images/combine-worksheets-flow-3.png" alt-text="用于切换为在控件字段输入框中输入整个数组的按钮。":::

## <a name="time-zone-differences"></a>时区差异

Excel文件没有固有位置或时区。 用户每次打开工作簿时，其会话都会使用该用户的本地时区进行日期计算。 Power Automate始终使用 UTC。

如果您的脚本使用日期或时间，则在本地测试脚本与在本地测试脚本时的行为Power Automate。 Power Automate允许你转换、设置格式和调整时间。 有关如何[在](https://flow.microsoft.com/blog/working-with-dates-and-times/) Power Automate [`main` 和 Parameters： Pass data to a script](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script) 中使用这些函数的说明，请参阅在流内使用日期和时间，以了解如何为脚本提供该时间信息。

## <a name="script-parameter-fields-or-returned-output-not-appearing-in-power-automate"></a>脚本参数字段或返回的输出未显示在Power Automate

脚本的参数或返回的数据未准确反映到流生成器中的两Power Automate原因。

- 脚本签名 (添加) **Business Excel Online** (连接器后) 更改的参数或返回值。
- 脚本签名使用不受支持的类型。 根据参数下的列表验证类型[，](../develop/power-automate-integration.md#main-parameters-pass-data-to-a-script)并返回 Run [Office Scripts with Power Automate](../develop/power-automate-integration.md)部分。[](../develop/power-automate-integration.md#return-data-from-a-script)

创建脚本时，脚本的签名与 **Excel Business (Online)** 连接器一起存储。 删除旧连接器并创建一个新连接器，获取最新的参数并返回 **Run 脚本操作** 的值。

## <a name="see-also"></a>另请参阅

- [脚本Office疑难解答](troubleshooting.md)
- [使用Office脚本运行Power Automate](../develop/power-automate-integration.md)
- [Excel Online (Business) 连接器参考文档](/connectors/excelonlinebusiness/)
