---
title: 在 Excel 网页版中使用 Office 脚本读取工作簿数据
description: 有关从工作簿中读取数据并评估脚本中的数据的 Office 脚本教程。
ms.date: 01/06/2021
localization_priority: Priority
ms.openlocfilehash: 0848a24e7333842b5b3b1f82ec8f270514c34d2f
ms.sourcegitcommit: 9df67e007ddbfec79a7360df9f4ea5ac6c86fb08
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/06/2021
ms.locfileid: "49772966"
---
# <a name="read-workbook-data-with-office-scripts-in-excel-on-the-web"></a>在 Excel 网页版中使用 Office 脚本读取工作簿数据

本教程将介绍如何在 Excel 网页版中使用 Office 脚本从工作簿中读取数据。 你将编写一个新脚本，该脚本可设置银行对帐单的格式并规范化该对帐单中的数据。 在此数据清理过程中，你的脚本将从事务单元格中读取值，将一个简单的公式应用到每个值，并将生成的答案写入工作簿。 通过从工作簿中读取数据，可在脚本中自动执行某些决策过程。

> [!TIP]
> 如果你不熟悉 Office 脚本，建议先查看[在 Excel 网页版中录制、编辑和创建 Office 脚本](excel-tutorial.md)教程。 [Office 脚本使用 TypeScript](../overview/code-editor-environment.md)，本教程面向在 JavaScript 或 TypeScript 方面具备初级到中级知识的人员。 如果你不熟悉 JavaScript，建议从 [Mozilla JavaScript 教程](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Introduction)入手。

## <a name="prerequisites"></a>先决条件

[!INCLUDE [Tutorial prerequisites](../includes/tutorial-prerequisites.md)]

## <a name="read-a-cell"></a>读取单元格

使用操作录制器创建的脚本只能将信息写入工作簿。 借助代码编辑器，可以编辑并创建也从工作簿中读取数据的脚本。

我们来创建一个读取数据并根据读取的数据执行操作的脚本。 我们将使用示例银行帐单。 此帐单是结合了支票和信贷的帐单。 遗憾的是，它们会以不同的方式报告余额变化。 支票帐单将收入作为正面信贷，将费用作为负面借记。 信贷帐单与之相反。

在本教程的其余部分中，我们将使用脚本对此数据进行标准化。 首先，让我们来了解如何从工作簿中读取数据。

1. 在用于教程其余部分的工作簿中创建新工作表。
2. 复制以下数据，并将其粘贴到新工作表中，从单元格 **A1** 开始。

    |日期 |帐户 |说明 |借记 |信贷 |
    |:--|:--|:--|:--|:--|
    |2019 年 10 月 10 日 |支票 |Coho Vineyard |-20.05 | |
    |2019 年 10 月 11 日 |信贷 |The Phone Company |99.95 | |
    |2019 年 10 月 13 日 |信贷 |Coho Vineyard |154.43 | |
    |2019 年 10 月 15 日 |支票 |外部存款 | |1000 |
    |2019 年 10 月 20 日 |信贷 |Coho Vineyard - 退款 | |-35.45 |
    |2019 年 10 月 25 日 |支票 |Best For You Organics Company | -85.64 | |
    |2019 年 11 月 1 日 |支票 |外部存款 | |1000 |

3. 打开“**所有脚本**”，然后选择“**新脚本**”。
4. 让我们来清理格式。 这是一个财务文档，因此更改“借记”和“信贷”列中的数字格式以将值显示为美元金额。 我们还调整列宽以适应数据。

    将脚本内容替换为以下代码：

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

5. 现在，让我们从数字列之一中读取一个值。 将以下代码添加到脚本的末尾（在结束 `}` 之前）：

    ```TypeScript
    // Get the value of cell D2.
    let range = selectedSheet.getRange("D2");
    console.log(range.getValues());
    ```

6. 运行脚本。
7. 应在控制台中看到 `[Array[1]]`。 这不是数字，因为区域是数据的二维数组。 该二维区域直接记录到控制台。 幸运的是，代码编辑器让你能够看到数组的内容。
8. 将二维数组记录到控制台时，它会对每行下面的列值进行分组。 按蓝色三角形展开数组日志。
9. 按新出现的蓝色三角形展开数组的第二级别。 现在，你应该会看到：

    ![控制台日志显示嵌套在两个数组下的输出“-20.05”](../images/tutorial-4.png)

## <a name="modify-the-value-of-a-cell"></a>修改单元格的值

现在，我们可以读取数据，让我们使用该数据来修改工作簿。 使单元格 **D2** 的值与 `Math.abs` 函数呈正相关。 [Math](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/math) 对象包含许多脚本具有访问权限的函数。 可在[使用 Office 脚本中的内置 JavaScript 对象](../develop/javascript-objects.md)中找到有关 `Math` 和其他内置对象的详细信息。

1. 我们将使用 `getValue` 和 `setValue` 方法更改单元格的值。 这些方法适用于单个单元格。 处理多单元格区域时，需使用 `getValues` 和 `setValues`。 将以下代码添加到脚本末尾：

    ```TypeScript
    // Run the `Math.abs` function with the value at D2 and apply that value back to D2.
    let positiveValue = Math.abs(range.getValue() as number);
    range.setValue(positiveValue);
    ```

    > [!NOTE]
    > 我们正使用 `as` 关键字将 `range.getValue()` 的返回值 [转换](https://www.typescripttutorial.net/typescript-tutorial/type-casting/) 为 `number`。  这样做很有必要，因为区域可能是字符串、数字或布尔值。 在本实例中，我们明确需要数字。

2. 单元格 **D2** 的值现在应为正值。

## <a name="modify-the-values-of-a-column"></a>修改列的值

现在，我们知道如何读取和写入单个单元格，让我们对脚本进行一般化，使其适用于整个“借记”和“信贷”列。

1. 删除仅影响单个单元格的代码（先前的绝对值代码），以便你的脚本现在如下所示：

    ```TypeScript
    function main(workbook: ExcelScript.Workbook) {
        // Get the current worksheet.
        let selectedSheet = workbook.getActiveWorksheet();

        // Format the range to display numerical dollar amounts.
        selectedSheet.getRange("D2:E8").setNumberFormat("$#,##0.00");

        // Fit the width of all the used columns to the data.
        selectedSheet.getUsedRange().getFormat().autofitColumns();
    }
    ```

2. 在脚本末尾添加循环访问最后两列中的行的循环。 对于每个单元格，脚本将值设置为当前值的绝对值。

    请注意，定义单元格位置的数组是从零开始的。 这意味着单元格 **A1** 为 `range[0][0]`。

    ```TypeScript
    // Get the values of the used range.
    let range = selectedSheet.getUsedRange();
    let rangeValues = range.getValues();

    // Iterate over the fourth and fifth columns and set their values to their absolute value.
    let rowCount = range.getRowCount();
    for (let i = 1; i < rowCount; i++) {
        // The column at index 3 is column "4" in the worksheet.
        if (rangeValues[i][3] != 0) {
            let positiveValue = Math.abs(rangeValues[i][3] as number);
            selectedSheet.getCell(i, 3).setValue(positiveValue);
        }

        // The column at index 4 is column "5" in the worksheet.
        if (rangeValues[i][4] != 0) {
            let positiveValue = Math.abs(rangeValues[i][4] as number);
            selectedSheet.getCell(i, 4).setValue(positiveValue);
        }
    }
    ```

    此部分的脚本执行几项重要任务。 首先，获取已用区域的值和行计数。 这样，我们就可以查看值并知道何时停止。 其次，循环访问已用区域，检查“借记”或“信贷”列中的每个单元格。 最后，如果单元格中的值不为 0，则该值将替换为其绝对值。 我们正在避免使用零，因此可以将空白单元格保留原样。

3. 运行脚本。

    现在，你的银行帐单如下所示：

    ![银行帐单作为仅包含正值的格式表](../images/tutorial-5.png)

## <a name="next-steps"></a>后续步骤

打开“代码编辑器”，然后尝试使用一些 [Excel 网页版中的 Office 脚本的示例脚本](../resources/excel-samples.md)。 还可以访问 [Excel 网页版中的 Office 脚本的脚本基础知识](../develop/scripting-fundamentals.md)，了解有关创建 Office 脚本的详细信息。

下一系列的 Office 脚本教程重点介绍如何将 Office 脚本与 Power Automate 一起使用。 在[使用 Power Automate 运行 Office 脚本](../develop/power-automate-integration.md)中了解有关结合两个平台的优势的更多信息，或尝试[通过 Power Automate 手动流调用脚本](excel-power-automate-manual.md)教程来创建使用 Office 脚本的 Power Automate 流。
