---
title: 将工作簿合并为单个工作簿
description: 了解如何使用脚本Office脚本Power Automate创建从其他工作簿合并到单个工作簿的工作表。
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 6d2c9492e0e2164fe34cff21d92f3df4c9bee3fe
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59337942"
---
# <a name="combine-worksheets-into-a-single-workbook"></a>将工作表合并到单个工作簿中

此示例演示如何将多个工作簿的数据提取到单个集中式工作簿中。 它使用两个脚本：一个从工作簿检索信息，另一个脚本使用该信息创建新的工作表。 它将脚本组合在一个Power Automate中，该流作用于整个 OneDrive 文件夹。

> [!IMPORTANT]
> 此示例仅复制其他工作簿中的值。 它不保留格式、图表、表格或其他对象。

## <a name="scenario"></a>应用场景

1. 在脚本Excel一个新的脚本OneDrive并添加此示例中的两个脚本。
1. 在文件夹中创建OneDrive并添加一个或多个包含数据的工作簿。
1. 构建一个流，获取该文件夹的所有文件。
1. 使用 **"返回工作表数据** "脚本从每个工作簿的每个工作表获取数据。
1. 使用 **"添加工作表"** 脚本在单个工作簿中为所有其他文件的每一个工作表创建新工作表。

## <a name="sample-code-return-worksheet-data"></a>示例代码：返回工作表数据

```TypeScript
/**
 * This script returns the values from the used ranges on each worksheet.
 */
function main(workbook: ExcelScript.Workbook): WorksheetData[]
{
  // Create an object to return the data from each worksheet.
  let worksheetInformation: WorksheetData[] = [];

  // Get the data from every worksheet, one at a time.
  workbook.getWorksheets().forEach((sheet) => {
    let values = sheet.getUsedRange()?.getValues();
    worksheetInformation.push({
       name: sheet.getName(),
       data: values as string[][]
    });
  });

  return worksheetInformation;
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="sample-code-add-worksheets"></a>示例代码：添加工作表

```TypeScript
/**
 * This script creates a new worksheet in the current workbook for each WorksheetData object provided.
 */
function main(workbook: ExcelScript.Workbook, workbookName: string, worksheetInformation: WorksheetData[])
{
  // Add each new worksheet.
  worksheetInformation.forEach((value) => {
    let sheet = workbook.addWorksheet(`${workbookName}.${value.name}`);

    // If there was any data in the worksheet, add it to a new range.
    if (value.data) {
      let range = sheet.getRangeByIndexes(0, 0, value.data.length, value.data[0].length);
      range.setValues(value.data);
    }
  });
}

// An interface to pass the worksheet name and cell values through a flow.
interface WorksheetData {
  name: string;
  data: string[][];
}
```

## <a name="power-automate-flow-combine-worksheets-into-a-single-workbook"></a>Power Automate流：将工作表合并到单个工作簿中

1. 登录到 [Power Automate](https://flow.microsoft.com)并创建新的 **即时云流**。
1. 选择 **"手动触发流"，** 然后选择"创建 **"。**
1. 获取文件夹中的所有文件。 本示例中，我们将使用名为"output"的文件夹。 添加一 **个新** 步骤，该步骤使用 **OneDrive for Business** 连接器和 **"在文件夹操作中列出文件**"。 提供包含文件的文件夹.csv路径。
    * **文件夹**： /output

    :::image type="content" source="../../images/combine-worksheets-flow-1.png" alt-text="已完成的OneDrive for Business连接器Power Automate。":::
1. 运行 **"返回工作表数据** "脚本，从每个工作簿获取所有数据。 使用 **"Excel脚本 (添加)** Online) **Business 连接器**。 对操作使用以下值。 请注意，添加文件的 *ID* 时，Power Automate操作包装在"应用到每个控件"中，因此该操作将在每个文件上执行。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件***：id* (文件夹中 **列表文件的动态**) 
    * **脚本**：返回工作表数据
1. 对 **新建的文件运行**"添加Excel脚本。 这将添加所有其他工作簿的数据。 执行上一 **个 Run 脚本** 操作后，在"应用到每个控件"内，使用"运行脚本"Excel (**Business)** 连接器 **。** 对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：你的文件
    * **脚本**：添加工作表
    * **workbookName：***从* (文件夹中 **的列表文件命名动态**) 
    * **worksheetInformation** *：run* script (中的结果 **)**

    :::image type="content" source="../../images/combine-worksheets-flow-2.png" alt-text="Apply to each 控件内的两个 Run 脚本操作。":::
    > [!NOTE]
    > 选择 **"切换到输入整个数组** "按钮以直接添加数组对象，而不是数组的单个项目。
    >
    > :::image type="content" source="../../images/combine-worksheets-flow-3.png" alt-text="用于切换为在控件字段输入框中输入整个数组的按钮。":::
1. 保存流。 使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。
1. 现在Excel文件应具有新工作表。
