---
title: 将 CSV 文件转换为Excel工作簿
description: 了解如何使用脚本Office脚本Power Automate从.xlsx创建.csv文件。
ms.date: 02/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5e501368015840d4181c5565662638b65e213fed
ms.sourcegitcommit: 49f527a7f54aba00e843ad4a92385af59c1d7bfa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/08/2022
ms.locfileid: "63352123"
---
# <a name="convert-csv-files-to-excel-workbooks"></a>将 CSV 文件转换为Excel工作簿

许多服务将数据导出为 CSV 文件 (逗号) 值。 此解决方案可自动执行将 CSV 文件转换为Excel文件格式.xlsx工作簿的过程。 它使用 Power Automate [](https://flow.microsoft.com) 流在 OneDrive 文件夹中查找具有 .csv 扩展名的文件，并使用 Office 脚本将数据从 .csv 文件复制到新的 Excel 工作簿。

## <a name="solution"></a>解决方案

1. 将.csv文件以及空白的"模板".xlsx存储在一个OneDrive文件夹中。
1. 创建一Office脚本以将 CSV 数据解析为一个范围。
1. 创建Power Automate流以读取.csv文件，并传递其内容到脚本。

## <a name="sample-files"></a>示例文件

下载 <a href="https://github.com/OfficeDev/office-scripts-docs/blob/master/docs/resources/samples/convert-csv-example.zip?raw=true">convert-csv-example.zip</a> 文件，获取Template.xlsx和两个示例.csv文件。 将文件解压缩到文件文件夹中OneDrive。 本示例假定该文件夹名为"output"。

添加以下脚本，然后使用给定的步骤构建一个流，以自己尝试示例！

## <a name="sample-code-insert-comma-separated-values-into-a-workbook"></a>示例代码：在工作簿中插入逗号分隔的值

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);

    // Check for blanks at the start of the row.
    if (row[0].charAt(0) === ',') {
      row.unshift("");
    }
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });

  // Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);

  // Add any formatting or table creation that you want.
}
```

## <a name="power-automate-flow-create-new-xlsx-files"></a>Power Automate流：创建新的.xlsx文件

1. 登录 [Power Automate并](https://flow.microsoft.com)创建新的 **计划云流**。
1. 将流程设置为" **每** "1"天重复一次，然后选择"创建 **"**。
1. 获取模板Excel文件。 这是所有已转换的已转换.csv的基础。 添加一 **个使用** OneDrive for Business 连接器和"获取文件内容 **"****操作的新** 步骤。 提供指向"Template.xlsx"文件的文件路径。
    * **文件**：/output/Template.xlsx
1. 通过 **进入该** 步骤 (在连接器) 右上角的"获取文件内容 **(...)** 菜单并选择"重命名"选项，重命名"获取文件内容"步骤。 将步骤名称更改为"获取Excel模板"。

     :::image type="content" source="../../images/convert-csv-flow-1.png" alt-text="OneDrive for Business中Power Automate连接器，重命名为&quot;获取Excel模板。":::
1. 获取"output"文件夹中的所有文件。 添加一 **个新** 步骤，该步骤使用 **OneDrive for Business** 连接器和 **"在文件夹操作中列出文件**"。 提供包含文件文件夹.csv路径。
    * **文件夹**：/output

    :::image type="content" source="../../images/convert-csv-flow-2.png" alt-text="已完成的OneDrive for Business连接器Power Automate。":::
1. 添加一个条件，以便流仅对文件.csv运行。 添加 **一个作为** 条件控件 **的新** 步骤。 对 Condition 使用以下 **值**。
    * **选择一个值**： *从* (文件夹中的列表 **文件命名动态**) 。 请注意，此动态内容具有多个结果，因此"应用到每个 *值"* 控件将围绕 **Condition。**
    * **以 (** 下拉列表列表中的) 
    * **选择一个值**：.csv

    :::image type="content" source="../../images/convert-csv-flow-3.png" alt-text="具有应用于其周围的每个控件的已完成条件控件。":::
1. 流程的其余部分位于"如果是"部分下，因为我们只想处理.csv文件。 通过添加.csv连接器和"获取文件内容 **"操作** 的新OneDrive for Business获取 **单个文件**。 使用 **文件夹中** "列表文件 **"中的动态内容的 Id**。
    * **文件**： *id* (文件夹步骤步骤中的 **列表文件动态**) 
1. 将新的 **"获取文件内容"** 步骤重命名为"get .csv file"。 这有助于将此文件与模板Excel区。
1. 创建新的.xlsx文件，使用Excel模板作为基本内容。 添加一 **个使用** **OneDrive for Business 连接器和****"创建文件"操作的新** 步骤。 使用以下值。
    * **文件夹路径**：/output
    * **文件名：***不带* 扩展名.xlsx (从文件夹中的"列表文件"中选择"没有扩展名动态内容的名称"，.xlsx键入") 
    * **文件内容**：*从"获取 (* 模板"Excel动态内容 **)**

     :::image type="content" source="../../images/convert-csv-flow-4.png" alt-text="获取.csv流的&quot;获取文件&quot;和&quot;创建Power Automate步骤。":::
1. 运行脚本将数据复制到新工作簿。 使用 **"运行Excel" (添加) Online**) **Business 连接器**。 对操作使用以下值。
    * **位置**：OneDrive for Business
    * **文档库**：OneDrive
    * **文件**：*创建* (中的动态 **内容的 id)**
    * **脚本**：转换 CSV
    * **csv**：*获取 (* 文件"中的文件 **.csv动态)**

    :::image type="content" source="../../images/convert-csv-flow-5.png" alt-text="已完成的 Excel Online (Business) 连接器Power Automate。":::
1. 保存流。 使用" **流** 编辑器"页上的"测试"按钮，或通过"我的流" **选项卡运行** 流。请务必在系统提示时允许访问。
1. 您应该在".xlsx"文件夹中找到新的文件，以及原始.csv文件。 新工作簿包含的数据与 CSV 文件相同。

## <a name="troubleshooting"></a>疑难解答

脚本需要逗号分隔的值来生成矩形区域。 如果您的 .csv 文件包含列数不同的行，则您将收到一个错误，指出"输入数组中的行数或列数与区域的大小或维度不匹配"。 如果数据无法符合矩形形状，请改为使用以下脚本。 此脚本一次添加一行数据，而不是一个范围。 此脚本的效率较低，对于大型数据集，速度明显变慢。

```TypeScript
function main(workbook: ExcelScript.Workbook, csv: string) {
  let sheet = workbook.getWorksheet("Sheet1");

  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();

  // Split each line into a row.
  let rows = csv.split("\r\n");
  rows.forEach((value, index) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);

    // Check for blanks at the start of the row.
    if (row[0].charAt(0) === ',') {
      row.unshift("");
    }

    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });

    // Create a 2D-array with one row.
    let data: string[][] = [];
    data.push(row);

    // Put the data in the worksheet.
    let range = sheet.getRangeByIndexes(index, 0, 1, data[0].length);
    range.setValues(data);
  });

  // Add any formatting or table creation that you want.
}
```
