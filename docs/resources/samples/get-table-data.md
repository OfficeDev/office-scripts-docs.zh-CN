---
title: 将 Excel 数据输出为 JSON
description: 了解如何将 Excel 表数据输出为 JSON 以在 Power Automate 中使用。
ms.date: 03/18/2021
localization_priority: Normal
ms.openlocfilehash: db6eb8f8645079eebc369e0a0622539075853953
ms.sourcegitcommit: 45ffe3dbd2c834b78592ad35928cf8096f5e80bc
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/14/2021
ms.locfileid: "51754794"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="3006e-103">将 Excel 表数据输出为 JSON，以在 Power Automate 中使用</span><span class="sxs-lookup"><span data-stu-id="3006e-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="3006e-104">Excel 表数据可以表示为 JSON 形式的对象数组。</span><span class="sxs-lookup"><span data-stu-id="3006e-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="3006e-105">每个对象代表表格中的一行。</span><span class="sxs-lookup"><span data-stu-id="3006e-105">Each object represents a row in the table.</span></span> <span data-ttu-id="3006e-106">这有助于以用户可见的一致格式从 Excel 中提取数据。</span><span class="sxs-lookup"><span data-stu-id="3006e-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="3006e-107">然后，可通过 Power Automate 流向其他系统提供数据。</span><span class="sxs-lookup"><span data-stu-id="3006e-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="3006e-108">_输入表数据_</span><span class="sxs-lookup"><span data-stu-id="3006e-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="显示输入表数据的工作表。":::

<span data-ttu-id="3006e-110">此示例的变体还包括其中一个表格列中的超链接。</span><span class="sxs-lookup"><span data-stu-id="3006e-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="3006e-111">这允许在 JSON 中显示其他级别的单元格数据。</span><span class="sxs-lookup"><span data-stu-id="3006e-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="3006e-112">_包含超链接的输入表数据_</span><span class="sxs-lookup"><span data-stu-id="3006e-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="显示格式化为超链接的表格数据的列的工作表。":::

<span data-ttu-id="3006e-114">_用于编辑超链接的对话框_</span><span class="sxs-lookup"><span data-stu-id="3006e-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="显示更改超链接的选项的&quot;编辑超链接&quot;对话框。":::

## <a name="sample-excel-file"></a><span data-ttu-id="3006e-116">示例 Excel 文件</span><span class="sxs-lookup"><span data-stu-id="3006e-116">Sample Excel file</span></span>

<span data-ttu-id="3006e-117">下载这些 <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> 中使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="3006e-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> used in these samples and try it out yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="3006e-118">示例代码：以 JSON 返回表数据</span><span class="sxs-lookup"><span data-stu-id="3006e-118">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="3006e-119">您可以更改 `interface TableData` 结构以匹配表列。</span><span class="sxs-lookup"><span data-stu-id="3006e-119">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="3006e-120">请注意，对于包含空格的列名，请务必将键放在引号中，如 示例中的 `"Event ID"` 。</span><span class="sxs-lookup"><span data-stu-id="3006e-120">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('PlainTable').getTables()[0];
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][]): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="3006e-121">示例输出</span><span class="sxs-lookup"><span data-stu-id="3006e-121">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="3006e-122">示例代码：使用超链接文本以 JSON 格式返回表数据</span><span class="sxs-lookup"><span data-stu-id="3006e-122">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="3006e-123">脚本始终从表的第 4 列 (0) 超链接。</span><span class="sxs-lookup"><span data-stu-id="3006e-123">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="3006e-124">通过修改注释下的代码，可以更改该顺序或包含多个列作为超链接数据 `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="3006e-124">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];
  const range = table.getRange();
  // If you know the table name, you can also do the following:
  // const table = workbook.getTable('Table13436');
  const texts = table.getRange().getTexts();
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts, range);
  } 
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(values: string[][], range: ExcelScript.Range): TableData[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        obj[objKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        obj[objKeys[j]] = values[i][j];
      }
    }
    objArray.push(obj);
  }
  return objArray as TableData[];
}

interface BasicObj {
  [key: string]: string
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  "Search link": string
  Speakers: string
}
```

### <a name="sample-output"></a><span data-ttu-id="3006e-125">示例输出</span><span class="sxs-lookup"><span data-stu-id="3006e-125">Sample output</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers&quot;: &quot;Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers&quot;: &quot;Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers&quot;: &quot;Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers&quot;: &quot;Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="3006e-126">在 Power Automate 中的使用</span><span class="sxs-lookup"><span data-stu-id="3006e-126">Use in Power Automate</span></span>

<span data-ttu-id="3006e-127">有关在 Power Automate 中如何使用此类脚本的信息，请参阅 [使用 Power Automate 创建自动化工作流](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)。</span><span class="sxs-lookup"><span data-stu-id="3006e-127">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
