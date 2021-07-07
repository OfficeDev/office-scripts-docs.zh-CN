---
title: 输出Excel JSON
description: 了解如何将Excel数据输出为 JSON，以用于Power Automate。
ms.date: 06/29/2021
localization_priority: Normal
ms.openlocfilehash: 63379d1323f5e2084f4aa39af3f4b6e5e6d7e7bb
ms.sourcegitcommit: 211c157ca746e266eeb079f5fa1925a1e35ab702
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/07/2021
ms.locfileid: "53313944"
---
# <a name="output-excel-table-data-as-json-for-usage-in-power-automate"></a><span data-ttu-id="89148-103">输出Excel数据作为 JSON，用于Power Automate</span><span class="sxs-lookup"><span data-stu-id="89148-103">Output Excel table data as JSON for usage in Power Automate</span></span>

<span data-ttu-id="89148-104">Excel表数据可以表示为 JSON 形式的对象数组。</span><span class="sxs-lookup"><span data-stu-id="89148-104">Excel table data can be represented as an array of objects in the form of JSON.</span></span> <span data-ttu-id="89148-105">每个对象代表表格中的一行。</span><span class="sxs-lookup"><span data-stu-id="89148-105">Each object represents a row in the table.</span></span> <span data-ttu-id="89148-106">这有助于以用户可见的Excel格式从数据中提取数据。</span><span class="sxs-lookup"><span data-stu-id="89148-106">This helps extract the data from Excel in a consistent format that is visible to the user.</span></span> <span data-ttu-id="89148-107">然后，可通过流向其他系统Power Automate数据。</span><span class="sxs-lookup"><span data-stu-id="89148-107">The data can then be given to other systems through Power Automate flows.</span></span>

<span data-ttu-id="89148-108">_输入表数据_</span><span class="sxs-lookup"><span data-stu-id="89148-108">_Input table data_</span></span>

:::image type="content" source="../../images/table-input.png" alt-text="显示输入表数据的工作表。":::

<span data-ttu-id="89148-110">此示例的变体还包括其中一个表格列中的超链接。</span><span class="sxs-lookup"><span data-stu-id="89148-110">A variation of this sample also includes the hyperlinks in one of the table columns.</span></span> <span data-ttu-id="89148-111">这允许在 JSON 中显示其他级别的单元格数据。</span><span class="sxs-lookup"><span data-stu-id="89148-111">This allows additional levels of cell data to be surfaced in the JSON.</span></span>

<span data-ttu-id="89148-112">_包含超链接的输入表数据_</span><span class="sxs-lookup"><span data-stu-id="89148-112">_Input table data that includes hyperlinks_</span></span>

:::image type="content" source="../../images/table-hyperlink-view.png" alt-text="显示格式化为超链接的表格数据的列的工作表。":::

<span data-ttu-id="89148-114">_用于编辑超链接的对话框_</span><span class="sxs-lookup"><span data-stu-id="89148-114">_Dialog to edit hyperlink_</span></span>

:::image type="content" source="../../images/table-hyperlink-edit.png" alt-text="显示更改超链接的选项的&quot;编辑超链接&quot;对话框。":::

## <a name="sample-excel-file"></a><span data-ttu-id="89148-116">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="89148-116">Sample Excel file</span></span>

<span data-ttu-id="89148-117">下载适用于 <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> 工作簿的文件文件。</span><span class="sxs-lookup"><span data-stu-id="89148-117">Download the file <a href="table-data-with-hyperlinks.xlsx">table-data-with-hyperlinks.xlsx</a> for a ready-to-use workbook.</span></span> <span data-ttu-id="89148-118">添加以下脚本以自己试用示例！</span><span class="sxs-lookup"><span data-stu-id="89148-118">Add the following script to try the sample yourself!</span></span>

## <a name="sample-code-return-table-data-as-json"></a><span data-ttu-id="89148-119">示例代码：以 JSON 返回表数据</span><span class="sxs-lookup"><span data-stu-id="89148-119">Sample code: Return table data as JSON</span></span>

> [!NOTE]
> <span data-ttu-id="89148-120">您可以更改 `interface TableData` 结构以匹配表列。</span><span class="sxs-lookup"><span data-stu-id="89148-120">You can change the `interface TableData` structure to match your table columns.</span></span> <span data-ttu-id="89148-121">请注意，对于包含空格的列名，请务必将键放在引号中，如 示例中的 `"Event ID"` 。</span><span class="sxs-lookup"><span data-stu-id="89148-121">Note that for column names with spaces, be sure to place your key in quotation marks, such as with `"Event ID"` in the sample.</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('PlainTable').getTables()[0];

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

// This function converts a 2D-array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j]
    }

    objectArray.push(object);
  }

  return objectArray as TableData[];
}

interface TableData {
  "Event ID": string
  Date: string
  Location: string
  Capacity: string
  Speakers: string
}
```

### <a name="sample-output-from-the-plaintable-worksheet"></a><span data-ttu-id="89148-122">"PlainTable"工作表中的示例输出</span><span class="sxs-lookup"><span data-stu-id="89148-122">Sample output from the "PlainTable" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="sample-code-return-table-data-as-json-with-hyperlink-text"></a><span data-ttu-id="89148-123">示例代码：使用超链接文本以 JSON 格式返回表数据</span><span class="sxs-lookup"><span data-stu-id="89148-123">Sample code: Return table data as JSON with hyperlink text</span></span>

> [!NOTE]
> <span data-ttu-id="89148-124">脚本始终从表的第 4 列 (0) 超链接。</span><span class="sxs-lookup"><span data-stu-id="89148-124">The script always extracts hyperlinks from the 4th column (0 index) of the table.</span></span> <span data-ttu-id="89148-125">通过修改注释下的代码，可以更改该顺序或包含多个列作为超链接数据 `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span><span class="sxs-lookup"><span data-stu-id="89148-125">You can change that order or include multiple columns as hyperlink data by modifying the code under the comment `// For the 4th column (0 index), extract the hyperlink and use that instead of text.`</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "WithHyperLink" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet('WithHyperLink').getTables()[0];

  // Get all the values from the table as text.
  const range = table.getRange();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0)  {
    returnObjects = returnObjectFromValues(range);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));  
  return returnObjects
}

function returnObjectFromValues(range: ExcelScript.Range): TableData[] {
  let values = range.getTexts();
  let objectArray = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i]
      continue;
    }

    let object = {}
    for (let j = 0; j < values[i].length; j++) {
      // For the 4th column (0 index), extract the hyperlink and use that instead of text. 
      if (j === 4) {
        object[objectKeys[j]] = range.getCell(i, j).getHyperlink().address;
      } else {
        object[objectKeys[j]] = values[i][j];
      }
    }

    objectArray.push(object);
  }
  return objectArray as TableData[];
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

### <a name="sample-output-from-the-withhyperlink-worksheet"></a><span data-ttu-id="89148-126">"WithHyperLink"工作表中的示例输出</span><span class="sxs-lookup"><span data-stu-id="89148-126">Sample output from the "WithHyperLink" worksheet</span></span>

```json
[{
    "Event ID": "E107",
    "Date": "2020-12-10",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Debra Berger"
}, {
    "Event ID": "E108",
    "Date": "2020-12-11",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Delia Dennis"
}, {
    "Event ID": "E109",
    "Date": "2020-12-12",
    "Location": "Montgomery",
    "Capacity": "10",
    "Search link": "https://www.google.com/search?q=Montgomery",
    "Speakers": "Diego Siciliani"
}, {
    "Event ID": "E110",
    "Date": "2020-12-13",
    "Location": "Boise",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Boise",
    "Speakers": "Gerhart Moller"
}, {
    "Event ID": "E111",
    "Date": "2020-12-14",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Grady Archie"
}, {
    "Event ID": "E112",
    "Date": "2020-12-15",
    "Location": "Fremont",
    "Capacity": "25",
    "Search link": "https://www.google.com/search?q=Fremont",
    "Speakers": "Irvin Sayers"
}, {
    "Event ID": "E113",
    "Date": "2020-12-16",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Isaiah Langer"
}, {
    "Event ID": "E114",
    "Date": "2020-12-17",
    "Location": "Salt Lake City",
    "Capacity": "20",
    "Search link": "https://www.google.com/search?q=salt+lake+city",
    "Speakers": "Johanna Lorenz"
}]
```

## <a name="use-in-power-automate"></a><span data-ttu-id="89148-127">在 Power Automate</span><span class="sxs-lookup"><span data-stu-id="89148-127">Use in Power Automate</span></span>

<span data-ttu-id="89148-128">若要了解如何在工作流中使用此Power Automate，请参阅使用 Power Automate 创建[自动Power Automate。](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate)</span><span class="sxs-lookup"><span data-stu-id="89148-128">For how to use such a script in Power Automate, see [Create an automated workflow with Power Automate](../../tutorials/excel-power-automate-returns.md#create-an-automated-workflow-with-power-automate).</span></span>
