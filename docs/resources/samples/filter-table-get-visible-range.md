---
title: 筛选 Excel 表并获取可见区域
description: 了解如何使用 Office 脚本筛选 Excel 表，并获取作为对象数组的可见区域。
ms.date: 03/16/2021
localization_priority: Normal
ms.openlocfilehash: c0a5842af4a62162225e3fc10203c261b91e010a
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571242"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="94e73-103">筛选 Excel 表并获取可见区域作为 JSON 对象</span><span class="sxs-lookup"><span data-stu-id="94e73-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="94e73-104">此示例筛选 Excel 表，并返回可见区域作为 JSON 对象。</span><span class="sxs-lookup"><span data-stu-id="94e73-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="94e73-105">此 JSON 可以作为较大解决方案的一部分提供给 Power Automate 流。</span><span class="sxs-lookup"><span data-stu-id="94e73-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="94e73-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="94e73-106">Example scenario</span></span>

* <span data-ttu-id="94e73-107">将筛选器应用于表列。</span><span class="sxs-lookup"><span data-stu-id="94e73-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="94e73-108">筛选后提取可见区域。</span><span class="sxs-lookup"><span data-stu-id="94e73-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="94e73-109">组合并返回具有特定 [JSON 结构的对象](#sample-json)。</span><span class="sxs-lookup"><span data-stu-id="94e73-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="94e73-110">示例代码：筛选表并获取可见区域</span><span class="sxs-lookup"><span data-stu-id="94e73-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="94e73-111">以下脚本筛选表并获取可见区域。</span><span class="sxs-lookup"><span data-stu-id="94e73-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="94e73-112">下载示例文件 <a href="table-filter.xlsx">table-filter.xlsx</a> 并使用此脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="94e73-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(v => v[0] as string);
  const uniqueKeys = keyColumnValues.filter((v, i, a) => a.indexOf(v) === i);

  console.log(uniqueKeys);
  const returnObj: ReturnTemplate = {}

  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    const rangeView = table1.getRange().getVisibleView();
    returnObj[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  })
  table1.getColumnByName('Station').getFilter().clear();
  console.log(JSON.stringify(returnObj));
  return returnObj
}

function returnObjectFromValues(values: string[][]): BasicObj[] {
  let objArray = [];
  let objKeys: string[] = [];
  for (let i=0; i < values.length; i++) {
    if (i===0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j=0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray;
}

interface BasicObj {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObj[]
}
```

### <a name="sample-json"></a><span data-ttu-id="94e73-113">示例 JSON</span><span class="sxs-lookup"><span data-stu-id="94e73-113">Sample JSON</span></span>

<span data-ttu-id="94e73-114">每个键表示表的唯一值。</span><span class="sxs-lookup"><span data-stu-id="94e73-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="94e73-115">每个数组实例表示应用相应筛选器时可见的行。</span><span class="sxs-lookup"><span data-stu-id="94e73-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason": ""
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason": ""
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason": ""
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason": ""
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason": ""
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="94e73-116">培训视频：筛选 Excel 表并获取可见区域</span><span class="sxs-lookup"><span data-stu-id="94e73-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="94e73-117">[![观看有关如何筛选 Excel 表和获取可见范围的分步视频](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "如何筛选 Excel 表和获取可见范围的分步视频")</span><span class="sxs-lookup"><span data-stu-id="94e73-117">[![Watch step-by-step video on how to filter an Excel table and get the visible range](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "Step-by-step video on how to filter an Excel table and get the visible range")</span></span>
