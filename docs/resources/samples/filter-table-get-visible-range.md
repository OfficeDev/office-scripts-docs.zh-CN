---
title: 筛选Excel并获取可见区域
description: 了解如何使用 Office Scripts 筛选 Excel 表，并获取作为对象数组的可见区域。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: a310857e6055b3da57c353dc7ad78a6fbdd86d4e
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232373"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a><span data-ttu-id="59450-103">筛选Excel，并获取作为 JSON 对象的可见区域</span><span class="sxs-lookup"><span data-stu-id="59450-103">Filter Excel table and get visible range as a JSON object</span></span>

<span data-ttu-id="59450-104">此示例筛选一Excel，并作为 JSON 对象返回可见区域。</span><span class="sxs-lookup"><span data-stu-id="59450-104">This sample filters an Excel table and returns the visible range as a JSON object.</span></span> <span data-ttu-id="59450-105">此 JSON 可以作为较大解决方案的Power Automate提供给一个流。</span><span class="sxs-lookup"><span data-stu-id="59450-105">This JSON could be provided to a Power Automate flow as part of a larger solution.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="59450-106">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="59450-106">Example scenario</span></span>

* <span data-ttu-id="59450-107">将筛选器应用于表列。</span><span class="sxs-lookup"><span data-stu-id="59450-107">Apply a filter to a table column.</span></span>
* <span data-ttu-id="59450-108">筛选后提取可见区域。</span><span class="sxs-lookup"><span data-stu-id="59450-108">Extract the visible range after filtering.</span></span>
* <span data-ttu-id="59450-109">组合并返回具有特定 [JSON 结构的对象](#sample-json)。</span><span class="sxs-lookup"><span data-stu-id="59450-109">Assemble and return an object with a [specific JSON structure](#sample-json).</span></span>

## <a name="sample-code-filter-a-table-and-get-visible-range"></a><span data-ttu-id="59450-110">示例代码：筛选表并获取可见区域</span><span class="sxs-lookup"><span data-stu-id="59450-110">Sample code: Filter a table and get visible range</span></span>

<span data-ttu-id="59450-111">以下脚本筛选表并获取可见区域。</span><span class="sxs-lookup"><span data-stu-id="59450-111">The following script filters a table and gets the visible range.</span></span>

<span data-ttu-id="59450-112">下载示例文件 <a href="table-filter.xlsx">table-filter.xlsx</a> 并使用此脚本尝试一下！</span><span class="sxs-lookup"><span data-stu-id="59450-112">Download the sample file <a href="table-filter.xlsx">table-filter.xlsx</a> and use it with this script to try it out yourself!</span></span>

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

### <a name="sample-json"></a><span data-ttu-id="59450-113">示例 JSON</span><span class="sxs-lookup"><span data-stu-id="59450-113">Sample JSON</span></span>

<span data-ttu-id="59450-114">每个键表示表的唯一值。</span><span class="sxs-lookup"><span data-stu-id="59450-114">Each key represents a unique value of a table.</span></span> <span data-ttu-id="59450-115">每个数组实例表示应用相应筛选器时可见的行。</span><span class="sxs-lookup"><span data-stu-id="59450-115">Each array instance represents the row that is visible when the corresponding filter is applied.</span></span>

```json
{
  "Station-1": [{
    "Station": "Station-1",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Debra Berger",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "27-Oct-20",
    "Responsible": "Delia Dennis",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-1",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Lidia Holloway",
    "Reason&quot;: &quot;"
  }],
  "Station-2": [{
    "Station": "Station-2",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Gerhart Moller",
    "Reason&quot;: &quot;"
  }, {
    "Station": "Station-2",
    "Shift": "Afternoon",
    "Date": "28-Oct-20",
    "Responsible": "Grady Archie",
    "Reason&quot;: &quot;"
  }],
  "Station-3": [{
    "Station": "Station-3",
    "Shift": "Morning",
    "Date": "27-Oct-20",
    "Responsible": "Isaiah Langer",
    "Reason&quot;: &quot;"
  }]
}
```

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a><span data-ttu-id="59450-116">培训视频：筛选Excel表并获取可见区域</span><span class="sxs-lookup"><span data-stu-id="59450-116">Training video: Filter an Excel table and get the visible range</span></span>

<span data-ttu-id="59450-117">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/Mv7BrvPq84A)。</span><span class="sxs-lookup"><span data-stu-id="59450-117">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/Mv7BrvPq84A).</span></span>
