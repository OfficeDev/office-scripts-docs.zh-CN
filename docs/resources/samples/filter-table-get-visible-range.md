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
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>筛选 Excel 表并获取可见区域作为 JSON 对象

此示例筛选 Excel 表，并返回可见区域作为 JSON 对象。 此 JSON 可以作为较大解决方案的一部分提供给 Power Automate 流。

## <a name="example-scenario"></a>示例应用场景

* 将筛选器应用于表列。
* 筛选后提取可见区域。
* 组合并返回具有特定 [JSON 结构的对象](#sample-json)。

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>示例代码：筛选表并获取可见区域

以下脚本筛选表并获取可见区域。

下载示例文件 <a href="table-filter.xlsx">table-filter.xlsx</a> 并使用此脚本尝试一下！

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

### <a name="sample-json"></a>示例 JSON

每个键表示表的唯一值。 每个数组实例表示应用相应筛选器时可见的行。

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>培训视频：筛选 Excel 表并获取可见区域

[![观看有关如何筛选 Excel 表和获取可见范围的分步视频](../../images/visible-range-as-objects-vid.jpg)](https://youtu.be/Mv7BrvPq84A "如何筛选 Excel 表和获取可见范围的分步视频")
