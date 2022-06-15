---
title: 筛选Excel表并获取可见范围
description: 了解如何使用Office脚本筛选Excel表，并获取可见范围作为对象数组。
ms.date: 03/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 103ec97111720ab872c0be843aa0573781d98c44
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088083"
---
# <a name="filter-excel-table-and-get-visible-range-as-a-json-object"></a>筛选Excel表，并以 JSON 对象形式获取可见范围

此示例筛选Excel表，并将可见范围返回为 [JSON](https://www.w3schools.com/whatis/whatis_json.asp) 对象。 此 JSON 可作为更大解决方案的一部分提供给Power Automate流。

## <a name="example-scenario"></a>示例方案

* 将筛选器应用于表列。
* 筛选后提取可见范围。
* 组合并返回具有 [特定 JSON 结构](#sample-json)的对象。

## <a name="sample-excel-file"></a>示例Excel文件

下载现成工作簿 <a href="table-filter.xlsx"> 的table-filter.xlsx</a> 。 添加以下脚本以自行尝试示例！

## <a name="sample-code-filter-a-table-and-get-visible-range"></a>示例代码：筛选表并获取可见范围

```TypeScript
function main(workbook: ExcelScript.Workbook): ReturnTemplate {
  // Get the "Station" column to use as key values in the filter.
  const table1 = workbook.getTable("Table1");
  const keyColumnValues: string [] = table1.getColumnByName('Station').getRangeBetweenHeaderAndTotal().getValues().map(value => value[0] as string);

  // Filter out repeated keys. This call to `filter` only returns the first instance of every unique element in the array.
  const uniqueKeys = keyColumnValues.filter((value, index, array) => array.indexOf(value) === index);
  console.log(uniqueKeys);

  const stationData: ReturnTemplate = {};

  // Filter the table to show only rows corresponding to each key.
  uniqueKeys.forEach((key: string) => {
    table1.getColumnByName('Station').getFilter()
      .applyValuesFilter([key]);
    
    // Get the visible view when a single filter is active.
    const rangeView = table1.getRange().getVisibleView();

    // Create a JSON object with every visible row.
    stationData[key] = returnObjectFromValues(rangeView.getValues() as string[][]);
  });

  // Remove the filters.
  table1.getColumnByName('Station').getFilter().clear();

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(stationData));
  return stationData;
}

// This function converts a 2D-array of values into a generic JSON object.
function returnObjectFromValues(values: string[][]): BasicObject[] {
  let objectArray: BasicObject[] = [];
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

  return objectArray;
}

interface BasicObject {
  [key: string] : string
}

interface ReturnTemplate {
  [key: string]: BasicObject[]
}
```

### <a name="sample-json"></a>示例 JSON

每个键表示表的唯一值。 每个数组实例表示应用相应筛选器时可见的行。 有关使用 JSON 的详细信息，请阅读[使用 JSON 向Office脚本传递数据](../../develop/use-json.md)。

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

## <a name="training-video-filter-an-excel-table-and-get-the-visible-range"></a>培训视频：筛选Excel表并获取可见范围

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/Mv7BrvPq84A)。
