---
title: 使用 JSON 将数据传入和传入Office脚本
description: 了解如何将数据结构化为 JSON 对象以用于外部调用和Power Automate
ms.date: 06/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 753097183a18f5d20ca2c78a3748c7a1d968ad42
ms.sourcegitcommit: dd01979d34b3499360d2f79a56f8a8f24f480eed
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/15/2022
ms.locfileid: "66088147"
---
# <a name="use-json-to-pass-data-to-and-from-office-scripts"></a>使用 JSON 将数据传入和传入Office脚本

[JSON (JavaScript 对象表示法) ](https://www.w3schools.com/whatis/whatis_json.asp) 是用于存储和传输数据的格式。 每个 JSON 对象都是可在创建时定义的名称/值对的集合。 JSON 适用于Office脚本，因为它可以处理Excel中范围、表和其他数据模式的任意复杂性。 通过 JSON 可以分析[来自 Web 服务](external-calls.md)的传入数据，并通过[Power Automate流](power-automate-integration.md)传递复杂对象。

本文重点介绍如何将 JSON 与Office脚本配合使用。 建议首先从 W3 学校的 [JSON 简介](https://www.w3schools.com/js/js_json_intro.asp) 等文章中了解有关格式的详细信息。

## <a name="parse-json-data-into-a-range-or-table"></a>将 JSON 数据分析到区域或表中

JSON 对象数组提供一致的方式在应用程序和 Web 服务之间传递表数据行。 在这些情况下，每个 JSON 对象表示一行，而属性表示列。 Office脚本可以循环访问 JSON 数组，并将其重新组合为 2D 数组。 然后将此数组设置为区域的值，并存储在工作簿中。 也可以将属性名称添加为标头来创建表。

以下脚本显示要转换为表的 JSON 数据。 请注意，数据不是从外部源获取的。 本文稍后将介绍这一点。

```typescript
/**
 * Sample JSON data. This would be replaced by external calls or
 * parameters getting data from Power Automate in a production script.
 */
const jsonData = [
  { "Action": "Edit", /* Action property with value of "Edit". */
    "N": 3370, /* N property with value of 3370. */
    "Percent": 17.85 /* Percent property with value of 17.85. */
  },
  // The rest of the object entries follow the same pattern.
  { "Action": "Paste", "N": 1171, "Percent": 6.2 },
  { "Action": "Clear", "N": 599, "Percent": 3.17 },
  { "Action": "Insert", "N": 352, "Percent": 1.86 },
  { "Action": "Delete", "N": 350, "Percent": 1.85 },
  { "Action": "Refresh", "N": 314, "Percent": 1.66 },
  { "Action": "Fill", "N": 286, "Percent": 1.51 },
];

/**
 * This script converts JSON data to an Excel table.
 */
function main(workbook: ExcelScript.Workbook) {
  // Create a new worksheet to store the imported data.
  const newSheet = workbook.addWorksheet();
  newSheet.activate();

  // Determine the data's shape by getting the properties in one object.
  // This assumes all the JSON objects have the same properties.
  const columnNames = getPropertiesFromJson(jsonData[0]);

  // Create the table headers using the property names.
  const headerRange = newSheet.getRangeByIndexes(0, 0, 1, columnNames.length);
  headerRange.setValues([columnNames]);

  // Create a new table with the headers.
  const newTable = newSheet.addTable(headerRange, true);

  // Add each object in the array of JSON objects to the table.
  const tableValues = jsonData.map(row => convertJsonToRow(row));
  newTable.addRows(-1, tableValues);
}

/**
 * This function turns a JSON object into an array to be used as a table row.
 */
function convertJsonToRow(obj: object) {
  const array: (string | number)[] = [];

  // Loop over each property and get the value. Their order will be the same as the column headers.
  for (let value in obj) {
    array.push(obj[value]);
  }
  return array;
}

/**
 * This function gets the property names from a single JSON object.
 */
function getPropertiesFromJson(obj: object) {
  const propertyArray: string[] = [];
  
  // Loop over each property in the object and store the property name in an array.
  for (let property in obj) {
    propertyArray.push(property);
  }

  return propertyArray;
}
```

> [!TIP]
> 如果知道 JSON 的结构，可以创建自己的接口，以便更轻松地获取特定属性。 可以将 JSON 到数组转换步骤替换为类型安全引用。 下面的代码片段显示这些步骤 (现在注释出来) 替换为使用新 `ActionRow` 接口的调用。 请注意，这使函 `convertJsonToRow` 数不再必要。
>
> ```typescript
>   // const tableValues = jsonData.map(row => convertJsonToRow(row));
>   // newTable.addRows(-1, tableValues);
>   // }
>
>      const actionRows: ActionRow[] = jsonData as ActionRow[];
>      // Add each object in the array of JSON objects to the table.
>      const tableValues = actionRows.map(row => [row.Action, row.N, row.Percent]);
>      newTable.addRows(-1, tableValues);
>    }
>    
>    interface ActionRow {
>      Action: string;
>      N: number;
>      Percent: number;
>    }
> ```

### <a name="get-json-data-from-external-sources"></a>从外部源获取 JSON 数据

可以通过两种方法通过Office脚本将 JSON 数据导入工作簿。

- 作为具有Power Automate流的[参数](power-automate-integration.md#main-parameters-pass-data-to-a-script)。
- `fetch`调用[外部 Web 服务](external-calls.md)。

#### <a name="modify-the-sample-to-work-with-power-automate"></a>修改示例以使用Power Automate

Power Automate中的 JSON 数据可以作为泛型对象数组传递。 将属性 `object[]` 添加到脚本以接受该数据。

```typescript
// For Power Automate, replace the main signature in the previous sample with this one
// and remove the sample data.
function main(workbook: ExcelScript.Workbook, jsonData: object[]) {
```

然后，你将在Power Automate连接器中看到要添加`jsonData`到 **“运行脚本**”操作的选项。

:::image type="content" source="../images/json-parameter-power-automate.png" alt-text="Excel联机 (业务) 连接器，其中显示了具有 jsonData 参数的运行脚本操作。":::

#### <a name="modify-the-sample-to-use-a-fetch-call"></a>修改示例以使用调用`fetch`

Web 服务可以使用 JSON 数据回复 `fetch` 呼叫。 这会为脚本提供在保持Excel时所需的数据。 通过阅读[Office脚本中的外部 API 调用支持](external-calls.md)，了解有关`fetch`外部调用和外部调用的详细信息。

```typescript
// For external services, replace the main signature in the previous sample with this one,
// add the fetch call, and remove the sample data.
async function main(workbook: ExcelScript.Workbook) {
  // Replace WEB_SERVICE_URL with the URL of whatever service you need to call.
  const response = await fetch('WEB_SERVICE_URL');
  const jsonData: object[] = await response.json();
```

## <a name="create-json-from-a-range"></a>从某个范围创建 JSON

工作表的行和列通常表示其数据值之间的关系。 表的一行在概念上映射到编程对象，每个列都是该对象的属性。 请考虑下表中的数据。 每行表示电子表格中记录的事务。

|ID |Date     |Amount |供应商                        |
|:--|:--------|:------|:-----------------------------|
|1  |6/1/2022 |$43.54 |最适合你有机公司 |
|2  |6/3/2022 |$67.23 |自由面包店和咖啡馆       |
|3  |6/3/2022 |$37.12 |最适合你有机公司 |
|4  |6/6/2022 |$86.95 |Coho Vineyard                 |
|5  |6/7/2022 |$13.64 |自由面包店和咖啡馆       |

每个事务 (每行) 都有一组与之关联的属性：“ID”、“Date”、“Amount”和“Vendor”。 这可以在Office脚本中建模为对象。

```typescript
// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

示例表中的行对应于接口中的属性，因此脚本可以轻松地将每一行转换为 `Transaction` 对象。 在输出Power Automate的数据时，这很有用。 以下脚本循环访问表中的每一行，并将其添加到一个 `Transaction[]`。

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Create an array of Transactions and add each row to it.
  let transactions: Transaction[] = [];
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  for (let i = 0; i < dataValues.length; i++) {
    let row = dataValues[i];
    let currentTransaction: Transaction = {
      ID: row[table.getColumnByName("ID").getIndex()] as string,
      Date: row[table.getColumnByName("Date").getIndex()] as number,
      Amount: row[table.getColumnByName("Amount").getIndex()] as number,
      Vendor: row[table.getColumnByName("Vendor").getIndex()] as string
    };
    transactions.push(currentTransaction);
  }

  // Do something with the Transaction objects, such as return them to a Power Automate flow.
  console.log(transactions);
}

// An interface that wraps transaction details as JSON.
interface Transaction {
  "ID": string;
  "Date": number;
  "Amount": number;
  "Vendor": string;
}
```

:::image type="content" source="../images/create-json-console-output.png" alt-text="上一个脚本中的控制台输出，其中显示了对象的属性值。":::

### <a name="use-a-generic-object"></a>使用泛型对象

上一个示例假定表标头值是一致的。 如果表具有变量列，则需要创建泛型 JSON 对象。 以下脚本显示一个脚本，该脚本将任何表记录为 JSON。

```typescript
function main(workbook: ExcelScript.Workbook) {
  // Get the table on the current worksheet.
  const table = workbook.getActiveWorksheet().getTables()[0];

  // Use the table header names as JSON properties.
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Get each data row in the table.
  const dataValues = table.getRangeBetweenHeaderAndTotal().getValues();
  let jsonArray: object[] = [];

  // For each row, create a JSON object and assign each property to it based on the table headers.
  for (let i = 0; i < dataValues.length; i++) {
    // Create a blank generic JSON object.
    let jsonObject: { [key: string]: string } = {};
    for (let j = 0; j < dataValues[i].length; j++) {
      jsonObject[tableHeaders[j]] = dataValues[i][j] as string;
    }

    jsonArray.push(jsonObject);
  }

  // Do something with the objects, such as return them to a Power Automate flow.
  console.log(jsonArray);
}

```

## <a name="see-also"></a>另请参阅

- [Office 脚本中的外部 API 呼叫支持](external-calls.md)
- [示例：在Office脚本中使用外部提取调用](../resources/samples/external-fetch-calls.md)
- [使用Power Automate运行Office脚本](power-automate-integration.md)