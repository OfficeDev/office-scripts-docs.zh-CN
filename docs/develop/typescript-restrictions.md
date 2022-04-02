---
title: 脚本中的 TypeScript Office限制
description: TypeScript 编译器和 linter 的特定信息，Office脚本代码编辑器。
ms.date: 03/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ba0dfe60081a0bb65dec4e694c7d534cb8df63
ms.sourcegitcommit: 7023b9e23499806901a5ecf8ebc460b76887cca6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2022
ms.locfileid: "64585679"
---
# <a name="typescript-restrictions-in-office-scripts"></a>脚本中的 TypeScript Office限制

Office脚本使用 TypeScript 语言。 在大多数情况下，任何 TypeScript 或 JavaScript 代码都适用于Office脚本。 但是，代码编辑器会强制执行一些限制，以确保脚本能够一致且符合您的工作簿Excel运行。

## <a name="no-any-type-in-office-scripts"></a>在脚本中，没有"Office"类型

[在](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) TypeScript 中，写入类型是可选的，因为可以推断出这些类型。 但是Office脚本要求变量不能为[任意类型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。 在脚本中`any`不允许显式和隐式Office脚本。 这些情况报告为错误。

### <a name="explicit-any"></a>显式 `any`

您不能在脚本脚本中`any`显式声明Office类型 (，即) `let value: any;` 。 类型`any`导致由事件处理时Excel。 例如，需要知道`Range`值是 、 `string``number`或 `boolean`。 如果脚本中的类型 `any` 明确定义为 (，在运行脚本之前，) 编译时错误。

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="代码编辑器悬停文本中的显式&quot;any&quot;消息。":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="控制台窗口中的显式&quot;any&quot;错误。":::

在上一个屏幕截图中 `[2, 14] Explicit Any is not allowed` ，指示行 #2，列 #14 定义 `any` 类型。 这可以帮助您找到错误。

若要解决此问题，请始终定义变量的类型。 如果不确定变量的类型，可以使用联合 [类型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。 对于保留值的`Range` `string | number | boolean` `string``number``boolean` `Range`变量（可以是 、 或 (值的类型是以下值之一) 。

### <a name="implicit-any"></a>隐式 `any`

TypeScript 变量类型可以 [隐式](https://www.typescriptlang.org/docs/handbook/type-inference.html) 定义。 如果 TypeScript 编译器无法确定变量 (或者由于类型未显式定义或类型推断不可行) `any` ，则它是隐式的，您将收到编译时错误。

:::image type="content" source="../images/implicit-any-editor-message.png" alt-text="代码编辑器悬停文本中的隐式&quot;any&quot;消息。":::

任何隐式上的最常见情况 `any` 是在变量声明中，例如 `let value;`。 有两种方法可以避免这种情况：

* 将变量分配给隐式可识别的类型 (`let value = 5;` 或) `let value = workbook.getWorksheet();` 。
* 显式键入变量 (`let value: number;`) 

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>不继承Office脚本类或接口

在脚本中创建的类和接口Office[或实现](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)脚本Office脚本类或接口。 换句话说，命名空间中没有任何 `ExcelScript` 内容可以有子类或子接口。

## <a name="incompatible-typescript-functions"></a>不兼容的 TypeScript 函数

Office脚本 API 不能用于以下项：

* [生成器函数](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` 不支持

出于安全考虑，不支持 JavaScript [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 函数。

## <a name="restricted-identifiers"></a>受限标识符

以下单词不能用作脚本中的标识符。 它们是保留条款。

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>仅数组回调中的箭头函数

当为 Array 方法 [提供回调参数时](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) ，脚本只能使用 [箭头](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) 函数。 不能将任何类型的标识符或"传统"函数传递给这些方法。

```TypeScript
const myArray = [1, 2, 3, 4, 5, 6];
let filteredArray = myArray.filter((x) => {
  return x % 2 === 0;
});
/*
  The following code generates a compiler error in the Office Scripts Code Editor.
  filteredArray = myArray.filter(function (x) {
    return x % 2 === 0;
  });
*/
```

## <a name="unions-of-excelscript-types-and-user-defined-types-arent-supported"></a>不支持类型和 `ExcelScript` 用户定义类型的联合

Office脚本在运行时从同步代码块转换为异步代码块。 脚本创建者将隐藏通过 [承诺与](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) 工作簿的通信。 此转换不支持包含 [类型和](https://www.typescriptlang.org/docs/handbook/2/everyday-types.html#union-types) 用户定义 `ExcelScript` 类型的联合类型。 在这种情况下，将`Promise`返回到 脚本，但脚本Office无法预期它，并且脚本创建者无法`Promise`与 交互。

下面的代码示例演示自定义接口和自定义接口 `ExcelScript.Table` 之间的不受支持的联合 `MyTable` 。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  const selectedSheet = workbook.getActiveWorksheet();

  // This union is not supported.
  const tableOrMyTable: ExcelScript.Table | MyTable = selectedSheet.getTables()[0];

  // `getName` returns a promise that can't be resolved by the script.
  const name = tableOrMyTable.getName();

  // This logs "{}" instead of the table name.
  console.log(name);
}

interface MyTable {
  getName(): string
}
```

## <a name="constructors-dont-support-office-scripts-apis-and-console-statements"></a>构造函数不支持脚本Office和`console`语句

`console`语句和许多Office脚本 API 需要与工作簿Excel同步。 这些同步使用 `await` 脚本的已编译运行时版本中的语句。 `await` 在构造函数 [中不受支持](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Classes/constructor)。 如果需要具有构造函数的类，请避免Office脚本 API `console` 或这些代码块中的语句。

以下代码示例演示了此方案。 它生成一个显示 的错误 `failed to load [code] [library]`。

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  class MyClass {
    constructor() {
      // Console statements and Office Scripts APIs aren't supported in constructors.
      console.log("This won't print.");
    }
  }

  let test = new MyClass();
}
```

## <a name="performance-warnings"></a>性能警告

如果脚本可能有性能问题，代码编辑器 [的 linter](https://wikipedia.org/wiki/Lint_(software)) 会发出警告。 改进脚本的性能中记录了这些案例[及其Office记录](web-client-performance.md)。

## <a name="external-api-calls"></a>外部 API 调用

有关详细信息[，请参阅 Office Scripts](external-calls.md) 中的外部 API 调用支持。

## <a name="see-also"></a>另请参阅

* [Excel 网页版中 Office 脚本的脚本基础知识](scripting-fundamentals.md)
* [提高脚本Office性能](web-client-performance.md)
