---
title: Office 脚本中的 TypeScript 限制
description: Office 脚本代码编辑器使用的 TypeScript 编译器和 linter 的具体信息。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 87a070b9f342fa5a1f5109fa647bba591832e0cf
ms.sourcegitcommit: 345f1dd96d80471b246044b199fe11126a192a88
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50242016"
---
# <a name="typescript-restrictions-in-office-scripts"></a>Office 脚本中的 TypeScript 限制

Office 脚本使用 TypeScript 语言。 在大多数情况下，任何 TypeScript 或 JavaScript 代码都将在 Office 脚本中运行。 但是，代码编辑器会强制执行一些限制，以确保脚本可以一致且符合 Excel 工作簿的运行要求。

## <a name="no-any-type-in-office-scripts"></a>Office 脚本中无"任何"类型

在[](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) TypeScript 中，写入类型是可选的，因为可以推断出这些类型。 但是，Office 脚本要求变量不能为任何 [类型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。 Office 脚本 `any` 中不允许显式和隐式。 这些情况报告为错误。

### <a name="explicit-any"></a>显式 `any`

您无法在 Office 脚本中显式声明变量的类型 `any` (即 `let someVariable: any;`) 。 该 `any` 类型导致 Excel 处理时出现问题。 例如， `Range` 需要知道值是 `string` ， 或 `number` `boolean` 。 如果在脚本中将任何变量显式定义为 (，则运行脚本脚本之前) 出现编译时 `any` 错误。

![在代码编辑器的悬停文本中显式显示任何消息](../images/explicit-any-editor-message.png)

![控制台窗口中的显式任何错误](../images/explicit-any-error-message.png)

在以上屏幕截图 `[5, 16] Explicit Any is not allowed` 中，指示行#5、列#16定义 `any` 类型。 这可以帮助您找到错误。

若要解决此问题，请始终定义变量的类型。 如果不确定变量的类型，可以使用联合 [类型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。 这可用于保留值（可以是类型）的变量，或者 (值的类型是以下值 `Range` `string` `number` `boolean` `Range` `string | number | boolean`) 。

### <a name="implicit-any"></a>隐式 `any`

TypeScript 变量类型可以 [隐式](https://www.typescriptlang.org/docs/handbook/type-inference.html) 定义。 如果 TypeScript 编译器无法确定变量 (或者由于类型未显式定义或类型推断无法进行) ，则它是隐式的，您将收到编译时错误。 `any`

任何隐式的最常见情况 `any` 都位于变量声明中，例如 `let value;` 。 有两种方法可以避免这种情况：

* 将变量分配给隐式可识别类型 (`let value = 5;` 或 `let value = workbook.getWorksheet();`) 。
* 显式键入变量 `let value: number;` () 

## <a name="no-inheriting-office-script-classes-or-interfaces"></a>没有继承 Office 脚本类或接口

在 Office 脚本中创建的类和接口无法 [扩展或实现](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office 脚本类或接口。 换句话说，命名空间中没有任何内容 `ExcelScript` 可以有子类或子接口。

## <a name="incompatible-typescript-functions"></a>不兼容的 TypeScript 函数

Office 脚本 API 不能用于以下项：

* [生成器函数](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [Array.sort](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a>`eval` 不支持

出于安全考虑，不支持 JavaScript [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 函数。

## <a name="restricted-identifers"></a>受限标识

以下单词不能用作脚本中的标识符。 它们是保留条款。

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a>仅数组回调中的箭头函数

当为 Array 方法 [提供回调参数](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) 时，脚本只能使用 [箭头](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) 函数。 不能将任何标识符或"传统"函数传递给这些方法。

```typescript
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

## <a name="performance-warnings"></a>性能警告

如果脚本可能有性能问题，则代码编辑器 [的 linter](https://wikipedia.org/wiki/Lint_(software)) 会发出警告。 这些案例及其处理的方法记录在 ["提高 Office 脚本的性能"中](web-client-performance.md)。

## <a name="external-api-calls"></a>外部 API 调用

有关详细信息 [，请参阅 Office 脚本中的外部 API](external-calls.md) 调用支持。

## <a name="see-also"></a>另请参阅

* [Excel 网页版中 Office 脚本的脚本基础知识](scripting-fundamentals.md)
* [提高 Office 脚本的性能](web-client-performance.md)
