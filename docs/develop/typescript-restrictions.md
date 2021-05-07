---
title: 脚本中的 TypeScript Office限制
description: TypeScript 编译器和 linter 的特定信息，Office脚本代码编辑器。
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 40eb6923d7b0c47dfeb4c846cdcc745e5d893c13
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232457"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="98140-103">脚本中的 TypeScript Office限制</span><span class="sxs-lookup"><span data-stu-id="98140-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="98140-104">Office脚本使用 TypeScript 语言。</span><span class="sxs-lookup"><span data-stu-id="98140-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="98140-105">在大多数情况下，任何 TypeScript 或 JavaScript 代码都适用于Office脚本。</span><span class="sxs-lookup"><span data-stu-id="98140-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="98140-106">但是，代码编辑器会强制执行一些限制，以确保脚本一致且符合您的工作簿Excel运行。</span><span class="sxs-lookup"><span data-stu-id="98140-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="98140-107">在脚本中，没有"Office"类型</span><span class="sxs-lookup"><span data-stu-id="98140-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="98140-108">在[](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) TypeScript 中，写入类型是可选的，因为可以推断出这些类型。</span><span class="sxs-lookup"><span data-stu-id="98140-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="98140-109">但是，Office脚本要求变量不能为[任何 类型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。</span><span class="sxs-lookup"><span data-stu-id="98140-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="98140-110">在脚本中 `any` 不允许显式和隐式Office脚本。</span><span class="sxs-lookup"><span data-stu-id="98140-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="98140-111">这些情况报告为错误。</span><span class="sxs-lookup"><span data-stu-id="98140-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="98140-112">显式 `any`</span><span class="sxs-lookup"><span data-stu-id="98140-112">Explicit `any`</span></span>

<span data-ttu-id="98140-113">您不能在脚本脚本中显式声明Office类型 `any` (，即 `let someVariable: any;`) 。</span><span class="sxs-lookup"><span data-stu-id="98140-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="98140-114">类型 `any` 导致由事件处理时Excel。</span><span class="sxs-lookup"><span data-stu-id="98140-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="98140-115">例如， `Range` 需要知道值是 、 `string` 或 `number` `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="98140-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="98140-116">如果脚本中的类型明确定义为 (，在运行脚本脚本之前) 出现编译时 `any` 错误。</span><span class="sxs-lookup"><span data-stu-id="98140-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

:::image type="content" source="../images/explicit-any-editor-message.png" alt-text="代码编辑器悬停文本中的显式&quot;any&quot;消息":::

:::image type="content" source="../images/explicit-any-error-message.png" alt-text="控制台窗口中的 Explicit Any 错误":::

<span data-ttu-id="98140-119">在上一 `[5, 16] Explicit Any is not allowed` 张屏幕截图中，#5列#16定义 `any` 类型。</span><span class="sxs-lookup"><span data-stu-id="98140-119">In the previous screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="98140-120">这可以帮助您找到错误。</span><span class="sxs-lookup"><span data-stu-id="98140-120">This helps you locate the error.</span></span>

<span data-ttu-id="98140-121">若要解决此问题，请始终定义变量的类型。</span><span class="sxs-lookup"><span data-stu-id="98140-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="98140-122">如果不确定变量的类型，可以使用联合 [类型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="98140-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="98140-123">对于保留值的变量（可以是 、 或 (值的类型是以下值之一 `Range` `string` `number` `boolean` `Range` `string | number | boolean`) 。</span><span class="sxs-lookup"><span data-stu-id="98140-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="98140-124">隐式 `any`</span><span class="sxs-lookup"><span data-stu-id="98140-124">Implicit `any`</span></span>

<span data-ttu-id="98140-125">TypeScript 变量类型可以 [隐式](https://www.typescriptlang.org/docs/handbook/type-inference.html) 定义。</span><span class="sxs-lookup"><span data-stu-id="98140-125">TypeScript variable types can be [implicitly](https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="98140-126">如果 TypeScript 编译器无法确定变量 (或者因为类型未显式定义或类型推断不可行) ，则它是隐式的，您将收到编译 `any` 时错误。</span><span class="sxs-lookup"><span data-stu-id="98140-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="98140-127">任何隐式上的最常见情况 `any` 是在变量声明中，例如 `let value;` 。</span><span class="sxs-lookup"><span data-stu-id="98140-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="98140-128">有两种方法可以避免这种情况：</span><span class="sxs-lookup"><span data-stu-id="98140-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="98140-129">将变量分配给隐式可识别的类型 (`let value = 5;` 或 `let value = workbook.getWorksheet();`) 。</span><span class="sxs-lookup"><span data-stu-id="98140-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="98140-130">显式键入变量 `let value: number;` () </span><span class="sxs-lookup"><span data-stu-id="98140-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="98140-131">不继承Office脚本类或接口</span><span class="sxs-lookup"><span data-stu-id="98140-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="98140-132">在脚本中创建的类和Office无法[扩展](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance)或实现Office脚本类或接口。</span><span class="sxs-lookup"><span data-stu-id="98140-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="98140-133">换句话说，命名空间中 `ExcelScript` 没有任何内容可以有子类或子接口。</span><span class="sxs-lookup"><span data-stu-id="98140-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="98140-134">不兼容的 TypeScript 函数</span><span class="sxs-lookup"><span data-stu-id="98140-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="98140-135">Office脚本 API 不能用于以下项：</span><span class="sxs-lookup"><span data-stu-id="98140-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="98140-136">生成器函数</span><span class="sxs-lookup"><span data-stu-id="98140-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="98140-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="98140-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="98140-138">`eval` 不支持</span><span class="sxs-lookup"><span data-stu-id="98140-138">`eval` is not supported</span></span>

<span data-ttu-id="98140-139">出于安全考虑，不支持 JavaScript [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 函数。</span><span class="sxs-lookup"><span data-stu-id="98140-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="98140-140">受限标识</span><span class="sxs-lookup"><span data-stu-id="98140-140">Restricted identifers</span></span>

<span data-ttu-id="98140-141">以下单词不能用作脚本中的标识符。</span><span class="sxs-lookup"><span data-stu-id="98140-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="98140-142">它们是保留条款。</span><span class="sxs-lookup"><span data-stu-id="98140-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="only-arrow-functions-in-array-callbacks"></a><span data-ttu-id="98140-143">仅数组回调中的箭头函数</span><span class="sxs-lookup"><span data-stu-id="98140-143">Only arrow functions in array callbacks</span></span>

<span data-ttu-id="98140-144">当为 Array 方法 [提供回调参数时](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) ，脚本只能使用 [箭头](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) 函数。</span><span class="sxs-lookup"><span data-stu-id="98140-144">Your scripts can only use [arrow functions](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Functions/Arrow_functions) when providing callback arguments for [Array](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array) methods.</span></span> <span data-ttu-id="98140-145">不能将任何类型的标识符或"传统"函数传递给这些方法。</span><span class="sxs-lookup"><span data-stu-id="98140-145">You cannot pass any sort of identifier or "traditional" function to these methods.</span></span>

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

## <a name="performance-warnings"></a><span data-ttu-id="98140-146">性能警告</span><span class="sxs-lookup"><span data-stu-id="98140-146">Performance warnings</span></span>

<span data-ttu-id="98140-147">如果脚本可能有性能问题，代码编辑器 [的 linter](https://wikipedia.org/wiki/Lint_(software)) 会发出警告。</span><span class="sxs-lookup"><span data-stu-id="98140-147">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="98140-148">这些案例及其处理的方法记录在 改进 Office[脚本 的性能中](web-client-performance.md)。</span><span class="sxs-lookup"><span data-stu-id="98140-148">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="98140-149">外部 API 调用</span><span class="sxs-lookup"><span data-stu-id="98140-149">External API calls</span></span>

<span data-ttu-id="98140-150">有关详细信息[，请参阅 Office Scripts](external-calls.md)中的外部 API 调用支持。</span><span class="sxs-lookup"><span data-stu-id="98140-150">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="98140-151">另请参阅</span><span class="sxs-lookup"><span data-stu-id="98140-151">See also</span></span>

* [<span data-ttu-id="98140-152">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="98140-152">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="98140-153">提高脚本Office性能</span><span class="sxs-lookup"><span data-stu-id="98140-153">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
