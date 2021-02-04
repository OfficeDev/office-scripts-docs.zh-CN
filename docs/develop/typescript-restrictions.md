---
title: Office 脚本中的 TypeScript 限制
description: Office 脚本代码编辑器使用的 TypeScript 编译器和 linter 的具体信息。
ms.date: 01/29/2021
localization_priority: Normal
ms.openlocfilehash: 41584ff23b333d17b2e267fdb3b0ec8741f3d203
ms.sourcegitcommit: df2b64603f91acb37bf95230efd538db0fbf9206
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/04/2021
ms.locfileid: "50099897"
---
# <a name="typescript-restrictions-in-office-scripts"></a><span data-ttu-id="411ff-103">Office 脚本中的 TypeScript 限制</span><span class="sxs-lookup"><span data-stu-id="411ff-103">TypeScript restrictions in Office Scripts</span></span>

<span data-ttu-id="411ff-104">Office 脚本使用 TypeScript 语言。</span><span class="sxs-lookup"><span data-stu-id="411ff-104">Office Scripts use the TypeScript language.</span></span> <span data-ttu-id="411ff-105">大多数情况下，任何 TypeScript 或 JavaScript 代码都适用于 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="411ff-105">For the most part, any TypeScript or JavaScript code will work in an Office Script.</span></span> <span data-ttu-id="411ff-106">但是，代码编辑器会强制执行一些限制，以确保脚本一致且按预期方式与 Excel 工作簿一起工作。</span><span class="sxs-lookup"><span data-stu-id="411ff-106">However, there are a few restrictions enforced by the Code Editor to ensure your script works consistently and as intended with your Excel workbook.</span></span>

## <a name="no-any-type-in-office-scripts"></a><span data-ttu-id="411ff-107">Office 脚本中无"任何"类型</span><span class="sxs-lookup"><span data-stu-id="411ff-107">No 'any' type in Office Scripts</span></span>

<span data-ttu-id="411ff-108">在 [TypeScript](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) 中，写入类型是可选的，因为可以推断出这些类型。</span><span class="sxs-lookup"><span data-stu-id="411ff-108">Writing [types](https://www.typescriptlang.org/docs/handbook/typescript-in-5-minutes.html) is optional in TypeScript, because the types can be inferred.</span></span> <span data-ttu-id="411ff-109">但是，Office 脚本要求变量不能为 [任何类型](https://www.typescriptlang.org/docs/handbook/basic-types.html#any)。</span><span class="sxs-lookup"><span data-stu-id="411ff-109">However, Office Script requires that a variable can't be of [type any](https://www.typescriptlang.org/docs/handbook/basic-types.html#any).</span></span> <span data-ttu-id="411ff-110">Office 脚本 `any` 中不允许显式和隐式。</span><span class="sxs-lookup"><span data-stu-id="411ff-110">Both explicit and implicit `any` are not allowed in an Office Script.</span></span> <span data-ttu-id="411ff-111">这些情况报告为错误。</span><span class="sxs-lookup"><span data-stu-id="411ff-111">These cases are reported as errors.</span></span>

### <a name="explicit-any"></a><span data-ttu-id="411ff-112">显式 `any`</span><span class="sxs-lookup"><span data-stu-id="411ff-112">Explicit `any`</span></span>

<span data-ttu-id="411ff-113">您不能在 Office 脚本中显式声明一个类型为 `any` (，即 `let someVariable: any;`) 。</span><span class="sxs-lookup"><span data-stu-id="411ff-113">You cannot explicitly declare a variable to be of type `any` in Office Scripts (that is, `let someVariable: any;`).</span></span> <span data-ttu-id="411ff-114">该 `any` 类型导致 Excel 处理时出现问题。</span><span class="sxs-lookup"><span data-stu-id="411ff-114">The `any` type causes issues when processed by Excel.</span></span> <span data-ttu-id="411ff-115">例如，需要知道值是 ， `Range` `string` 或 `number` `boolean` 。</span><span class="sxs-lookup"><span data-stu-id="411ff-115">For example, a `Range` needs to know that a value is a `string`, `number`, or `boolean`.</span></span> <span data-ttu-id="411ff-116">如果在脚本中将任何变量显式定义为 (，则运行脚本之前，您会收到编译时错误) 错误 `any` 。</span><span class="sxs-lookup"><span data-stu-id="411ff-116">You will receive a compile-time error (an error prior to running the script) if any variable is explicitly defined as the `any` type in the script.</span></span>

![代码编辑器悬停文本中的显式任何消息](../images/explicit-any-editor-message.png)

![控制台窗口中的显式任何错误](../images/explicit-any-error-message.png)

<span data-ttu-id="411ff-119">在以上屏幕截图 `[5, 16] Explicit Any is not allowed` 中，指示行#5，列#16定义 `any` 类型。</span><span class="sxs-lookup"><span data-stu-id="411ff-119">In the above screenshot `[5, 16] Explicit Any is not allowed` indicates that line #5, column #16 defines `any` type.</span></span> <span data-ttu-id="411ff-120">这可以帮助您找到错误。</span><span class="sxs-lookup"><span data-stu-id="411ff-120">This helps you locate the error.</span></span>

<span data-ttu-id="411ff-121">若要解决此问题，请始终定义变量的类型。</span><span class="sxs-lookup"><span data-stu-id="411ff-121">To get around this issue, always define the type of the variable.</span></span> <span data-ttu-id="411ff-122">如果不确定变量的类型，可以使用联合 [类型](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html)。</span><span class="sxs-lookup"><span data-stu-id="411ff-122">If you are uncertain about the type of a variable, you can use a [union type](https://www.typescriptlang.org/docs/handbook/unions-and-intersections.html).</span></span> <span data-ttu-id="411ff-123">这可用于保留值（可以是类型）的变量，或者 (值的类型是以下值 `Range` `string` `number` `boolean` `Range` `string | number | boolean`) 。</span><span class="sxs-lookup"><span data-stu-id="411ff-123">This can be useful for variables that hold `Range` values, which can be of type `string`, `number`, or `boolean` (the type for `Range` values is a union of those: `string | number | boolean`).</span></span>

### <a name="implicit-any"></a><span data-ttu-id="411ff-124">隐式 `any`</span><span class="sxs-lookup"><span data-stu-id="411ff-124">Implicit `any`</span></span>

<span data-ttu-id="411ff-125">TypeScript 变量类型可以[隐式] ( (https://www.typescriptlang.org/docs/handbook/type-inference.html) 定义。</span><span class="sxs-lookup"><span data-stu-id="411ff-125">TypeScript variable types can be [implicitly]((https://www.typescriptlang.org/docs/handbook/type-inference.html) defined.</span></span> <span data-ttu-id="411ff-126">如果 TypeScript 编译器无法确定变量 (或者由于类型未明确定义或类型推断无法进行) ，则它是隐式的，并且您将收到编译时错误。 `any`</span><span class="sxs-lookup"><span data-stu-id="411ff-126">If the TypeScript compiler is unable to determine the type of a variable (either because type is not defined explicitly or type inference isn't possible), then it's an implicit `any` and you will receive a compilation-time error.</span></span>

<span data-ttu-id="411ff-127">任何隐式的最常见情况 `any` 都位于变量声明中，例如 `let value;` 。</span><span class="sxs-lookup"><span data-stu-id="411ff-127">The most common case on any implicit `any` is in a variable declaration, such as `let value;`.</span></span> <span data-ttu-id="411ff-128">有两种方法可以避免这种情况：</span><span class="sxs-lookup"><span data-stu-id="411ff-128">There are two ways to avoid this:</span></span>

* <span data-ttu-id="411ff-129">将变量分配给隐式可识别的类型 (`let value = 5;` 或 `let value = workbook.getWorksheet();`) 。</span><span class="sxs-lookup"><span data-stu-id="411ff-129">Assign the variable to an implicitly identifiable type (`let value = 5;` or `let value = workbook.getWorksheet();`).</span></span>
* <span data-ttu-id="411ff-130">显式键入变量 `let value: number;` () </span><span class="sxs-lookup"><span data-stu-id="411ff-130">Explicitly type the variable (`let value: number;`)</span></span>

## <a name="no-inheriting-office-script-classes-or-interfaces"></a><span data-ttu-id="411ff-131">没有继承 Office 脚本类或接口</span><span class="sxs-lookup"><span data-stu-id="411ff-131">No inheriting Office Script classes or interfaces</span></span>

<span data-ttu-id="411ff-132">在 Office 脚本中创建的类和接口无法 [扩展或实现](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office 脚本类或接口。</span><span class="sxs-lookup"><span data-stu-id="411ff-132">Classes and interfaces that are created in your Office Script cannot [extend or implement](https://www.typescriptlang.org/docs/handbook/classes.html#inheritance) Office Scripts classes or interfaces.</span></span> <span data-ttu-id="411ff-133">换句话说，命名空间中没有任何内容 `ExcelScript` 可以有子类或子接口。</span><span class="sxs-lookup"><span data-stu-id="411ff-133">In other words, nothing in the `ExcelScript` namespace can have subclasses or subinterfaces.</span></span>

## <a name="incompatible-typescript-functions"></a><span data-ttu-id="411ff-134">不兼容的 TypeScript 函数</span><span class="sxs-lookup"><span data-stu-id="411ff-134">Incompatible TypeScript functions</span></span>

<span data-ttu-id="411ff-135">Office 脚本 API 不能用于以下项：</span><span class="sxs-lookup"><span data-stu-id="411ff-135">Office Scripts APIs cannot be used in the following:</span></span>

* [<span data-ttu-id="411ff-136">生成器函数</span><span class="sxs-lookup"><span data-stu-id="411ff-136">Generator functions</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Guide/Iterators_and_Generators#generator_functions)
* [<span data-ttu-id="411ff-137">Array.sort</span><span class="sxs-lookup"><span data-stu-id="411ff-137">Array.sort</span></span>](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Array/sort)

## <a name="eval-is-not-supported"></a><span data-ttu-id="411ff-138">`eval` 不支持</span><span class="sxs-lookup"><span data-stu-id="411ff-138">`eval` is not supported</span></span>

<span data-ttu-id="411ff-139">出于安全考虑，不支持 JavaScript [eval](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) 函数。</span><span class="sxs-lookup"><span data-stu-id="411ff-139">The JavaScript [eval function](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/eval) is not supported for security reasons.</span></span>

## <a name="restricted-identifers"></a><span data-ttu-id="411ff-140">受限制的标识</span><span class="sxs-lookup"><span data-stu-id="411ff-140">Restricted identifers</span></span>

<span data-ttu-id="411ff-141">以下单词不能用作脚本中的标识符。</span><span class="sxs-lookup"><span data-stu-id="411ff-141">The following words can't be used as identifiers in a script.</span></span> <span data-ttu-id="411ff-142">它们是保留条款。</span><span class="sxs-lookup"><span data-stu-id="411ff-142">They are reserved terms.</span></span>

* `Excel`
* `ExcelScript`
* `console`

## <a name="performance-warnings"></a><span data-ttu-id="411ff-143">性能警告</span><span class="sxs-lookup"><span data-stu-id="411ff-143">Performance warnings</span></span>

<span data-ttu-id="411ff-144">如果脚本可能有性能问题，代码编辑器 [的 linter](https://wikipedia.org/wiki/Lint_(software)) 会发出警告。</span><span class="sxs-lookup"><span data-stu-id="411ff-144">The Code Editor's [linter](https://wikipedia.org/wiki/Lint_(software)) gives warnings if the script might have performance issues.</span></span> <span data-ttu-id="411ff-145">这些案例及其处理过程记录在 ["提高 Office 脚本的性能"中](web-client-performance.md)。</span><span class="sxs-lookup"><span data-stu-id="411ff-145">The cases and how to work around them are documented in [Improve the performance of your Office Scripts](web-client-performance.md).</span></span>

## <a name="external-api-calls"></a><span data-ttu-id="411ff-146">外部 API 调用</span><span class="sxs-lookup"><span data-stu-id="411ff-146">External API calls</span></span>

<span data-ttu-id="411ff-147">有关详细信息 [，请参阅 Office 脚本中的外部 API](external-calls.md) 调用支持。</span><span class="sxs-lookup"><span data-stu-id="411ff-147">See [External API call support in Office Scripts](external-calls.md) for more information.</span></span>

## <a name="see-also"></a><span data-ttu-id="411ff-148">另请参阅</span><span class="sxs-lookup"><span data-stu-id="411ff-148">See also</span></span>

* [<span data-ttu-id="411ff-149">Excel 网页版中 Office 脚本的脚本基础知识</span><span class="sxs-lookup"><span data-stu-id="411ff-149">Scripting fundamentals for Office Scripts in Excel on the web</span></span>](scripting-fundamentals.md)
* [<span data-ttu-id="411ff-150">提高 Office 脚本的性能</span><span class="sxs-lookup"><span data-stu-id="411ff-150">Improve the performance of your Office Scripts</span></span>](web-client-performance.md)
