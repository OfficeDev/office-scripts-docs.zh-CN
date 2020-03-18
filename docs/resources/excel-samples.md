---
title: Web 上的 Excel 中 Office 脚本的示例脚本
description: 要用于 web 上 Excel 中的 Office 脚本的一组代码示例。
ms.date: 02/19/2020
localization_priority: Normal
ms.openlocfilehash: abb4064dfde8b644035e725832e481e6463e979e
ms.sourcegitcommit: b075eed5a6f275274fbbf6d62633219eac416f26
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/10/2020
ms.locfileid: "42700134"
---
# <a name="sample-scripts-for-office-scripts-in-excel-on-the-web-preview"></a><span data-ttu-id="ba0ca-103">Excel 网页版中 Office 脚本的示例脚本（预览）</span><span class="sxs-lookup"><span data-stu-id="ba0ca-103">Sample scripts for Office Scripts in Excel on the web (preview)</span></span>

<span data-ttu-id="ba0ca-104">下面的示例是您在自己的工作簿中尝试的简单脚本。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-104">The following samples are simple scripts for you to try on your own workbooks.</span></span> <span data-ttu-id="ba0ca-105">若要在 Excel 网页上使用它们，请执行以下操作：</span><span class="sxs-lookup"><span data-stu-id="ba0ca-105">To use them in Excel on the web:</span></span>

1. <span data-ttu-id="ba0ca-106">打开 "**自动**" 选项卡。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-106">Open the **Automate** tab.</span></span>
2. <span data-ttu-id="ba0ca-107">按**代码编辑器**。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-107">Press **Code Editor**.</span></span>
3. <span data-ttu-id="ba0ca-108">在代码编辑器的任务窗格中，按 "**新建脚本**"。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-108">Press **New Script** in the Code Editor's task pane.</span></span>
4. <span data-ttu-id="ba0ca-109">将整个脚本替换为您选择的示例。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-109">Replace the entire script with the sample of your choice.</span></span>
5. <span data-ttu-id="ba0ca-110">在代码编辑器的任务窗格中按 "**运行**"。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-110">Press **Run** in the Code Editor's task pane.</span></span>

[!INCLUDE [Preview note](../includes/preview-note.md)]

## <a name="scripting-basics"></a><span data-ttu-id="ba0ca-111">脚本基础</span><span class="sxs-lookup"><span data-stu-id="ba0ca-111">Scripting basics</span></span>

<span data-ttu-id="ba0ca-112">这些示例演示 Office 脚本的基本构建基块。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-112">These samples demonstrate fundamental building blocks for Office Scripts.</span></span> <span data-ttu-id="ba0ca-113">将这些应用程序添加到脚本以扩展解决方案并解决常见问题。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-113">Add these to your scripts to extend your solution and solve common problems.</span></span>

### <a name="read-and-log-one-cell"></a><span data-ttu-id="ba0ca-114">读取和记录一个单元格</span><span class="sxs-lookup"><span data-stu-id="ba0ca-114">Read and log one cell</span></span>

<span data-ttu-id="ba0ca-115">此示例读取**A1**的值并将其打印到控制台。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-115">This sample reads the value of **A1** and prints it to the console.</span></span>

``` TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  range.load("values");
  await context.sync();

  // Print the value of A1.
  console.log(range.values);
}
```

### <a name="work-with-dates"></a><span data-ttu-id="ba0ca-116">使用日期</span><span class="sxs-lookup"><span data-stu-id="ba0ca-116">Work with dates</span></span>

<span data-ttu-id="ba0ca-117">此示例使用 JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date)对象获取当前日期和时间，然后将这些值写入活动工作表中的两个单元格。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-117">This sample uses the JavaScript [Date](https://developer.mozilla.org/docs/web/javascript/reference/global_objects/date) object to get the current date and time and then writes those values to two cells in the active worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the cells at A1 and B1.
  let dateRange = context.workbook.worksheets.getActiveWorksheet().getRange("A1");
  let timeRange = context.workbook.worksheets.getActiveWorksheet().getRange("B1");

  // Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());

  // Add the date string to A1.
  dateRange.values = [[date.toLocaleDateString()]];
  
  // Add the time string to B1.
  timeRange.values = [[date.toLocaleTimeString()]];
}
```

## <a name="display-data"></a><span data-ttu-id="ba0ca-118">显示数据</span><span class="sxs-lookup"><span data-stu-id="ba0ca-118">Display data</span></span>

<span data-ttu-id="ba0ca-119">这些示例演示如何使用工作表数据，并为用户提供更好的视图或组织。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-119">These samples demonstrate how to work with worksheet data and provide users with a better view or organization.</span></span>

### <a name="apply-conditional-formatting"></a><span data-ttu-id="ba0ca-120">应用条件格式</span><span class="sxs-lookup"><span data-stu-id="ba0ca-120">Apply conditional formatting</span></span>

<span data-ttu-id="ba0ca-121">此示例向工作表中当前使用的区域应用条件格式。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-121">This sample applies conditional formatting to the currently used range in the worksheet.</span></span> <span data-ttu-id="ba0ca-122">条件格式是前10% 的数值的绿色填充。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-122">The conditional formatting is a green fill for the top 10% of values.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();

  // Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);
  conditionalFormat.topBottom.format.fill.color = "green";
  conditionalFormat.topBottom.rule = {
    rank: 10, // The percentage threshold.
    type: Excel.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  };
}
```

### <a name="create-a-sorted-table"></a><span data-ttu-id="ba0ca-123">创建已排序的表</span><span class="sxs-lookup"><span data-stu-id="ba0ca-123">Create a sorted table</span></span>

<span data-ttu-id="ba0ca-124">本示例从当前工作表的已用区域创建一个表格，然后基于第一列对其进行排序。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-124">This sample creates a table from the current worksheet's used range, then sorts it based on the first column.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.tables.add(usedRange, true);

  // Sort the table using the first column.
  newTable.sort.apply([{ key: 0, ascending: true }]);
}
```

## <a name="collaboration"></a><span data-ttu-id="ba0ca-125">协作</span><span class="sxs-lookup"><span data-stu-id="ba0ca-125">Collaboration</span></span>

<span data-ttu-id="ba0ca-126">这些示例演示如何使用 Excel 的与协作相关的功能，如注释。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-126">These samples demonstrate how to work with collaboration-related features of Excel, such as comments.</span></span>

### <a name="delete-resolved-comments"></a><span data-ttu-id="ba0ca-127">删除已解决的注释</span><span class="sxs-lookup"><span data-stu-id="ba0ca-127">Delete resolved comments</span></span>

<span data-ttu-id="ba0ca-128">此示例从当前工作表中删除所有已解析的注释。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-128">This sample deletes all resolved comments from the current worksheet.</span></span>

```TypeScript
async function main(context: Excel.RequestContext) {
  // Get the current worksheet.
  let selectedSheet = context.workbook.worksheets.getActiveWorksheet();

  // Get the comments on this worksheet.
  let comments = selectedSheet.comments;
  comments.load("items/resolved");
  await context.sync();

  // Delete the resolved comments.
  comments.items.forEach((comment) => {
      if (comment.resolved) {
          comment.delete();
      }
  });
}
```

## <a name="scenario-samples"></a><span data-ttu-id="ba0ca-129">方案示例</span><span class="sxs-lookup"><span data-stu-id="ba0ca-129">Scenario samples</span></span>

<span data-ttu-id="ba0ca-130">有关 showcasing 大型的真实解决方案的示例，请访问[Office 脚本的示例方案](scenarios/sample-scenario-overview.md)。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-130">For samples showcasing larger, real-world solutions, visit [Sample scenarios for Office Scripts](scenarios/sample-scenario-overview.md).</span></span>

## <a name="suggest-new-samples"></a><span data-ttu-id="ba0ca-131">建议新示例</span><span class="sxs-lookup"><span data-stu-id="ba0ca-131">Suggest new samples</span></span>

<span data-ttu-id="ba0ca-132">我们欢迎您提出新示例建议。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-132">We welcome suggestions for new samples.</span></span> <span data-ttu-id="ba0ca-133">如果有一个可帮助其他脚本开发人员的常见方案，请在下面的 "反馈" 部分告诉我们。</span><span class="sxs-lookup"><span data-stu-id="ba0ca-133">If there is a common scenario that would help other script developers, please tell us in the feedback section below.</span></span>
