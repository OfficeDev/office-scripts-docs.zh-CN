---
title: 在 Office 脚本中使用数据透视表
description: 了解 Office 脚本 JavaScript API 中数据透视表的对象模型。
ms.date: 04/20/2022
ms.localizationpriority: medium
ms.openlocfilehash: a457c41bd1205f4e17636c43d7ba78addc80d0e4
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572582"
---
# <a name="work-with-pivottables-in-office-scripts"></a>在 Office 脚本中使用数据透视表

借助数据透视表，可以快速分析大量数据集合。 随着他们的力量来复杂。 使用 Office 脚本 API 可以自定义数据透视表以满足你的需求，但 API 集的范围使入门成为一项挑战。 本文演示如何执行常见的数据透视表任务，并说明重要的类和方法。

> [!NOTE]
> 若要更好地了解 API 使用的术语的上下文，请先阅读 Excel 的数据透视表文档。 首先 [创建数据透视表来分析工作表数据](https://support.microsoft.com/office/a9a84538-bfe9-40a9-a8e9-f99134456576)。

## <a name="object-model"></a>对象模型

:::image type="content" source="../images/pivottable-object-model.png" alt-text="使用数据透视表时使用的类、方法和属性的简化图片。":::

[数据透视表](/javascript/api/office-scripts/excelscript/excelscript.pivottable)是 Office 脚本 API 中数据透视表的中心对象。

- [Workbook](/javascript/api/office-scripts/excelscript/excelscript.workbook) 对象包含所有[数据透视表的](/javascript/api/office-scripts/excelscript/excelscript.pivottable)集合。 每个 [工作表](/javascript/api/office-scripts/excelscript/excelscript.worksheet) 还包含该工作表的本地数据透视表集合。
- [数据透视表](/javascript/api/office-scripts/excelscript/excelscript.pivottable)包含 [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy)。 层次结构可以视为表中的列。
- [PivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) 可以添加为行或列 ([RowColumnPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.rowcolumnpivothierarchy)) 、 [数据 (DataPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.datapivothierarchy)) 或 [FilterPivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy))  (筛选器。
- 每个 [PivotHierarchy](/javascript/api/office-scripts/excelscript/excelscript.pivothierarchy) 只包含一个 [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield)。 Excel 外部的数据透视表结构可能包含每个层次结构中的多个字段，因此存在此设计以支持将来的选项。 对于 Office 脚本，字段和层次结构映射到相同的信息。
- [PivotField](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) 包含多个 [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem)。 每个 PivotItem 都是字段中的唯一值。 将每个项视为表列中的值。 如果字段用于数据，则项也可以是聚合值，例如总和。
- [PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) 定义如何显示 [PivotFields](/javascript/api/office-scripts/excelscript/excelscript.pivotfield) 和 [PivotItems](/javascript/api/office-scripts/excelscript/excelscript.pivotitem)。
- [数据透视文件使用](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters) 不同的条件从 [数据透视表](/javascript/api/office-scripts/excelscript/excelscript.pivottable) 筛选数据。

了解这些关系在实践中的工作原理。 以下数据介绍了各种农场的水果销售情况。 它是本文中所有示例的基础。 使用 [pivottable-sample.xlsx](pivottable-sample.xlsx) 继续操作。

:::image type="content" source="../images/pivottable-raw-data.png" alt-text="不同农场不同类型的水果销售的集合。":::

## <a name="create-a-pivottable-with-fields"></a>使用字段创建数据透视表

数据透视表是使用对现有数据的引用创建的。 范围和表都可以是数据透视表的源。 它们还需要在工作簿中存在一个位置。 由于数据透视表的大小是动态的，因此只指定目标范围的左上角。

以下代码片段基于一系列数据创建数据透视表。 数据透视表没有层次结构，因此数据尚未以任何方式分组。

```typescript
  const dataSheet = workbook.getWorksheet("Data");
  const pivotSheet = workbook.getWorksheet("Pivot");

  const farmPivot = pivotSheet.addPivotTable(
    "Farm Pivot", /* The name of the PivotTable. */
    dataSheet.getUsedRange(), /* The source data range. */
    pivotSheet.getRange("A1") /* The location to put the new PivotTable. */);
```

:::image type="content" source="../images/pivottable-empty.png" alt-text="名为“Farm Pivot”的数据透视表，没有层次结构。":::

### <a name="hierarchies-and-fields"></a>层次结构和字段

数据透视表是通过层次结构组织的。 添加为特定类型的层次结构时，这些层次结构用于透视数据。 有四种类型的层次结构。

- **行**：在水平行中显示项。
- **列**：显示垂直列中的项。
- **数据**：显示基于行和列的值聚合。
- **筛选器**：添加或删除数据透视表中的项。

数据透视表可以分配给这些特定层次结构的字段数量或数量。 数据透视表至少需要一个数据层次结构来显示汇总的数字数据，至少需要一行或一列来透视该摘要。 以下代码片段添加了两个行层次结构和两个数据层次结构。

```typescript
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Farm"));
  farmPivot.addRowHierarchy(farmPivot.getHierarchy("Type"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold at Farm"));
  farmPivot.addDataHierarchy(farmPivot.getHierarchy("Crates Sold Wholesale"));
```

:::image type="content" source="../images/pivottable-data-hierarchy.png" alt-text="一个数据透视表，显示根据来自农场的不同水果的总销售额。":::

## <a name="layout-ranges"></a>布局范围

数据透视表的每个部分都映射到一个区域。 这样，脚本就可以从数据透视表获取数据，供稍后在脚本中使用或在 [Power Automate 流](power-automate-integration.md)中返回。 这些范围是通过从`PivotTable.getLayout()`中获取[的 PivotLayout](/javascript/api/office-scripts/excelscript/excelscript.pivotlayout) 对象访问的。 下图显示了方法在 `PivotLayout`其中返回的范围。

:::image type="content" source="../images/pivottable-layout-breakdown.png" alt-text="显示布局的获取范围函数返回数据透视表的哪些部分的图表。":::

## <a name="filters-and-slicers"></a>筛选器和切片器

有三种方法可以筛选数据透视表。

- [FilterPivotHierarchies](/javascript/api/office-scripts/excelscript/excelscript.filterpivothierarchy)
- [PivotFilters](/javascript/api/office-scripts/excelscript/excelscript.pivotfilters)
- [Slicers](/javascript/api/office-scripts/excelscript/excelscript.slicer)

### <a name="filterpivothierarchies"></a>FilterPivotHierarchies

`FilterPivotHierarchies` 添加一个附加层次结构以筛选每个数据行。 从数据透视表及其摘要中排除包含项目的任何行。 由于这些筛选器基于项，因此它们只处理离散值。 如果“分类”是示例中的筛选器层次结构，则用户可以选择筛选器的“有机”和“常规”值。 同样，如果选择“箱销售批发”，筛选器选项将是单个数字，如 120 和 150，而不是数字范围。

`FilterPivotHierarchies` 创建时已选中所有值。 这意味着，在用户手动与筛选器控件交互或 `PivotManualFilter` 在属于 `FilterPivotHierarchy`筛选器的字段上设置筛选控件之前，不会筛选任何内容。

以下代码片段将“分类”添加为筛选器层次结构。

```typescript
  farmPivot.addFilterHierarchy(farmPivot.getHierarchy("Classification"));
```

:::image type="content" source="../images/pivottable-filter-hierarchy.png" alt-text="使用数据透视表的“分类”的筛选器控件。":::

### <a name="pivotfilters"></a>PivotFilters

该 `PivotFilters` 对象是应用于单个字段的筛选器集合。 由于每个层次结构只有一个字段，因此在 `PivotHierarchy.getFields()` 应用筛选器时应始终使用第一个字段。 有四种筛选器类型。

- **日期筛选器**：基于日历日期的筛选。
- **标签筛选器**：文本比较筛选。
- **手动筛选**：自定义输入筛选。
- **值筛选器**：数字比较筛选。 这会将关联层次结构中的项与指定数据层次结构中的值进行比较。

通常，在四种类型的筛选器中，只有一种创建并应用到该字段。 如果脚本尝试使用不兼容的筛选器，则会引发错误，其文本为“参数无效或缺失或格式不正确”。

以下代码片段添加了两个筛选器。 第一种是手动筛选器，用于选择现有“分类”筛选器层次结构中的项。 第二个筛选器删除了“批发销售箱”少于 300 个的任何农场。 请注意，这会筛选出这些场的“Sum”，而不是原始数据中的单行。

```typescript
  const classificationField = farmPivot.getFilterHierarchy("Classification").getFields()[0];
  classificationField.applyFilter({
    manualFilter: { 
      selectedItems: ["Organic"] /* The included items. */
    }
  });

  const farmField = farmPivot.getHierarchy("Farm").getFields()[0];
  farmField.applyFilter({
    valueFilter: {
      condition: ExcelScript.ValueFilterCondition.greaterThan, /* The relationship of the value to the comparator. */
      comparator: 300, /* The value to which items are compared. */
      value: "Sum of Crates Sold Wholesale" /* The name of the data hierarchy. Note the "Sum of" prefix. */
      }
  });
```

:::image type="content" source="../images/pivottable-filters.png" alt-text="应用值筛选器和手动筛选器后的数据透视表。":::

### <a name="slicers"></a>切片器

[切片器](https://support.microsoft.com/office/249f966b-a9d5-4b0f-b31a-12651785d29d) 筛选数据透视表 (或标准表) 中的数据。 它们是工作表中的可移动对象，允许快速筛选选择。 切片器的操作方式与手动筛选器 `PivotFilterHierarchy`类似。 要从数据透视表中 `PivotField` 添加或排除这些项的项。

以下代码片段为“Type”字段添加切片器。 它将所选项设置为“Lemon”和“Lime”，然后将切片器向左移动 400 像素。

```typescript
  const fruitSlicer = pivotSheet.addSlicer(
    farmPivot, /* The table or PivotTale to be sliced. */
    farmPivot.getHierarchy("Type").getFields()[0] /* What source to use as the slicer options. */
  );
  fruitSlicer.selectItems(["Lemon", "Lime"]);
  fruitSlicer.setLeft(400);
```

:::image type="content" source="../images/slicer.png" alt-text="在数据透视表上筛选数据的切片器。":::

## <a name="see-also"></a>另请参阅

- [Excel 网页版中 Office 脚本的脚本基础知识](scripting-fundamentals.md)
- [Office 脚本 API 参考](/javascript/api/office-scripts/overview)
