---
title: 在外接程序中添加Excel
description: 了解如何使用Office脚本在工作表中添加注释。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3ff9d56934520a98dd1de7d31077396294bde29d
ms.sourcegitcommit: d3ed4bdeeba805d97c930394e172e8306a0cf484
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/15/2021
ms.locfileid: "59332395"
---
# <a name="add-comments-in-excel"></a>在外接程序中添加Excel

此示例演示如何向单元格添加注释，包括 [@mentioning添加注释](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 。

## <a name="example-scenario"></a>示例应用场景

* 团队主管负责维护轮班计划。 团队主管向轮班记录分配员工 ID。
* 团队主管希望通知员工。 通过向员工添加@mentions注释，员工将通过电子邮件从工作表收到自定义邮件。
* 随后，员工可以在方便时查看工作簿并回复注释。

## <a name="solution"></a>解决方案

1. 该脚本从员工工作表中提取员工信息。
1. 然后，该脚本添加注释 (，包括相关的员工电子邮件) 到班次记录中的相应单元格。
1. 添加新注释之前，将删除单元格中的现有注释。

## <a name="sample-excel-file"></a>示例Excel文件

下载 <a href="excel-comments.xlsx">excel-comments.xlsx</a> 工作簿的工作簿。 添加以下脚本以自己试用示例！

## <a name="sample-code-add-comments"></a>示例代码：添加注释

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Get the list of employees.
  const employees = workbook.getWorksheet('Employees').getUsedRange().getTexts();
  console.log(employees); 
  
  // Get the schedule information from the schedule table.
  const scheduleSheet = workbook.getWorksheet('Schedule');
  const table = scheduleSheet.getTables()[0];
  const range = table.getRangeBetweenHeaderAndTotal();
  const scheduleData = range.getTexts();

  // Look through the schedule for a matching employee.
  for (let i = 0; i < scheduleData.length; i++) {
    let employeeId = scheduleData[i][3];

    // Compare the employee ID in the schedule against the employee information table.
    let employeeInfo = employees.find(employeeRow => employeeRow[0] === employeeId);
    if (employeeInfo) {
      console.log("Found a match " + employeeInfo);
      let adminNotes = scheduleData[i][4];

      // Look for and delete old comments, so we avoid conflicts.
      let comment = workbook.getCommentByCell(range.getCell(i, 5));
      if (comment) {
        comment.delete();
      }

      // Add a comment using the admin notes as the text.
      workbook.addComment(range.getCell(i,5), {
        mentions: [{
          email: employeeInfo[1],
          id: 0, // This ID maps this mention to the `id=0` text in the comment.
          name: employeeInfo[2]
        }],
        richContent: `<at id=\"0\">${employeeInfo[2]}</at> ${adminNotes}`
      }, ExcelScript.ContentType.mention);        
      
    } else {
      console.log("No match for: " + employeeId);
    }
  }
}
```

## <a name="training-video-add-comments"></a>培训视频：添加注释

[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/CpR78nkaOFw)。
