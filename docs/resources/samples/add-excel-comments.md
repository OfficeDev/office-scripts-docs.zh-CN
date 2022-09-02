---
title: 在 Excel 中添加注释
description: 了解如何使用 Office 脚本在工作表中添加注释。
ms.date: 06/29/2021
ms.localizationpriority: medium
ms.openlocfilehash: 90f072805e6798a4f9d6e74889ccca15610c87bd
ms.sourcegitcommit: a6504f8b0d6b717457c6e0b5306c35ad3900914e
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/02/2022
ms.locfileid: "67572491"
---
# <a name="add-comments-in-excel"></a>在 Excel 中添加注释

此示例演示如何向单元格添加注释，包括 [@mentioning](https://support.microsoft.com/office/90701709-5dc1-41c7-aa48-b01d4a46e8c7) 同事。

## <a name="example-scenario"></a>示例方案

* 团队主管维护轮班计划。 团队主管将员工 ID 分配给轮班记录。
* 团队主管希望通知员工。 通过添加@mentions员工的注释，将向员工发送来自工作表的自定义消息。
* 随后，员工可以在方便时查看工作簿并响应批注。

## <a name="solution"></a>解决方案

1. 该脚本从员工工作表中提取员工信息。
1. 然后，脚本将注释添加 (包括相关员工电子邮件) 到班次记录中的相应单元格。
1. 在添加新注释之前，将删除单元格中的现有注释。

## <a name="sample-excel-file"></a>示例 Excel 文件

下载现成工作簿 [ 的excel-comments.xlsx](excel-comments.xlsx) 。 添加以下脚本以自行尝试示例！

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

[观看苏迪 · 拉马穆尔西在 YouTube 上浏览这个示例](https://youtu.be/CpR78nkaOFw)。
