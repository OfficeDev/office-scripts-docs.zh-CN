---
title: 安排在Teams
description: 了解如何使用 Office 脚本从Teams发送Excel会议。
ms.date: 05/25/2021
localization_priority: Normal
ms.openlocfilehash: f93d9ceca6603ddb9e7123a393787fcf54597cca
ms.sourcegitcommit: 339ecbb9914d54f919e3475018888fb5d00abe89
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/28/2021
ms.locfileid: "52697772"
---
# <a name="office-scripts-sample-scenario-schedule-interviews-in-teams"></a>Office脚本示例方案：安排在 Teams

在此方案中，你是一名 HR 招聘人员，负责安排与Teams。 在一个管理文件中管理应聘者的Excel计划。 你需要向候选人和Teams发送会议邀请。 然后，你需要更新Excel文件，并确认Teams会议已发送。

解决方案有三个步骤组合在单个流Power Automate流。

1. 脚本从表中提取数据，并返回对象数组作为 JSON 数据。
1. 然后，数据将发送到Teams **创建Teams会议操作** 以发送邀请。
1. 相同的 JSON 数据将发送到另一个脚本以更新邀请的状态。

## <a name="scripting-skills-covered"></a>涵盖的脚本编写技能

* Power Automate流
* Teams集成
* 表分析

## <a name="sample-excel-file"></a>示例Excel文件

下载此 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> 中使用的文件，然后自己试用！ 请务必更改至少一个电子邮件地址，以便收到邀请。

## <a name="sample-code-extract-table-data-to-schedule-invites"></a>示例代码：提取表数据以计划邀请

为此流 **命名此脚本 Schedule Interviews。**

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  const MEETING_DURATION = workbook.getWorksheet("Constants").getRange("B1").getValue() as number;
  const MESSAGE_TEMPLATE = workbook.getWorksheet("Constants").getRange("B2").getValue() as string;

  // Get the interview candidate information.
  const sheet = workbook.getWorksheet("Interviews");
  const table = sheet.getTables()[0];
  const dataRows = table.getRangeBetweenHeaderAndTotal().getValues();

  // Convert the table rows into InterviewInvite objects for the flow.
  let invites: InterviewInvite[] = [];
  dataRows.forEach((row) => {
    const inviteSent = row[1] as boolean;
    if (!inviteSent) {
      const startTime = new Date(Math.round(((row[6] as number) - 25569) * 86400 * 1000));
      const finishTime = new Date(startTime.getTime() + MEETING_DURATION * 60 * 1000);
      const candidateName = row[2] as string;
      const interviewerName = row[4] as string;

      invites.push({
        ID: row[0] as string,
        Candidate: candidateName,
        CandidateEmail: row[3] as string,
        Interviewer: row[4] as string,
        InterviewerEmail: row[5] as string,
        StartTime: startTime.toISOString(),
        FinishTime: finishTime.toISOString(),
        Message: generateInviteMessage(MESSAGE_TEMPLATE, candidateName, interviewerName)
      });
    }    
  });

  console.log(JSON.stringify(invites));
  return invites;
}

function generateInviteMessage(
  messageTemplate: string,
   candidate: string,
   interviewer: string) : string {
  return messageTemplate.replace("_Candidate_", candidate).replace("_Interviewer_", interviewer);
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-code-mark-rows-as-invited"></a>示例代码：将行标记为受邀

将此脚本 **命名 为流记录已** 发送的邀请。

```TypeScript
function main(workbook: ExcelScript.Workbook, invites: InterviewInvite[]) {
  const table = workbook.getWorksheet("Interviews").getTables()[0];

  // Get the ID and Invite Sent columns from the table.
  const idColumn = table.getColumnByName("ID");
  const idRange = idColumn.getRangeBetweenHeaderAndTotal().getValues();
  const inviteSentColumn = table.getColumnByName("Invite Sent?");

  const dataRowCount = idRange.length;

  // Find matching IDs to mark the correct row.
  for (let row = 0; row < dataRowCount; row++){
    let inviteSent = invites.find((invite) => {
      return invite.ID == idRange[row][0] as string;
    });

    if (inviteSent) {
      inviteSentColumn.getRangeBetweenHeaderAndTotal().getCell(row, 0).setValue(true);
      console.log(`Invite for ${inviteSent.Candidate} has been sent.`);
    }
  } 
}

// The interview invite information.
interface InterviewInvite {
  ID: string
  Candidate: string
  CandidateEmail: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
  Message: string
}
```

## <a name="sample-flow-run-the-interview-scheduling-scripts-and-send-the-teams-meetings"></a>示例流程：运行访谈计划脚本并发送Teams会议

1. 创建新的即时 **云流**。
1. 选择 **"手动触发流"，** 然后按"**创建"。**
1. 添加使用 **Excel** Online (**Business) 连接器** 和 **"运行** 脚本"操作的新步骤。 使用下列值完成连接器。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件***：hr-interviews.xlsx (浏览器选项选择)*
    1. **脚本**：计划面试 已完成的 Excel Online (Business) 连接器的屏幕截图，用于从 Power Automate 中的工作簿获取 :::image type="content" source="../../images/schedule-interviews-1.png" alt-text="面试Power Automate":::
1. 添加一 **个使用**"创建会议 **Teams的新** 步骤。 当你从连接器选择动态内容Excel，将针对你的流生成应用到每个块。 使用下列值完成连接器。
    1. **日历 ID**：日历
    1. **主题**：Contoso Interview
    1. **邮件****： (Excel** 值) 
    1. **时区：** 太平洋标准时间
    1. **开始时间****：StartTime** (Excel值) 
    1. **结束时间****：FinishTime** (Excel值) 
    1. **必选与会者**： **CandidateEmail** ;**ScreenshotEmail** (Excel值) 已完成的 Teams :::image type="content" source="../../images/schedule-interviews-2.png" alt-text="连接器的屏幕截图，用于":::安排Power Automate
1. 在同一 **个"应用到每个** 块"中，使用"运行脚本Excel添加 (**Business)** 连接器。  使用以下值。
    1. **位置**：OneDrive for Business
    1. **文档库**：OneDrive
    1. **文件***：hr-interviews.xlsx (浏览器选项选择)*
    1. **脚本**：记录已发送的邀请
    1. **邀请**：**结果** (Excel值) 已完成的 Excel Online :::image type="content" source="../../images/schedule-interviews-3.png" alt-text=" (Business)":::连接器的屏幕截图，用于录制已Power Automate
1. 保存流并试用。

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a>培训视频：从Teams发送Excel会议

[观看 Sudhi Ramamurthy 在 YouTube](https://youtu.be/HyBdx52NOE8)上演练此示例的版本。 他的版本使用更强大的脚本来处理更改的列和过时的会议时间。
