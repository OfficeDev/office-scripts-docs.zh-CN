---
title: 从Teams数据发送Excel会议
description: 了解如何使用 Office 脚本从Teams发送Excel会议。
ms.date: 04/28/2021
localization_priority: Normal
ms.openlocfilehash: b0a3d5732727fd399fe34f3645336840ba4c156d
ms.sourcegitcommit: f7a7aebfb687f2a35dbed07ed62ff352a114525a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/06/2021
ms.locfileid: "52232730"
---
# <a name="send-teams-meeting-from-excel-data"></a><span data-ttu-id="a74be-103">从Teams数据发送Excel会议</span><span class="sxs-lookup"><span data-stu-id="a74be-103">Send Teams meeting from Excel data</span></span>

<span data-ttu-id="a74be-104">此解决方案演示如何使用 Office 脚本和 Power Automate 操作从 Excel 文件选择行，并使用它发送 Teams 会议邀请，然后更新Excel。</span><span class="sxs-lookup"><span data-stu-id="a74be-104">This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="a74be-105">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="a74be-105">Example scenario</span></span>

* <span data-ttu-id="a74be-106">人力资源招聘人员管理应聘者在一个职位Excel计划。</span><span class="sxs-lookup"><span data-stu-id="a74be-106">An HR recruiter manages the interview schedule of candidates in an Excel file.</span></span>
* <span data-ttu-id="a74be-107">招聘人员需要将Teams会议邀请发送给候选人和面试者。</span><span class="sxs-lookup"><span data-stu-id="a74be-107">The recruiter needs to send the Teams meeting invite to the candidate and interviewers.</span></span> <span data-ttu-id="a74be-108">业务规则包括：</span><span class="sxs-lookup"><span data-stu-id="a74be-108">The business rules are to select:</span></span>

    <span data-ttu-id="a74be-109"> () 邀请仅发送给未在文件列中记录的邀请的发送者。</span><span class="sxs-lookup"><span data-stu-id="a74be-109">(a) Invites to only those for whom the invite isn't already sent as recorded in the file column.</span></span>

    <span data-ttu-id="a74be-110"> (b) 将来的面试日期 (任何过去的日期) 。</span><span class="sxs-lookup"><span data-stu-id="a74be-110">(b) Interview dates in the future (no past dates).</span></span>

* <span data-ttu-id="a74be-111">招聘人员需要更新Excel文件，确认已针对符合条件的记录Teams所有会议。</span><span class="sxs-lookup"><span data-stu-id="a74be-111">The recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records.</span></span>

<span data-ttu-id="a74be-112">解决方案有 3 个部分：</span><span class="sxs-lookup"><span data-stu-id="a74be-112">The solution has 3 parts:</span></span>

1. <span data-ttu-id="a74be-113">Office用于根据条件从表中提取数据并作为 JSON 数据返回对象数组的脚本。</span><span class="sxs-lookup"><span data-stu-id="a74be-113">Office Script to extract data from a table based on conditions and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="a74be-114">然后，数据将发送到Teams **创建Teams会议操作** 以发送邀请。</span><span class="sxs-lookup"><span data-stu-id="a74be-114">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span> <span data-ttu-id="a74be-115">在 JSON Teams每个实例发送一个会议。</span><span class="sxs-lookup"><span data-stu-id="a74be-115">Send one Teams meeting per instance in the JSON array.</span></span>
1. <span data-ttu-id="a74be-116">将相同的 JSON 数据发送到另一Office脚本以更新邀请的状态。</span><span class="sxs-lookup"><span data-stu-id="a74be-116">Send the same JSON data to another Office Script to update the status of the invitation.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="a74be-117">示例Excel文件</span><span class="sxs-lookup"><span data-stu-id="a74be-117">Sample Excel file</span></span>

<span data-ttu-id="a74be-118">下载此 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> 中使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="a74be-118">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span>

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a><span data-ttu-id="a74be-119">示例代码：从表中选择筛选的行作为 JSON</span><span class="sxs-lookup"><span data-stu-id="a74be-119">Sample code: Select filtered rows from table as JSON</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook): InterviewInvite[] {
  console.log("Current date time: " + new Date().toUTCString())
  const MEETING_DURATION = workbook.getNamedItem('MeetingDuration').getRange().getValue() as number;
  const sheet = workbook.getWorksheet('Interviews');
  const table = sheet.getTables()[0];
  const dataRows: string[][] = table.getRange().getTexts();
  // OR use the following statement if there's no table:
  // let dataRows = sheet.getUsedRange().getValues();
  const selectedRows = dataRows.filter((row, i) => {
    // Select header row and any data row with the status column equal to approach value.
    return (row[1] === 'FALSE' || i === 0)
  })
  const recordDetails: RecordDetail[] = returnObjectFromValues(selectedRows as string[][]);
  const inviteRecords = generateInterviewRecords(recordDetails, MEETING_DURATION);
  console.log(JSON.stringify(inviteRecords));
  return inviteRecords;
}

/**
 * This helper function converts table values into an object array.
 */
function returnObjectFromValues(values: string[][]): RecordDetail[] {
  let objArray: BasicObj[] = [];
  let objKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objKeys = values[i]
      continue;
    }
    let obj = {}
    for (let j = 0; j < values[i].length; j++) {
      obj[objKeys[j]] = values[i][j]
    }
    objArray.push(obj);
  }
  return objArray as RecordDetail[];
}

/**
 * Generate interview records by selecting required columns.
 * @param records Input records
 * @param mins Number of minutes to add to the start date-time
 */
function generateInterviewRecords(records: RecordDetail[], mins: number): InterviewInvite[] {
  const interviewInvites: InterviewInvite[] = []

  records.forEach((record) => {
    // Interviewer 1
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time1'])) > new Date()) {
      console.log("selected " + new Date(record['Start time1']).toUTCString());
      let startTime = new Date(record['Start time1']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time1']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer1,
        InterviewerEmail: record['Interviewer1 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time1']).toUTCString()))
    }
    // Interviewer 2 
    // If the start date-time is greater than current date-time, add to output records.
    if ((new Date(record['Start time2'])) > new Date()) {
      console.log("selected " + new Date(record['Start time2']).toUTCString());


      let startTime = new Date(record['Start time2']).toISOString();
      // Compute the finish time of the meeting.
      let finishTime = addMins(new Date(record['Start time2']), mins).toISOString();
      interviewInvites.push({
        ID: record.ID,
        Candidate: record.Candidate,
        CandidateEmail: record['Candidate email'] as string,
        CandidateContact: record['Candidate contact'] as string,
        Interviewer: record.Interviewer2,
        InterviewerEmail: record['Interviewer2 email'],
        StartTime: startTime,
        FinishTime: finishTime
      })
    } else {
      console.log("Rejected " + (new Date(record['Start time2']).toUTCString()))

    }
  })
  return interviewInvites;
}

/**
 * Add minutes to start date-time.
 * @param startDateTime Start date-time
 * @param mins Minutes to add to the start date-time
 */
function addMins(startDateTime: Date, mins: number) {
  return new Date(startDateTime.getTime() + mins * 60 * 1000);
}

// Basic key-value pair object.
interface BasicObj {
  [key: string]: string | number | boolean
}

// Input record that matches the table data.
interface RecordDetail extends BasicObj {
  ID: string
  'Invite to interview': string
  Candidate: string
  'Candidate email': string
  'Candidate contact': string
  Interviewer1: string
  'Interviewer1 email': string
  Interviewer2: string
  'Interviewer2 email': string
  'Start time1': string
  'Start time2': string
}

// Output record.
interface InterviewInvite extends BasicObj {
  ID: string
  Candidate: string
  CandidateEmail: string
  CandidateContact: string
  Interviewer: string
  InterviewerEmail: string
  StartTime: string
  FinishTime: string
}
```

## <a name="sample-code-mark-as-invited"></a><span data-ttu-id="a74be-120">示例代码：标记为受邀</span><span class="sxs-lookup"><span data-stu-id="a74be-120">Sample code: Mark as invited</span></span>

```TypeScript
function main(workbook: ExcelScript.Workbook, completedInvitesString: string) {
    completedInvitesString = `[
  {
    "ID": "10",
    "Candidate": "Adele ",
    "CandidateEmail": "AdeleV@M365x904181.OnMicrosoft.com",
    "CandidateContact": "1234567899",
    "Interviewer": "Megan",
    "InterviewerEmail": "MeganB@M365x904181.OnMicrosoft.com",
    "StartTime": "2020-11-03T18:30:00Z",
    "FinishTime": "2020-11-03T22:45:00Z"
  },
  {
    "ID": "30",
    "Candidate": "Allan ",
    "CandidateEmail": "AllanD@M365x904181.OnMicrosoft.com",
    "CandidateContact": "1234567978",
    "Interviewer": "Raul",
    "InterviewerEmail": "RaulR@M365x904181.OnMicrosoft.com",
    "StartTime": "2020-11-03T23:00:00Z",
    "FinishTime": "2020-11-03T23:45:00Z"
  }
]`;
    let completedInvites = JSON.parse(completedInvitesString) as InterviewInvite[];
    const sheet = workbook.getWorksheet('Interviews');
    const range = sheet.getTables()[0].getRange();
    const dataRows = range.getValues();
    for (let i=0; i < dataRows.length; i++) {
        for (let invite of completedInvites) {
            if (String(dataRows[i][0]) === invite.ID) {
                range.getCell(i,1).setValue(true);
            }
        }
    }
    return;
}


// Invite record.
interface InterviewInvite  {
    ID: string
    Candidate: string
    CandidateEmail: string
    CandidateContact: string
    Interviewer: string
    InterviewerEmail: string
    StartTime: string
    FinishTime: string
}
```

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="a74be-121">培训视频：从Teams发送Excel会议</span><span class="sxs-lookup"><span data-stu-id="a74be-121">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="a74be-122">[观看 Sudhi Ramamurthy 在 YouTube 上演练此示例](https://youtu.be/HyBdx52NOE8)。</span><span class="sxs-lookup"><span data-stu-id="a74be-122">[Watch Sudhi Ramamurthy walk through this sample on YouTube](https://youtu.be/HyBdx52NOE8).</span></span>
