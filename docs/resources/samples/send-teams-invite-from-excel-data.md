---
title: 从 Excel 数据发送 Teams 会议
description: 了解如何使用 Office 脚本从 Excel 数据发送 Teams 会议。
ms.date: 03/30/2021
localization_priority: Normal
ms.openlocfilehash: 807c9228049504c089c8dafe63a5d9ccaab94399
ms.sourcegitcommit: 5d24e77df70aa2c1c982275d53213c2a9323ff86
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/02/2021
ms.locfileid: "51571286"
---
# <a name="send-teams-meeting-from-excel-data"></a><span data-ttu-id="bdfbf-103">从 Excel 数据发送 Teams 会议</span><span class="sxs-lookup"><span data-stu-id="bdfbf-103">Send Teams meeting from Excel data</span></span>

<span data-ttu-id="bdfbf-104">此解决方案演示如何使用 Office 脚本和 Power Automate 操作从 Excel 文件选择行，并使用它发送 Teams 会议邀请，然后更新 Excel。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-104">This solution shows how to use Office Scripts and Power Automate actions to select rows from Excel file and use it to send a Teams meeting invite then update Excel.</span></span>

## <a name="example-scenario"></a><span data-ttu-id="bdfbf-105">示例应用场景</span><span class="sxs-lookup"><span data-stu-id="bdfbf-105">Example scenario</span></span>

* <span data-ttu-id="bdfbf-106">HR 招聘人员管理 Excel 文件中候选人的面试计划。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-106">An HR recruiter manages the interview schedule of candidates in an Excel file.</span></span>
* <span data-ttu-id="bdfbf-107">招聘人员需要向候选人和面试者发送 Teams 会议邀请。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-107">The recruiter needs to send the Teams meeting invite to the candidate and interviewers.</span></span> <span data-ttu-id="bdfbf-108">业务规则包括：</span><span class="sxs-lookup"><span data-stu-id="bdfbf-108">The business rules are to select:</span></span>

    <span data-ttu-id="bdfbf-109"> () 邀请仅发送给未在文件列中记录的邀请的发送者。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-109">(a) Invites to only those for whom the invite isn't already sent as recorded in the file column.</span></span>

    <span data-ttu-id="bdfbf-110"> (b) 将来的面试日期 (任何过去的日期) 。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-110">(b) Interview dates in the future (no past dates).</span></span>

* <span data-ttu-id="bdfbf-111">招聘人员需要更新 Excel 文件，并确认已针对符合条件的记录发送了所有 Teams 会议。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-111">The recruiter needs to update the Excel file with the confirmation that all Teams meetings have been sent for the eligible records.</span></span>

<span data-ttu-id="bdfbf-112">解决方案有 3 个部分：</span><span class="sxs-lookup"><span data-stu-id="bdfbf-112">The solution has 3 parts:</span></span>

1. <span data-ttu-id="bdfbf-113">用于根据条件从表中提取数据并作为 JSON 数据返回对象数组的 Office 脚本。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-113">Office Script to extract data from a table based on conditions and returns an array of objects as JSON data.</span></span>
1. <span data-ttu-id="bdfbf-114">然后，将数据发送到 Teams **创建 Teams 会议** 操作以发送邀请。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-114">The data is then sent to the Teams **Create a Teams meeting** action to send invites.</span></span> <span data-ttu-id="bdfbf-115">JSON 数组中每个实例发送一个 Teams 会议。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-115">Send one Teams meeting per instance in the JSON array.</span></span>
1. <span data-ttu-id="bdfbf-116">将相同的 JSON 数据发送到另一个 Office 脚本以更新邀请的状态。</span><span class="sxs-lookup"><span data-stu-id="bdfbf-116">Send the same JSON data to another Office Script to update the status of the invitation.</span></span>

## <a name="sample-excel-file"></a><span data-ttu-id="bdfbf-117">示例 Excel 文件</span><span class="sxs-lookup"><span data-stu-id="bdfbf-117">Sample Excel file</span></span>

<span data-ttu-id="bdfbf-118">下载此 <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> 中使用的文件，然后自己试用！</span><span class="sxs-lookup"><span data-stu-id="bdfbf-118">Download the file <a href="hr-schedule.xlsx">hr-schedule.xlsx</a> used in this solution and try it out yourself!</span></span>

## <a name="sample-code-select-filtered-rows-from-table-as-json"></a><span data-ttu-id="bdfbf-119">示例代码：从表中选择筛选的行作为 JSON</span><span class="sxs-lookup"><span data-stu-id="bdfbf-119">Sample code: Select filtered rows from table as JSON</span></span>

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

## <a name="sample-code-mark-as-invited"></a><span data-ttu-id="bdfbf-120">示例代码：标记为受邀</span><span class="sxs-lookup"><span data-stu-id="bdfbf-120">Sample code: Mark as invited</span></span>

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

## <a name="training-video-send-a-teams-meeting-from-excel-data"></a><span data-ttu-id="bdfbf-121">培训视频：从 Excel 数据发送 Teams 会议</span><span class="sxs-lookup"><span data-stu-id="bdfbf-121">Training video: Send a Teams meeting from Excel data</span></span>

<span data-ttu-id="bdfbf-122">[![观看分步视频，了解如何从 Excel 数据发送 Teams 会议](../../images/teams-invite-vid.jpg)](https://youtu.be/HyBdx52NOE8 "如何从 Excel 数据发送 Teams 会议分步视频")</span><span class="sxs-lookup"><span data-stu-id="bdfbf-122">[![Watch step-by-step video on how to send a Teams meeting from Excel data](../../images/teams-invite-vid.jpg)](https://youtu.be/HyBdx52NOE8 "Step-by-step video on how to send a Teams meeting from Excel data")</span></span>
