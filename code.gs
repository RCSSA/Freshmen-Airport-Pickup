// faq
// 1. what if a student wants to change his/her time?
// if they have their time slot picked by some volunteer, first step is to contact the volunteer
// if the new time no longer works for the volunteer, submit another response should be fine
// if their times slots have not been picked by a volunteer, please go to the response sheet and
// manually remove the response, then the student can submit another response (future patches can potentially automate this process)

// 2. timezone specifications
// Javascript has well known date time issues
// because this event is meant to happen during the summer
// all responses in the student sheet will be recorded in Central Daylight Timezone (GMT -5, which is the same as EST)
// to change this, go to the sheet, click on File>>Settings>>Timezone


// features to be upgrade
// 1. in find slots: search for date can be improved to binary search
// 2. add the function so that volunteer can chooose how many students to pick up
// 3. remove weeks/days that have no students waiting to be picked up (reinit everyday?)
// 4. in addQuestions: don't add dates when no students are waiting to be picked up
// 5. When a new student added a time slot after the form is generated, make the new choice pop up on the form automatically
// 6. timezone issue: convert everthing to local?
// 7. generate info questions using code; set them as required
// 8. num of people that one can pick up: more than 4 people, contact RCSSA manually

// global constants
const formId = "1BK27EKKNY35bThmGf0JyxJO77NekqRfZ1PfH2ujETRo";
const form = FormApp.openById(formId);

// spread sheets
const studentSSId = "1xpvPnj1l3KI9pCac-usnUunV4u5cUXBguhkELRSiUMc";
const studentSS = SpreadsheetApp.openById(studentSSId);
const volSSId = "15ZnS8pw-AMXWOTL3NpgItrrMxMt8-ZnD__4AQMxz13k";
const volSS = SpreadsheetApp.openById(volSSId);

// sheets
const studentSheetName = "Form Responses 1";
const studentSheet = studentSS.getSheetByName(studentSheetName);
const volSheetName = "Form Responses 1";
const volSheet = volSS.getSheetByName(volSheetName);

// lock service
const wait_time = 20000; // 20 seconds
const lock = LockService.getScriptLock();

// time related
const timeZone = "est"  //note that est is the same as cdt, it is used here because the system doesn't support cdt 
const startWeek = new Date(`2022-08-14 ${timeZone}`);
const endWeek = new Date(`2022-08-14 ${timeZone}`);
const options = { weekday: 'short', year: 'numeric', month: 'short', day: 'numeric', timeZone: timeZone };

// new student sheet related - col number starts from 1
const studentEmailColNum = 2;
const studentNameColNum = 3;
const studentWechatColNum = 4;
const studentOtherContactColNum = 5;
const timeColNum = 6;  //  col number for responses containing student arrival date and time

const volNameColNum = 8;
const volWechatColNum = 9;
const volContactColNum = 10;
const volEmailColNum = 11;
const volVeriColNum = 12;
const emailSentColNum = 13;

// volunteer form related
const questionPrefix = {"week": "Step1", "date": "Step2", "time": "Step3"}
const nullChoiceStr = "暂时没有合适的时间，我愿意先加入志愿者群~"
const nameQIdx = 0;
const wechatQIdx = 1
const otherContactsQIdx = 2;
const numStudentQIdx = 3;
const firstMultChoiceQIdx = 4;

// format

// status
const UNASSIGNED = '';
const UNVERIFIED = 1;
const VERIFIED = 2;

// volunteer group QRcode
const qrCodeFileId = "1BCo66b-O-ablSuihrBZU1hpc1Kb-woso"
const qrCodeBlob = DriveApp.getFileById(qrCodeFileId).getAs('image/jpeg')

function getQuota(){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}

// routine 
function routine(){
  clean()
  addMultipleChoiceQuestions(); 
  // console.log(...date2Q.keys())
  // _updateForm(_getLastResponse())
}

function bind(){
    ScriptApp.newTrigger('myOnFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}


function clean(n=5){
  var items = form.getItems();
  var len = items.length
  console.log(`start cleaning from item no.${n}`);
  for (var i=n; i<len; i++) {
    console.log(`removing item: ${items[i].getTitle()}`);
    form.deleteItem(n);
  }
  console.log("cleaning done");
}

function addMultipleChoiceQuestions(){
  // create a question to select week
  // let weekSection = form.addPageBreakItem()
  //   .setTitle('接机周')
  //   .setHelpText('如果一周没有合适的时间话 可以看看其他周哦');
  let weekQ = form.addMultipleChoiceItem()
  let weekChoices = []; 
  let date = startWeek;
  let endDate, section;
  while (date <= endWeek){
    endDate = _addDays(date, 6);

    // add a section for each week
    section = form.addPageBreakItem()
    .setTitle(`${date.toLocaleDateString('en-US', options)} - ${endDate.toLocaleDateString('en-US', options)}周接机日期`);

    // add the select date question
    _addDayQ(date, endDate);

    // add a choice for each week, point to the section
    weekChoices.push(weekQ.createChoice(date.toLocaleDateString('en-US', options) + " - " +endDate.toLocaleDateString('en-US', options), section))

    // loop updates
    date = _addDays(date, 7);    
  }

  // fill in choices for select week questions
  weekQ.setChoices(weekChoices);
  weekQ.setTitle(`${questionPrefix.week}: 请选择接机周`).setRequired(true);
}


function _addDayQ(startDate, endDate){
  // create a question to select the date
  let dateQ = form.addMultipleChoiceItem()
  // let dateQ = form.addMultipleChoiceItem().setTitle("请选择接机日期");
  let dateChoices = [];
  let date = startDate;
  let section;
  while (date <= endDate){
    // add a section for each day
    section = form.addPageBreakItem()
    .setTitle(`${date.toLocaleDateString('en-US', options)} 接机时间`)
    .setHelpText('如果当前没有您想要的时间，您可以先提交问卷，我们会将接机志愿者+新生群的二维码通过邮件发给您，方便您获取最新的动态');

    // add a choice for each day, point to the section
    dateChoices.push(dateQ.createChoice(date.toLocaleDateString('en-US', options), section))
    console.log(`adding question for ${date.toLocaleDateString('en-US', options)}`)
    // add a select time question for each day
    _addTimeQ(date);
    
    // loop updates
    date = _addDays(date, 1);    
  }

  // fill in choices for select week questions
  dateQ.setChoices(dateChoices);
  dateQ.setTitle(`${questionPrefix.date}: 请选择接机日期`).setRequired(true);
}


function _addTimeQ(date){
  let timeSlots = _findTimes(date);
  let dateOptions = { month: 'numeric', day: 'numeric', timeZone: timeZone };
  // create a question to select the time
  let timeQ = form.addMultipleChoiceItem()
  let timeChoices = []
  timeSlots.map(slot => timeChoices.push(timeQ.createChoice(slot, FormApp.PageNavigationType.SUBMIT)));
  // a reminder choice
  timeChoices.push(timeQ.createChoice(nullChoiceStr, FormApp.PageNavigationType.SUBMIT));
  timeQ.setChoices(timeChoices);
  timeQ.setTitle(`${questionPrefix.time}: 请选择接机时间`).setRequired(true);
  // date2Q.set(date.toLocaleDateString([], dateOptions), timeQ);
}


//toDo: improve the complexity of this algorithm
// linear search -> binary search
// note that the argument "date" is in CST
function _findTimes(date){
  let times = _getColData(timeColNum, studentSheet);
  let volNames = _getColData(volNameColNum, studentSheet);
  let unpickedTimes = [];
  for (let i=0; i<times.length; i++){
    if (_isSameDate(new Date(times[i]), date) && volNames[i] == ''){
      unpickedTimes.push(times[i])
    }
  }
  let timeOptions =  { timeZone: timeZone , year: 'numeric', month: 'short', day: 'numeric', hour: "2-digit", minute: "2-digit"}
  let timesString = unpickedTimes.map(dp => dp.toLocaleString('en-US',timeOptions ))
  return [...new Set(timesString)].sort()
}

function _isSameDate(date1, date2){
  let cmp = {timeZone: timeZone, month: "numeric", day: "numeric"};
  let lang = "en-US";
  return date1.toLocaleDateString(lang, cmp) == date2.toLocaleDateString(lang, cmp);
}

function _isSameTime(date1, date2){
  let cmp = { timeZone: timeZone , year: 'numeric', month: 'short', day: 'numeric', hour: "2-digit", minute: "2-digit"}
  let lang = "en-US";
  console.log(date1.toLocaleDateString(lang, cmp) )
  console.log(date2.toLocaleDateString(lang, cmp) )
  return date1.toLocaleDateString(lang, cmp) == date2.toLocaleDateString(lang, cmp);
}



// ------------------- after submission ------------------------------------------//
// Function to be called after each submission
function myOnFormSubmit(event){
  Logger.log("Response from " + event.response.getRespondentEmail() + " is received.")
  lock.tryLock(200000);
  if (!lock.hasLock()) {
    Logger.log('Could not obtain lock after 200 seconds.');
  }else{
    let volEmail = event.response.getRespondentEmail();
    let formRespones = event.response.getItemResponses();
    let volName = formRespones[nameQIdx].getResponse();
    let volWechat = formRespones[wechatQIdx].getResponse();
    let volContact = formRespones[otherContactsQIdx].getResponse();
    let volNumStudent = formRespones[numStudentQIdx].getResponse();
    let volTime;    
    let multipleChoiceItems = formRespones.slice(firstMultChoiceQIdx);
    let timeQuestion;
    let itemResponse;
    let selected = false;
    // send a confirmation email regardless
    sendConfirmationEmail(volEmail);

    // find the question and the response containing the date+time information
    // filter out null choices 暂无合适时间 
    for (itemResponse of multipleChoiceItems){
      timeQuestion = itemResponse.getItem().asMultipleChoiceItem();
      if (timeQuestion.getTitle().startsWith(questionPrefix.time)){
        volTime = itemResponse.getResponse();
          if (!volTime.startsWith(nullChoiceStr)){
              selected = true;
              break;
          }
      }
    }

    if (!selected){
      Logger.log("this volunteer did not select any time slot.")
      return;
    }

    Logger.log("This volunteer selected " + volTime);
    // all these questions are required so there should not be any empty repsonses
    // if (!(_checkName(volName) || _checkContact(volContact) || _checkNumStudent(volNumStudent) || _checkWechat(volWechat))){
    //   _sendFailureEmail(volEmail, volTime);
    //   return
    // }

    result = _matchStudent(volName, volEmail, volWechat, volContact, volTime, volNumStudent);
    Logger.log(result);
    // due to concurrency, some volunteer might select a datetime that is already gone
    if(!result.haveFound){
        _sendFailureEmail(volEmail, volTime);
    }
    // if no more students are waiting to be picked up at that time
    // remove that choice from the form
    else if (!result.haveLeft){
      _updateForm(itemResponse);
    }

    lock.releaseLock();
  }
}

function _matchStudent(volName, volEmail,volWechat, volContact, volTime, volNumStudent){
  console.log("volunteer can pick up " + volNumStudent+ " students");
  volTime = new Date(volTime.toString() + " " + timeZone);
  let studentTimeCol = studentSheet.getRange(2, timeColNum, studentSheet.getLastRow()-1, 1).getValues().flat();
  let curVolNames = studentSheet.getRange(2, volNameColNum, studentSheet.getLastRow()-1, 1).getValues().flat();
  let studentTimeSlots = [];
  let haveFound = false;
  let haveLeft = false;
  studentTimeCol.map(datetime => studentTimeSlots.push(new Date(datetime)))
  // Iterate through all studen times and
  // 1. match all students until volNumStudent is reached
  // 2. check if there are any students left for the same time slot
  for (let i=0; i<studentTimeSlots.length; i++){
    let studentTime = studentTimeSlots[i];
    let curVolName = curVolNames[i];    
    if (volTime.getTime()==studentTime.getTime() && curVolName == UNASSIGNED){
      if (volNumStudent > 0){
        studentSheet.getRange(2+i, volNameColNum).setValue(volName);
        studentSheet.getRange(2+i, volEmailColNum).setValue(volEmail);
        studentSheet.getRange(2+i, volWechatColNum).setValue(volWechat);
        studentSheet.getRange(2+i, volContactColNum).setValue(volContact);
        // if we found one student, still return true to let them exchange  info
        haveFound = true;  
      }

      if (volNumStudent < 0){
        Logger.log("wrong subtraction")
      }    
      if (volNumStudent == 0){
        haveLeft = true;
        break;
      }
      volNumStudent--;
    }
  } 
  return {haveFound: haveFound, haveLeft: haveLeft};
}


function _updateForm(itemResponse){
  // get the question ID corresponding to the date selected
  _removeResponseFromChoices(itemResponse);
}

function sendEmail(){
  Logger.log("sending email");
  let cnt = 5;
  let emailSentCol = _getColData(emailSentColNum, studentSheet);
  let volVeriCol = _getColData(volVeriColNum, studentSheet);
  let volEmailCol = _getColData(volEmailColNum, studentSheet);
  let volNameCol = _getColData(volNameColNum, studentSheet);
  let volWechatCol = _getColData(volWechatColNum, studentSheet);
  let volContactCol = _getColData(volContactColNum, studentSheet)
  let studentNameCol = _getColData(studentNameColNum, studentSheet);
  let studentEmailCol = _getColData(studentEmailColNum, studentSheet);
  let studentTimeCol = _getColData(timeColNum, studentSheet);
  let studentWechatCol = _getColData(studentWechatColNum, studentSheet);
  let studentOtherContactsCol = _getColData(studentOtherContactColNum, studentSheet)

  for (let i=0; i<studentSheet.getLastRow(); i++){
    // exchange email for those who have been verified but not yet received an email
    if (emailSentCol[i] == UNASSIGNED && volVeriCol[i] != UNASSIGNED && volNameCol[i] != UNASSIGNED){
      Logger.log(`exchanging contacts: ${studentEmailCol[i]} ${volEmailCol[i]}`)
      sendSucessEmail(volEmailCol[i], studentNameCol[i], studentTimeCol[i], studentWechatCol[i], studentOtherContactsCol[i], studentEmailCol[i])
      sendSucessEmail(studentEmailCol[i], volNameCol[i], studentTimeCol[i], volWechatCol[i], volContactCol[i], volEmailCol[i])
      studentSheet.getRange(2+i, emailSentColNum).setValue(1);
      cnt -= 1;
      if (cnt == 0){
        return
      }
    }
  }
}

function sendConfirmationEmail(receiver){
  Logger.log(`sending confirmation email to ${receiver}`);
  MailApp.sendEmail({
    to: receiver,
    // multi line string!
    subject: `RCSSA: 2022接机群二维码！`,
    name: "James Li",
    body: 
    `RCSSA志愿者你好！
      再次感谢你参加本次接机活动。你的接机申请已收到！
      如果您已经选择好了接机时间，请稍作等待。出于安全原因考虑, 会有一位RCSSA志愿者人工审核您的信息。审核通过后您和匹配的新生都会再收到一封含有对方联系方式的电子邮件，请多多查收邮箱！
      您可以通过附件内的二维码来加入接机志愿者+新生群，获取最新的动态！    
    祝大家新学期一帆风顺！

    2022-2023 RCSSA全体成员
    `,
    attachments: [qrCodeBlob]
  });
  Logger.log(`Done sending confirmation email to ${receiver}`);
}

function sendSucessEmail(receiver, name, time, wechat, otherContacts, cemail){
  time = new Date(time.toString())
  let dateOption = {month: "numeric", day: "numeric", timeZone: timeZone};
  let timeOption = {week: "short", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit", timeZone: timeZone}
  let date = time.toLocaleDateString('en-US', dateOption)
  MailApp.sendEmail({
    to: receiver,
    // multi line string!
    subject: `RCSSA: 您${date}号的接机已成功匹配`,
    name: "RCSSA",
    body: 
    `RCSSA志愿者/新生你好！
      您${time.toLocaleDateString('en-US', timeOption)}的接机已成功匹配！以下是对方的姓名和联系方式：
      
      姓名： ${name}
      Email: ${cemail}
      微信号：${wechat}
      其他联系方式: ${otherContacts}
      附件内有本次接机活动的群二维码，欢迎加入来获取最新动态！
      在接机过程中可能会出现新生因为海关，行李，天气等原因导致飞机延误，志愿者也可能会因为突发情况无法按时抵达，请大家一定保持联系，互相理解！也祝大家在这个过程中结识更多的新生和学长学姐！
      再次感谢你参加本次接机活动，志愿者请记得领取接机福利，新生也欢迎关注RCSSA公众号: RCSSA 以及官网 rcssa.rice.edu
    
    祝大家新学期一帆风顺！

    20222-2023 RCSSA全体成员
    `,
    attachments: [qrCodeBlob]
  });
}


// remove the first occurrence of a value from an array
// nothing happen if the val does not exist
function removeArrayVal(arr, val){
  let index = arr.indexOf(val);
  if (index !== -1) {
    arr.splice(index, 1);
  }
}

// ------------------------ Error Checking ---------------------------
function _sendFailureEmail(receiver, time){
  Logger.log(`failed to match at ${time} `);
  time = new Date(time.toString());
  let dateOption = {month: "numeric", day: "numeric"};
  let timeOption = {week: "short", month: "short", day: "numeric", hour: "2-digit", minute: "2-digit"}
  let date = time.toLocaleDateString('en-US', dateOption)
  MailApp.sendEmail({
    to: receiver,
    // multi line string!
    subject: `RCSSA: 感谢您${date}号的接机报名`,
    name: "James Li",
    body: 
    `RCSSA志愿者你好！
      很抱歉，您${time.toLocaleDateString('en-US', timeOption)}的接机报名没有成功。
      由于google form本身的更新速度有限，您和其他多个志愿者同时选择了同一个接机时间。
      因为选择这个时间的志愿者的人数大于在这个时间抵达的新生的人数，系统自动根据提交问卷的先后顺序匹配接机报名。
      再次感谢你参加本次接机活动，欢迎你选择其他的时间帮助新生接机！

    祝新学期一帆风顺！

    20222-2023 RCSSA全体成员
    `,
  });
}

function _checkName(name){
  return name != "";
}

function _checkWechat(wechat){
  return wechat != "";
}

function _checkContact(contact){
  return contact != "";
}

function _checkContact(contact){
  return contact != "";
}

function _checkNumStudent(num){
  return num != 0;
}


// --------------------- helper function ----------------------------
function _addDays(date, days) {
  let result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function _removeResponseFromChoices(itemResponse){
  let question = itemResponse.getItem().asMultipleChoiceItem(); 
  let responseVal = itemResponse.getResponse();
  let choices = question.getChoices();
  for (let i=0; i<choices.length; i++){
    let choice = choices[i];
    if (choice.getValue() == responseVal){
      choices.splice(i, 1);
      break;
    }
  }
  question.setChoices(choices)
}

function _getColData(colNum, sheet){
  return sheet.getRange(2, colNum, studentSheet.getLastRow()-1, 1).getValues().flat()
}


// ---------------- legacy code --------------
// // update the form(flush the choices) when a time slot becomes unavailable
// // toDo: update linear search to binary search
// function _updateForm(response){
//   // get the question ID corresponding to the date selected
//   let timeChoices = _removeResponseFromChoices(timeQIdx);
//   // if no more time slots remain for that day, remove the date choice
//   // note that the select time question for that date still exists, 
//   // but there is no way to access that question as a user
//   // the question will be removed when the service reboots
//   if (timeChoices.length == 0){
//     let dateChoices = _removeResponseFromChoices(dateQIdx);
//     // similarly for date
//     if (dateChoices == 0){
//       let weekChoices = _removeResponseFromChoices(weekQIdx);
//       // if no more weeks are left, then we are done!
//       if (weekChoices.length == 0){
//           let weekResponse = response.getItemResponses()[weekQIdx];
//           let weekQuestion = weekResponse.getItem().asMultipleChoiceItem(); 
//           console.log("All students will be picked up by a RCSSA volunteer!");
//           weekChoices.push(weekQuestion.createChoice("谢谢您的参与！目前所有新生已经找到接机支援者！", FormApp.PageNavigationType.RESTART));
//       }
//     }

//   }   
//     // todo: deal with empty array
  
//   console.log(timeVal);
//   // let time = timeItem.
//   // console.log(items[dateQIdx].getItem().getTitle())
// }

// Get the latest response of the form
// using lock to handle concurrency issue
// function _getLastResponse(){
//   let lastFormResponse;
//   lastFormResponse = form.getResponses().pop(); 
//   for (let itemResponse of lastFormResponse.getItemResponses().slice(firstMultChoiceQIdx)){
//       console.log("3;"+ itemResponse.getItem().asMultipleChoiceItem().getTitle())
//       if (itemResponse.getItem().asMultipleChoiceItem().getTitle().startsWith(questionPrefix.time)){
//         volTime = itemResponse.getResponse();
//           if (!volTime.startsWith(nullChoiceStr)){
//             // due to concurrency, some volunteer might select a datetime that is already gone
//             console.log(_matchStudent("123", "123", "4312", volTime, 2))
//           }
//       }
//     }
//   return lastFormResponse;
// }

// function addInfoQuestions(){
//   let section1 = form.addPageBreakItem()
//     .setTitle('联系方式')
//     .setHelpText('感谢您参与2022年RCSSA接机！')
//   form.addTextItem().setTitle('您的姓名');
//   form.addTextItem().setTitle('您的联系方式').setHelpText("微信号和手机号更方便和新生联系~ 新生也会提供他们的微信号的手机号");
//   return section1
// }

// function addNumPeopleQuestion(){
//     let numStudentSection = form.addPageBreakItem()
//     .setTitle('接机人数')
//     .setHelpText('如果有多个新生在同一时间抵达 我们根据您填写的人数进行匹配 新生的行李较多（可能有2个托运箱子） 请您也考虑下行李所需要的空间');
//     form.addTextItem().setTitle('请填写接机人数');
//     return numStudentSection;
// }


