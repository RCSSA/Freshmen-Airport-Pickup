// global constants
const formId = "14ehXAOAtRVdbozdNl0rYJ2neJ2fSXFPOI3GGpTRxeSM";
const form = FormApp.openById(formId);
// group QRcode
const qrCodeFileId = "1BCo66b-O-ablSuihrBZU1hpc1Kb-woso"
const qrCodeBlob = DriveApp.getFileById(qrCodeFileId).getAs('image/jpeg')

// new student response (google form) related - col number starts from 0
const studentNameColNum = 1;
const studentWechatColNum = 2;
const studentOtherContactColNum = 3;
const timeColNum = 3;
const timeConfColNum = 4;

function bind(){
    ScriptApp.newTrigger('myOnFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}


function myOnFormSubmit(event) {
  let email = event.response.getRespondentEmail();
  let formRespones = event.response.getItemResponses();
  let studentTimeCol_1 = formRespones[timeColNum].getResponse();
  let studentTimeCol_2 = formRespones[timeConfColNum].getResponse();

  // If time entered twice is different, send time conflict failure email
  // Otherwise, send confirmation email
  if (studentTimeCol_1 != studentTimeCol_2){
    Logger.log(`sending email - input time conflict: ${email}}`)
    sendFailureEmailTimeConflict(email);
  } else {
    sendConfirmationEmail(email);
  }
}

function sendConfirmationEmail(receiver){
    MailApp.sendEmail({
    to: receiver,
    // multi line string!
    subject: `RCSSA: 2022接机群二维码！`,
    name: "James Li",
    body: 
    `Heeey! 
    欢迎加入Rice大家庭！你抵达休斯顿的时间我们已经收到。
    在等待志愿者匹配的同时，欢迎扫描附件内的二维码加入接机群，获取最新的动态。
    此外，志愿者将会于07.26日开始选择时间接机，当你的接机时间被一名RCSSA认证的志愿者选取后，你会再收到一封匹配成功的确认邮件，请多多查收邮箱！

    RCSSA 2022-2023 全体成员
    `,
    attachments: [qrCodeBlob]
  });
}

function sendFailureEmailTimeConflict(receiver){
  Logger.log(`sending failure email due to time inconsistence to ${receiver}`);
  MailApp.sendEmail({
    to: receiver,
    // multi line string!
    subject: `RCSSA: 请确定接机时间后提交申请`,
    name: "RCSSA",
    body: 
    `新生你好！
      很抱歉，您的接机报名没有成功。
      由于您两次填写的接机时间不一致，请您确定好需要接机时间后，如有需要，请您再次填写并提交报名。
      
    祝新学期一帆风顺！

    20222-2023 RCSSA全体成员
    `,
    attachments: [qrCodeBlob]
  });
}
