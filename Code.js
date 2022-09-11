// global constants
const formId = "1uyC3qT7OLisQn1_ywd5Ks0WANukBzslt8CP6tYw6J68";
const form = FormApp.openById(formId);
// group QRcode
const qrCodeFileId = "1BCo66b-O-ablSuihrBZU1hpc1Kb-woso"
const qrCodeBlob = DriveApp.getFileById(qrCodeFileId).getAs('image/jpeg')

function bind(){
    ScriptApp.newTrigger('myOnFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();
}

function myOnFormSubmit(event) {
  let email = event.response.getRespondentEmail();
  sendConfirmationEmail(email);
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
  
