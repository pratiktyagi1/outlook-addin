/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    msola.msalInit();
    msola.defaultEmail = Office.context.roamingSettings.get('default_email');
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("userAuthBtn").onclick = msola.msalUserLogin;
    document.getElementById("sendEmail").onclick = msola.getMailMime;
    document.getElementById("saveDefault").onclick = msola.setDefaultEmail;
    //document.getElementById("logout").onclick = msola.logOut;
    if(msola.defaultEmail){
      $("#toEmail").val(msola.defaultEmail);
    }
  }
});


var msola = {
  msalInstance:{},
  msAccessToken: '',
  isUserLoggedIn: false,
  defaultEmail:'',
  msalConfig:{
    auth: {
      clientId: "91390636-1b1c-4c9e-ae4d-d91f2cde8aa6",
      authority: "https://login.microsoftonline.com/common",
      postLogoutRedirectUri: "https://localhost:3000/taskpane.html"
    },  
    cache: {  
      cacheLocation: "localStorage"          
    }
  },
  graphConfig:{
    endPoint: "https://graph.microsoft.com/v1.0/me" 
  },
  permissionScope:{
    scopes: ["https://graph.microsoft.com/mail.readwrite","https://graph.microsoft.com/user.read", "https://graph.microsoft.com/mail.send"]
  },
  msalInit:function(){
    $(".progressLoader").show();
    msola.msalInstance = new Msal.UserAgentApplication(msola.msalConfig);
    msola.msalInstance.acquireTokenSilent(msola.permissionScope).then(msola.silentTokenSuccess,msola.silentTokenError); 
  },
  msalUserLogin: function(){
    $(".progressLoader").show();
    msola.msalInstance.loginPopup(msola.permissionScope).then(function (id_token) {
      msola.msalInstance.acquireTokenSilent(msola.permissionScope).then(function (result) { 
        //console.log(result.accessToken);
        msola.msAccessToken = result.accessToken; 
        $(".progressLoader").hide();
        $("#afterLogin").show();
        $("#beforeLogin").hide();
      });
    });
  },
  silentTokenSuccess:function(result){
    console.log('User Already Logged In!');
    //console.log(result.accessToken);
    msola.msAccessToken = result.accessToken;
    $("#beforeLogin").hide();
    $("#afterLogin").show();
    $(".progressLoader").hide();
  },
  silentTokenError: function(error){
    if (error.toString().indexOf("interaction_required" != -1)) {
        msola.msalInstance.acquireTokenPopup(msola.permissionScope).then(function (access_token) {
            console.log("Success acquiring access token");
            //console.log(access_token);
            msola.msAccessToken = access_token;
            $("#afterLogin").show();
            $("#beforeLogin").hide();
            $(".progressLoader").hide();
        }, function (error) {
            console.log("Failure acquiring token: " + error);
            $("#beforeLogin").show();
            $("#afterLogin").hide();
            $(".progressLoader").hide();
            
        });
    }
  },
  settingDialog: function(){
    var dialogOptions = { width: 20, height: 40, displayInIframe: true };
    Office.context.ui.displayDialogAsync('https://localhost:3000/settings.html', dialogOptions, function (result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
    });
  },
  setDefaultEmail: function(){
    var email = $("#toEmail").val();
    if(email){
        var _settings = Office.context.roamingSettings;
        _settings.set("default_email", email);
        _settings.saveAsync(msola.saveSoarSettingsCallback);
    }else{
      msola.showErrorMsg('Please enter email !');
    }
  },
  saveSoarSettingsCallback: function(asyncResult){
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      msola.showErrorMsg('Unable to save as default email !');
    }else{
      var email = $("#toEmail").val();
      msola.showSuccessMsg(email+' saved as default email !');
    }
  },
  sendEmail: function(postData = ''){
    $.ajax({  
      "crossDomain": true,  
      "url": msola.graphConfig.endPoint+"/sendMail",  
      "method": "POST",  
      "data": postData,
      "headers": { 
          'Authorization': 'Bearer ' + msola.msAccessToken, 
          "content-type": "application/json"  
      } 
    }).done(function(response, status, xhr){
      msola.showSuccessMsg('Email sent successfully !');
     
    }).fail(function(jqXHR, textStatus, errorThrown){
      msola.showErrorMsg('Unable to send email!');
    });
  },
  getMailMime: function(){
    var receipentEmail = $("#toEmail").val();
    if(receipentEmail){
      $(".progressLoader").show();
      var item = Office.context.mailbox.item;
      $.ajax({  
        "crossDomain": true,  
        "url": msola.graphConfig.endPoint+"/messages/"+item.itemId+"/$value",  
        "method": "GET",  
        "headers": { 
            'Authorization': 'Bearer ' + msola.msAccessToken, 
            "content-type": "application/json"  
        } 
      }).done(function(response, status, xhr){
        if(status == 'success'){
          var mailContent = msola.buildMessage(response);
          msola.sendEmail(mailContent);
        }
      }).fail(function(jqXHR, textStatus, errorThrown){
        console.log(errorThrown);
        msola.showErrorMsg('Please try later !');
      });
    }else{
      msola.showErrorMsg('Please enter email !');
    }
  },
  buildMessage: function(mailMimeContent){
    var receipentEmail = $("#toEmail").val();
    var encodedMailContent = btoa(mailMimeContent);
    var item = Office.context.mailbox.item;
    var emailContent = JSON.stringify({
      "message":{
        "subject": "Outlook Add-in Final Test",
        "body": {
          "contentType": "Html",
          "content": "<p>Hello Yair</p> This email was sent from outlook add-in. Please test the content of attached eml file. The final flow of the Add-in is ready to be reviewed.<p>Regards,<br/>Vivek Negi</p>"
        },
        "toRecipients": [
          {
            "emailAddress": {
              "address": receipentEmail
            }
          }
        ],
        "attachments": [
          {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": item.subject+".eml",
            "contentType": "application/octet-stream",
            "contentBytes": encodedMailContent
          }
        ]
      },
      "saveToSentItems": "true"
    });
    return emailContent;
  },
  showSuccessMsg: function(msg){
    $(".progressLoader").hide();
    $("#errMsgBar").hide();
    $(".ms-MessageBar-text").text(msg);
    $("#successMsgBar").show();
  },
  showErrorMsg: function(msg){
    $(".progressLoader").hide();
    $(".ms-MessageBar-text").text(msg);
    $("#successMsgBar").hide();
    $("#errMsgBar").show();
  }, 
  logOut: function(){
    msola.msalInstance.logout();
    $("#beforeLogin").show();
    $("#afterLogin").hide();
  }
}