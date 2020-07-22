/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // Get a reference to the current message
var item = Office.context.mailbox.item;

// Write message property value to the task pane
//document.getElementById("item-subject").innerHTML = "<b>Subject:</b> <br/>" + item.subject;
doAuth();
//generateToken();
/*Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);

  } else {
    // Handle the error.
  }
});*/
}
   

  function getCurrentItem(accessToken){
   // sendMail();
  }

  function doAuth(){


    $.ajax({  
      "async": true,  
      "crossDomain": true,  
      "url": "https://cors-anywhere.herokuapp.com/https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/oauth2/v2.0/authorize", // Pass your tenant name instead of sharepointtechie    
      "method": "POST",  
      "headers": {  
          "content-type": "application/x-www-form-urlencoded"  
      },  
      "data": {  
          "response_type": "code",  
          "client_id ": "e91c8169-bc9d-4e90-a893-62051c4f5378", //Provide your app id    
          "redirect_uri": "https://localhost:3000",
          "scope ": "offline_access user.read mail.read",
          "response_mode":'query',
          "state":12345  
      },  
      success: function(response) {  
          console.log(response);  
          var token = response.access_token;
          //document.getElementById('item-subject').innerHTML = "<b>Token:</b> <br/>" + token;
         // sendMail(token);
      },
      error: function (error) {
        console.log("Error in getting data: " + error); 
    }  

  }); 

}


  function generateToken(){


    $.ajax({  
      "async": true,  
      "crossDomain": true,  
      "url": "https://cors-anywhere.herokuapp.com/https://login.microsoftonline.com/f8cdef31-a31e-4b4a-93e4-5f571e91255a/oauth2/v2.0/token", // Pass your tenant name instead of sharepointtechie    
      "method": "POST",  
      "headers": {  
          "content-type": "application/x-www-form-urlencoded"  
      },  
      "data": {  
          "grant_type": "client_credentials",  
          "client_id ": "e91c8169-bc9d-4e90-a893-62051c4f5378", //Provide your app id    
          "client_secret": "D2_W-FZ97r544-F.NRZ-T--u1gQ_kzd7qu", //Provide your client secret genereated from your app
          "scope ": "https://graph.microsoft.com/.default"  
      },  
      success: function(response) {  
         // console.log(response);  
          var token = response.access_token;
          //document.getElementById('item-subject').innerHTML = "<b>Token:</b> <br/>" + token;
          sendMail(token);
      },
      error: function (error) {
        console.log("Error in getting data: " + error); 
    }  

  }); 

    /*
    var sendUrl = 'https://outlook.office.com/api/v2.0/me/sendmail';
    jQuery.ajax({
      type:'POST',
      url: sendUrl,
      dataType: 'json',
      headers: { 'Authorization': 'Bearer ' + accessToken }
      }).done(function(response){
      // Message is passed in `item`.
      //var subject = item.Subject;
      console.log(response); 
      
      }).fail(function(error){
      // Handle error.
      console.log('error');
      });
    //document.getElementById("item-subject").innerHTML = "<b>Token:</b> <br/>" + accessToken;
  }*/

/*var getMessageUrl = Office.context.mailbox.restUrl +
'/v2.0/me/messages/' + itemId;

var sendUrl = 'https://outlook.office.com/api/v2.0/me/sendmail';*/

/*$.ajax({
url: sendUrl,
dataType: 'json',
headers: { 'Authorization': 'Bearer ' + accessToken }
}).done(function(item){
// Message is passed in `item`.
var subject = item.Subject;

}).fail(function(error){
// Handle error.
});*/

}

function sendMail(token){
  console.log('Token in sendMail: '+token);

  $.ajax({  
    "async": true,  
    "crossDomain": true,  
    "url": "https://graph.microsoft.com/v1.0/users/prateektyagi1@outlook.com/sendMail",  
    "method": "POST",  
    "headers": { 
        'Authorization': 'Bearer ' + token, 
        "content-type": "application/json"  
    },  
    "processData": false, 
    "data": JSON.stringify({
      "message": {
        "subject": "Meet for lunch?",
        "body": {
          "contentType": "Text",
          "content": "The new cafeteria is open."
        },
        "toRecipients": [
          {
            "emailAddress": {
              "address": "pratiktyagi1@gmail.com"
            }
          }
        ]
      },
      "saveToSentItems": "false"
    }),  
    success: function(response) {  
        console.log(response);  
        //var token = response.access_token;
        //document.getElementById('item-subject').innerHTML = "<b>Token:</b> <br/>" + token;
        //sendMail(token);
    },
    error: function (error) {
      console.log(error);
  }  

});
}