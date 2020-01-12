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
  /**
   * Insert your Outlook code here
   */
  getSubject();
  getAllRecipients();
  
  // Get a reference to the current message
var item = Office.context.mailbox.item;
item.enhancedLocation.getAsync(callbackFunction);

// Write message property value to the task pane
//document.getElementById("item-subject").innerHTML = "<b>Titolo:</b> <br/>" + item.subject;
//document.getElementById('message').innerText += item.subject;
}

// Get the subject of the item that the user is composing.
function getSubject() {
  var item = Office.context.mailbox.item;
  item.subject.getAsync(
      function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed){
              write(asyncResult.error.message);
          }
          else {
              // Successfully got the subject, display it.
              write ('Il Titolo è: ' + asyncResult.value);
          }
      });
}

// Write to a div with id='message' on the page.
function write(message){
  document.getElementById('item-subject').innerText += message; 
}

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    var result = "Display name: " + place.displayName;
    result +=" Type: " + place.locationIdentifier.type;
      if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
          result +="Email address: " + place.emailAddress;
      }
      document.getElementById('item-location').innerText += result; 
      document.getElementById('tbSala').value = result; 
  });
}




// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
  var item = Office.context.mailbox.item;
  // Local objects to point to recipients of either
  // the appointment or message that is being composed.
  // bccRecipients applies to only messages, not appointments.
  var toRecipients, ccRecipients, bccRecipients;
  // Verify if the composed item is an appointment or message.
  if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
      toRecipients = item.requiredAttendees;
      ccRecipients = item.optionalAttendees;
  }
  else {
      toRecipients = item.to;
      ccRecipients = item.cc;
      bccRecipients = item.bcc;
  }
  
  // Use asynchronous method getAsync to get each type of recipients
  // of the composed item. Each time, this example passes an anonymous 
  // callback function that doesn't take any parameters.
  toRecipients.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed){
          write(asyncResult.error.message);
      }
      else {
          // Async call to get to-recipients of the item completed.
          // Display the email addresses of the to-recipients. 
          writePerson ('Obbligatorio:');
          displayAddresses(asyncResult);
      }    
  }); // End getAsync for to-recipients.

  // Get any cc-recipients.
  ccRecipients.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed){
          write(asyncResult.error.message);
      }
      else {
          // Async call to get cc-recipients of the item completed.
          // Display the email addresses of the cc-recipients.
          writePerson ('Facoltativo:');
          displayAddresses(asyncResult);
      }
  }); // End getAsync for cc-recipients.

  // If the item has the bcc field, i.e., item is message,
  // get any bcc-recipients.
  if (bccRecipients) {
      bccRecipients.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed){
          writePerson(asyncResult.error.message);
      }
      else {
          // Async call to get bcc-recipients of the item completed.
          // Display the email addresses of the bcc-recipients.
          writePerson ('BCC:');
          displayAddresses(asyncResult);
      }
                      
      }); // End getAsync for bcc-recipients.
   }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
  for (var i=0; i<asyncResult.value.length; i++)
      writePerson (asyncResult.value[i].displayName);
}

// Writes to a div with id='message' on the page.
function writePerson(message){
  document.getElementById('messagePerson').innerText += message + "\n\n"; 
}
