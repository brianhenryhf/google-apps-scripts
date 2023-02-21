/* 
UPDATE 2023:  This script is from ~2010 - this fxnality now built-in to gmail, so this script is defunct.

Script that provides snoozing labels, which you can move stuff to, and are then moved up the snooze chain, 
or eventually back to inbox, daily (assuming you set up the trigger to do so.).  
*/

var MARK_UNREAD = false;
var ADD_UNSNOOZED_LABEL = false;
var UNSNOOZE_VIA_FWD = true; //resend to self, to mark unread and bump to top of inbox.  really, mutex with at least mark_unread
var MAX_SNOOZE_DAYS = 7;

function getLabelName(i) {
    return "Snooze/Snooze " + i + " days";
}

//to run once
/*function setup() {
    // Create the labels we'll need for snoozing
    GmailApp.createLabel("Snooze");
    for (var i = 1; i <= MAX_SNOOZE_DAYS; ++i) {
        GmailApp.createLabel(getLabelName(i));
    }
    if (ADD_UNSNOOZED_LABEL) {
        GmailApp.createLabel("Unsnoozed");
    }
}*/


//run via "Day timer" and some off-peak time. 
function moveSnoozes() {
    var oldLabel, newLabel, page;
    for (var i = 1; i <= MAX_SNOOZE_DAYS; ++i) { 
        newLabel = oldLabel;
        oldLabel = GmailApp.getUserLabelByName(getLabelName(i));
        page = null;

        // Get threads in "pages" of 100 at a time 
        while (!page || page.length == 100) {
            page = oldLabel.getThreads(0, 100);
            if (page.length > 0) {
                if (newLabel) {
                    // Move the threads into "today's" label 
                    newLabel.addToThreads(page);
                } else {
                  if(!UNSNOOZE_VIA_FWD) {
                    // Unless it's time to unsnooze it 
                    GmailApp.moveThreadsToInbox(page);
                    if (MARK_UNREAD) {
                      GmailApp.markThreadsUnread(page);
                    }
                    if (ADD_UNSNOOZED_LABEL) {
                      GmailApp.getUserLabelByName("Unsnoozed").addToThreads(page);
                    }
                  } else {
                    for(var zz = 0; zz < page.length; zz++) {
                      var tmpMessages = page[zz].getMessages();
                      var tmpMsg = tmpMessages[tmpMessages.length - 1]; 
                      var subjLead = "!SNZ: ";
                      var subj = tmpMsg.getSubject();
                      var reSnooze = (subj.substr(0, subjLead.length) == subjLead);  // this is a snooze of a snooze
                      if(!reSnooze) subj = subjLead + subj;
                      tmpMsg.forward(Session.getEffectiveUser().getEmail(), {
                        name: "My Snooze Service",
                        subject: subj  
                      });  
                      if(!reSnooze) {
                        page[zz].markRead(); //note old one will still exist in archive.  label it for delete later..
                        GmailApp.getUserLabelByName("Snooze/forwarded").addToThread(page[zz]);
                      } else {
                        //safe enough to delete as this is just a forward from a prev snooze
                        page[zz].moveToTrash();
                      }
                    }
                  }
                }
                // Move the threads out of "yesterday's" label 
                oldLabel.removeFromThreads(page);
            }
        }
    }
  Logger.log("moved snoozes @ " + new Date())
}