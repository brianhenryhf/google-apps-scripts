/**
 * Script to cleanup older consuming email acct cruft. Run nightly or as desired.
 * 
 * This establishes labels that, when applied to threads (ideally via Gmail filters), will cause them to automatically
 * 'expire' and get trashed after a set number of days, to reduce manual work cleaning up marketing emails. Labels 
 * should be created manually (e.g., "daily deals", etc.).
 */


function killOldDailyDeals() { //daily
  __.debug("start")
  __.killWithLabelBeforeDate("daily deals", __.daysAgo(2));
}

function killOldEventsLists() {//assumes weekly
  __.killWithLabelBeforeDate("events lists", __.daysAgo(8));
}

function killOldTravelDeals() {//assumes every few days
  __.killWithLabelBeforeDate("travel deals", __.daysAgo(5));
}

function killShortTermSales() {//give them a week-ish?
  __.killWithLabelBeforeDate("short-term sales", __.daysAgo(6));
}

function killLongTermSales() {//give them a month-ish
  __.killWithLabelBeforeDate("long-term sales", __.daysAgo(30));
}

function killOldMonthlyEventsLists() {//give them a month-ish
  __.killWithLabelBeforeDate("monthly events lists", __.daysAgo(30));
}

var __ = function() { //internals
  //hiding the vars, but exposing the fxns - intent is to let Apps Script editor pick up the "public" fxns automatically in the run menu,
  //but sequester the private internal ones here, and stop access to the impl vars altogether
  var DEBUG = false;//turn on/off debug logs
  var DAY_IN_MILLIS = 86400000;  

  return {
    //note this version doesn't allow for starring, unread, '!!!' label or any other means of counteracting auto-kill
    killWithLabelBeforeDate: function(killLabel, killDate) {
      var page,
          threadIdx,
          thread;
      var targetLabel= GmailApp.getUserLabelByName(killLabel);

      //can't get total thread count, so best we can do is know we didn't get less than the max last iteration
      //checking for page nullness, which will be true first time through only
      while (!page || page.length == 100) { 
        page = targetLabel.getThreads(0, 100);
        this.debug("page.length:" + page.length); 
        if(page.length > 0) {
          for(threadIdx = 0; threadIdx < page.length; threadIdx++) {
            thread = page[threadIdx];
            this.debug("thread subj: " + thread.getFirstMessageSubject());
            if(thread.getLastMessageDate() < killDate) { //from whenever script is run
              thread.moveToTrash();
              //this.debug("i'd delete " + thread.getFirstMessageSubject() + " from " + thread.getLastMessageDate())
            }
          }
        }
      }
    },
    
    daysAgo: function(num) {
      return new Date(new Date().getTime() - (DAY_IN_MILLIS * num));
    },
    
    debug: function(msg) {
      if(DEBUG) Logger.log(msg); 
    }
  }
}();
