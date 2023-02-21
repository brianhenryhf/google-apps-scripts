/**
 * OOO Calendar Syncer
 * run time-based trigger every hour
 *
 * One-way sync of cal data from folks' public calendar 'ooo'/'out of office' events to a shared calendar for an org.
 * 
 * Notes: 
 * - A special 'type' of event for "out of office", as presented in the gcal 'create event' dialog is not real/exposed.  
 *   This script keys in on event names instead, which can result in false positives, but best we can do with current
 *   APIs/model from our Google.
 * - Events deleted from shared calendar are not replaced by later syncs (unless clearXXXRun() is called)
 * - Events deleted from source calendar ARE removed from shared calendar. (do not need to clear run)
 * - This only looks 3 months out.
 * - Import is generally smart enough to not duplicate events, and will instead update imported events instead of 
 *   duplicate (verified: name/description changes).  This is true after clearing runs as well.

 * Dev Notes
 * Swiped/modified from https://developers.google.com/gsuite/solutions/vacation-calendar
 * View logs in stackdriver logging, not Logger logs (so, NOT View->Logs)

 * Times and Quotas: 
 * This can take some time, and apps scripts time out around 4-6 mins, so watch that. Running this too many times in a row can bump into 
 * quota issues, likely with the running user, which will turn up in stackdriver logs.  Possible working in 
 * userIp or quotaUser could help (unexplored).  May be possible to increase quotas (see 
 * https: *developers.google.com/apps-script/guides/services/quotas)
 */

var KEYWORDS = ['vacation', 'vacations', 'ooo', 'out of office', 'offline', 'holiday', 'holidays'];
var MONTHS_IN_ADVANCE = 3;

// this can use one of two apis to get events to send to `DEST_CALENDAR_ID` - `groupEmail` can be supplied, but has been
// seen not to work with some aliases/groups for unresolved reasons. In that case, set that to null and be sure to
// supply `domain` to search cals in that domain (this mode may require more intense admin permissions for script-runner).
var MYGROUP_CALENDAR_CONFIG = {
  //e.g., c_abcdefghijk@group.calendar.google.com
  calId: PropertiesService.getScriptProperties().getProperty('DEST_CALENDAR_ID'),
  groupEmail: null,
  domain: PropertiesService.getScriptProperties().getProperty('DOMAIN'),
  propHandle: 'mygroup'
};

/** function to run on a time-based trigger (say, every few hours) */
function syncMygroupCalendar() {
  sync(MYGROUP_CALENDAR_CONFIG);
}

/** this function is for ad hoc debug/troubleshoot/etc */
function clearLastMygroupRun() {
  clearLastRun(MYGROUP_CALENDAR_CONFIG);
}


//----- 

// Clear out the last run property, which is used to optimize skipping events in the past and unchanged
// events in the future (only work with events new/changed since the last run).  Clearing out allows 
// reprocessing of events.
// Note that any events removed from group cal will be re-found and -added when reprocessing occurs.
function clearLastRun(targetCalConfig) {
  console.log('clearing last run for ' + targetCalConfig.propHandle);
  PropertiesService.getScriptProperties().deleteProperty(targetCalConfig.propHandle + '_lastRun');
}

/**
 * Look through the group members' public calendars and add any
 * 'vacation' or 'out of office' events to the team calendar.
 */
function sync(targetCalConfig) {
  console.log('starting import for ' + targetCalConfig.propHandle);
  // Define the calendar event date range to search.
  var today = new Date();
  var maxDate = new Date();
  maxDate.setMonth(maxDate.getMonth() + MONTHS_IN_ADVANCE);

  // Determine the time the the script was last run.
  var lastRun = PropertiesService.getScriptProperties().getProperty(targetCalConfig.propHandle + '_lastRun');
  lastRun = lastRun ? new Date(lastRun) : null;

  var users;
  if(targetCalConfig.groupEmail) {
    users = getGroupUsers(targetCalConfig.groupEmail);
  } else {
    users = getAllUsers(targetCalConfig.domain);
  }
  
  // For each user, find events having one or more of the keywords in the event
  // summary in the specified date range. Import each of those to the team
  // calendar.
  var count = 0;
  users.forEach(function(user) {
    KEYWORDS.forEach(function(keyword) {
      var events = findEvents(user, keyword, today, maxDate, lastRun);
      events.forEach(function(event) {
        importEvent(targetCalConfig, user.username, event);
        count++;
      }); // End foreach event.
    }); // End foreach keyword.
  }); // End foreach user.

  // NOTE this sometimes errors - "Data storage error".  in which case, script fails and next run will do a little more
  // work than it has to, but shouldn't duplicate events.
  PropertiesService.getScriptProperties().setProperty(targetCalConfig.propHandle + '_lastRun', today);
  console.log('Imported ' + count + ' events');
}

/**
 * Imports the given event from the user's calendar into the shared team
 * calendar.
 * @param {string} username The team member that is attending the event.
 * @param {Calendar.Event} event The event to import.
 */
function importEvent(targetCalConfig, username, event) {
  event.summary = '[' + username + '] ' + event.summary;
  event.organizer = {
    id: targetCalConfig.calId,
  };
  event.attendees = [];
  console.log('Importing: %s', event.summary);
  try {
    // info on this is not clear - not part of base calendar scripts api, but maybe using REST api
    // via https://developers.google.com/apps-script/advanced/calendar - though args are unclear from looking 
    //  at REST api docs.
    Calendar.Events.import(event, targetCalConfig.calId);
  } catch (e) {
    console.error('Error attempting to import event: %s. Skipping.',
        e.toString());
  }
}

/**
 * In a given user's calendar, look for occurrences of the given keyword
 * in events within the specified date range and return any such events
 * found.
 * @param {Object} user The user to retrieve events for.
 * @param {string} keyword The keyword to look for.
 * @param {Date} start The starting date of the range to examine.
 * @param {Date} end The ending date of the range to examine.
 * @param {Date} optSince A date indicating the last time this script was run.
 * @return {Calendar.Event[]} An array of calendar events.
 */
function findEvents(user, keyword, start, end, optSince) {
  var params = {
    q: keyword,
    timeMin: formatDateAsRFC3339(start),
    timeMax: formatDateAsRFC3339(end),
    showDeleted: true,
  };
  if (optSince) {
    // This prevents the script from examining events that have not been
    // modified since the specified date (that is, the last time the
    // script was run).
    params.updatedMin = formatDateAsRFC3339(optSince);
  }
  var pageToken = null;
  var events = [];
  do { //loop pages for user/kw list request
    var response;
    var MAX_TRIES = 5;
    var tryNum = 0;
    params.pageToken = pageToken;
    response = null;

    do { //retries for failure on a page
      tryNum++;
      try {
        response = Calendar.Events.list(user.email, params);
      } catch (e) {
        //this happens on random runs:  error tostring = "Empty response".  unclear why.  retrying once ususally fixes, but sometimes requires more.  
        // list can also fail for users who have not set their calendar to shared
        console.error('Error retriving events for %s, kw %s: %s; try %s of %s for page.',
            JSON.stringify(user), keyword, e.toString(), tryNum, MAX_TRIES);
        if(tryNum < MAX_TRIES) Utilities.sleep(1000);  //hang on a sec            
        else if(tryNum == MAX_TRIES) console.log('out of tries for page.');
      }
    } while(response == null && tryNum < MAX_TRIES)

    if(response == null) {
      console.log('response was never obtained for page.  moving on.');
      pageToken = null;  //we've done all we can do for this user/kw, so this stops outer loop
    } else {
      events = events.concat(response.items.filter(function(item) {
        return shoudImportEvent(user, keyword, item);
      }));
      pageToken = response.nextPageToken;
    }
  } while (pageToken);
  return events;
}

/**
 * Determines if the given event should be imported into the shared team
 * calendar.
 * @param {Object} user The user that is attending the event.
 * @param {string} keyword The keyword being searched for.
 * @param {Calendar.Event} event The event being considered.
 * @return {boolean} True if the event should be imported.
 */
function shoudImportEvent(user, keyword, event) {
  // Filter out events where the keyword did not appear in the summary
  // (that is, the keyword appeared in a different field, and are thus
  // is not likely to be relevant).
  if (event.summary.toLowerCase().indexOf(keyword) < 0) {
    return false;
  }
  if (!event.organizer || event.organizer.email == user.email) {
    // If the user is the creator of the event, always import it.
    return true;
  }
  // Only import events the user has accepted.
  if (!event.attendees) return false;
  var matching = event.attendees.filter(function(attendee) {
    return attendee.self;
  });
  return matching.length > 0 && matching[0].responseStatus == 'accepted';
}

/**
 * Return an RFC3339 formated date String corresponding to the given
 * Date object.
 * @param {Date} date a Date.
 * @return {string} a formatted date string.
 */
function formatDateAsRFC3339(date) {
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd\'T\'HH:mm:ssZ');
}


//returns a scaled down user obj with extracted bits common b/w admin api and group api
function getAllUsers(domain) {
  var pageToken;
  var page;
  var users = [];
  do {
    //groups api can be weird - possibly if executing user is not in a group, the group can be sort of found, but internally an error is thrown
    //so, traversal by group hits a roadblock; use admin api instead to get all users.
    page = AdminDirectory.Users.list({
      domain: domain,
      orderBy: 'email',
      maxResults: 200,
      pageToken: pageToken
    });
    if(page.users) users = users.concat(page.users);
    pageToken = page.nextPageToken;
  } while (pageToken);
  return adaptFromAdminApiUsers(users);
}  

//returns a scaled down user obj with extracted bits common b/w admin api and group api
function getGroupUsers(groupEmail) {
  return adaptFromGroupsAppUsers(GroupsApp.getGroupByEmail(groupEmail).getUsers());
}

function adaptFromAdminApiUsers(users) {
  return users.map(function(it) {
    return {
      email: it.primaryEmail,
      username: it.primaryEmail.split('@')[0]
    };
  });
}

function adaptFromGroupsAppUsers(users) {
  return users.map(function(it) {
    return {
      email: it.getEmail(),
      username: it.getEmail().split('@')[0]
    };
  });
}

/*
function setup() {
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    throw new Error('Triggers are already setup.');
  }
  ScriptApp.newTrigger('sync').timeBased().everyHours(1).create();
  // Run the first sync immediately.
  sync();
}
*/