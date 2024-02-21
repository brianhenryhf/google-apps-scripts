/**
 * Script to cleanup older email acct cruft. Run nightly via trigger, or ad hoc as desired.
 * 
 * This operates on labels of pattern "ttl-Nd-opt-descriptive-info" that, when applied to threads, will cause threads to
 * deleted after label-specified ttl days ("N" in pattern), to reduce manual work cleaning up marketing, job alert, etc. emails. Labels 
 * should be created manually in Gmail and ideally applied automatically via Gmail filters.
 * 
 * Benefit over more semantic label mechanism previously used (e.g, 'daily deals' config'd for N days in code) is this requires no code 
 * changes for arbitrary new expiration lengths or email types, yet allows for "opt-descriptive-info" space to put semantic info about
 * email type, if desired.
 * 
 * (NOTE: May time out after 360s. so if a lot to cleanup - on first run for example - this can result in only newer messages being cleaned up. Meant
 * to run daily and does nothing crucial; it'll eventually cover the right stuff.)
 */

const killAllExpirableLabeleds = () => {
  const targetLabelSpecs = __.parseExpirableLabels();
  targetLabelSpecs.forEach((spec) => {
    __.killWithLabelBeforeDate(spec.label, __.daysAgo(spec.ttlDays + 1));
  });
}

const __ = (() => { //internals
  //Hiding the vars, but exposing the fxns - intent is to let Apps Script editor pick up the "public" 'killxx' fxns automatically in the run menu,
  //but sequester the private internal ones here, and stop access to the impl vars altogether (but still allow ad hoc testing/use of internal 
  //helper fxns.)

  const DEBUG = false; 
  const STUB_DELETE = false;
  const DAY_IN_MILLIS = 86400000;
  //Note: to avoid probable rate limiting errors with low fetch max setting during dev/testing, jamming in a Utilities.sleep(500) in 
  //iteration seems to help.
  const MAX_THREADS_PER_PAGE = 100;

  return {
    parseExpirableLabels: function () {
      const allLabels = GmailApp.getUserLabels();
      
      return allLabels.reduce((acc, curr) => {
        const name = curr.getName();
        const result = name.match(/^ttl-(?<ttlDays>\d+)d?(?:-(?<desc>.*))?/i);
        
        if(result) {
          const {ttlDays, desc} = result.groups;
          acc.push({
            label: name,
            ttlDays: parseInt(ttlDays),
            desc
          });
        }
        return acc;
      }, []);
    },

    //Note this version doesn't allow for starring, unread, '!!!' label or any other means of counteracting auto-kill
    killWithLabelBeforeDate: function (killLabel, killDate) {
      this.debug(`- killing within label "${killLabel}"`);
      const targetLabel = GmailApp.getUserLabelByName(killLabel);
      let page,
          threadFetchStartIdx = 0;

      //Can't get total thread count, so best we can do is know we didn't get less than the max last iteration.
      //Also, checking for page nullness, which will be true first time through only.
      while (!page || page.length == MAX_THREADS_PER_PAGE) { 
        page = targetLabel.getThreads(threadFetchStartIdx, MAX_THREADS_PER_PAGE);
        this.debug(`page.length: ${page.length}`);

        //First thread fetch attempt may result in none; previous iteration that may have gotten the EXACTLY MAX_THREAD_FETCH remaining may 
        //also now result in none.
        if(page.length > 0) {
          for(const thread of page) {
            let threadActuallyDeleted = false;
            
            this.debug(`thread subj: ${thread.getFirstMessageSubject()}`);
            if(thread.getLastMessageDate() < killDate) {
              threadActuallyDeleted = this.deleteThread(thread);
            }
            //If we actually deleted, threads shift left in next fetch. If didn't (not old enough, or just stubbing), adjust start right to avoid
            // re-processing same threads.
            if(!threadActuallyDeleted) threadFetchStartIdx++;
          }
        }
      }
    },

    //Returns whether thread was actually deleted (for fetch start-shifting purposes)
    deleteThread: function (thread) {
      if(STUB_DELETE) {
        this.debug(`i'd delete ${thread.getFirstMessageSubject()} from ${thread.getLastMessageDate()}`);
        return false;
      } else {
        thread.moveToTrash();
        return true;
      }
    },

    daysAgo: (num) => new Date(new Date().getTime() - (DAY_IN_MILLIS * num)),
    
    debug: (msg) => {
      if(DEBUG) Logger.log(msg); 
    }
  }
})();