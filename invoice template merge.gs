//  IN OK SHAPE, fxnally (absolute mess otherwise) - add in field formatting, clean up (encapsulate?).  don't go nuts here.  MAYBE an eye for generalized single-doc mail merge.  
//ok, this prolly works ok. use it. check rounding.  polish for vanity/posting?

//todo - could have defaults that figure out the period (default), but allow override,  eh, this would be crappy in backing record.  anyway, that might want a scanning-for-tags approach (syntax could be like {{default|entered}})

//some inconsistencies (invoice num vs id, file naming format vs invoiceid format, for hitorical reasons.  in future, move these closer)

//we could really perhaps just map from 
// spreadsheet cols to keys.  but this is building a littel more smarts into metadata now, so  maybe not...

//null replacement key means not a field involved in replacing.  'required' is required for row to be considered complete.
//turns out this very explicit key to obj mapping is rarely useful ('generated'), but not hurting much.  Actually, could be more flexible to just use the
//replacement keys as col names.  But also more cryptic.  For now, not doing.  oh, also, this is future-crappy - if i change something (filename is compose of xyz instead of xyq), i have to append new cols at end.  more flexible to be able to add cols in middle, near old context.  tradeoff is i'd have to keep
// colnames pretty static.  can't see that being a problem.
// could also have formatting fxns.  and repalcement could instead be a function that works on other bits of data.  but this works for now.

// can insert a checkbox in a sheet - useful in some real way?  can you have a button or link to force a specific row to get generated?

//ok, might be nice to also just have ctor to make a row obj from data that answers these questions... formatting could be context-aware.
const DATA_KEY_MAP = {
  generated: { dataGridIdx: 0, replacementKey: null, required: false },
  customerId: { dataGridIdx: 1, replacementKey: null, required: true },
  invoiceNum: { dataGridIdx: 2, replacementKey: null, required: true },
  invoiceId: { dataGridIdx: 3, replacementKey: 'INVC_ID', required: true },
  invoiceDate: { dataGridIdx: 4, replacementKey: 'INVC_DT', valTransform: (val) => val.toLocaleDateString(), required: true },
  workPeriodStart: { dataGridIdx: 5, replacementKey: 'PERIOD_ST_INCL', valTransform: (val) => val.toLocaleDateString(), required: true },  //see INVC_DT
  workPeriodEnd: { dataGridIdx: 6, replacementKey: 'PERIOD_END_INCL', valTransform: (val) => val.toLocaleDateString(), required: true },  //see INVC_DT
  totalHours: { dataGridIdx: 7, replacementKey: 'TOT_HRS', valTransform: (val) => new Intl.NumberFormat('en-US', {minimumFractionDigits: 1}).format(val), required: true },  //TODO - what happens with weird rounding - 1.7 hours...?
  hourlyRate: { dataGridIdx: 8, replacementKey: 'HRLY_RT', valTransform: (val) => new Intl.NumberFormat('en-US', {style: 'currency', currency: 'USD'}).format(val), required: true },  //fmt (sheet fmt is discarded)
  totalAmtDue: { dataGridIdx: 9, replacementKey: 'TOT_DUE', valTransform: (val) => new Intl.NumberFormat('en-US', {style: 'currency', currency: 'USD'}).format(val), required: true },  //fmt
  templateFileId: {dataGridIdx : 10, replacementKey: null, required: true },
  payerInfo: { dataGridIdx: 11, replacementKey: 'PAYER_INFO', required: true },
  payeeInfo: { dataGridIdx: 12, replacementKey: 'PAYEE_INFO', required: true },
  lineItemDesc: { dataGridIdx: 13, replacementKey: 'LI_DESC', required: true },
  invoiceNote: { dataGridIdx: 14, replacementKey: 'INV_NOTE', required: false },
  terms: { dataGridIdx: 15, replacementKey: 'TERMS', required: true },  //old ones won't display this.  not gonna backfill data for now.
  dueDate: { dataGridIdx: 16, replacementKey: 'DUE_DT', valTransform: (val) => val.toLocaleDateString(), required: true }
};

// this is hard to read - the data_key_map thing should be an object that translates a rows's datat to easy field accessors.

//val:  rowFields[DATA_KEY_MAP.customerId.dataGridIdx]
const row = (rowFields) => {
  return Object.entries(DATA_KEY_MAP).reduce((agg, curr) => {
    agg[curr[0]] = agg[curr[1]  /*..xfer keys of sub objects to this object. explicit naming is fine...  prolly some easy es5 way to merge in fields? actually, this may be it already?  */   ]
  }, {})
}

/** barring a better way to do this, this primes a test fxn that behaves like an actual handler of real sheet event.  */
const TEST_SheetEditEvent = () => {
  const evt = {
    source: SpreadsheetApp.openById('1epL9J8JOQJS1rRnZB6tscSLq9JGyAUQaunsthckEjWc')
  };

  handleOnEdit(evt);
}

/** for a given row's array of cols, this does any needed transforms to produce presentable result */
const buildReplacementMap = (rowFields) => {
  return Object.values(DATA_KEY_MAP).reduce((agg, curr) => {
    if(curr.replacementKey === null) return agg;
    agg[curr.replacementKey] = curr.valTransform ? curr.valTransform(rowFields[curr.dataGridIdx]) : rowFields[curr.dataGridIdx];
    return agg;
  }, {});
}

/*
trigger is a bit tricky - i guess, on edit (which could be even one cell), if ANY row is complete for required data, AND is not marked created, run the op to generate.  then mark created.  or could store the prev row count, then onedit check current row count ( i guess replacing a row would not trigger thenn, which is weird.)
*/

//note doc number formatting is not real and is ignored here.  force any fmting you want here..

/** this is the main entrypoint - edit a sheet, face the fiery wrath of this function. */
const handleOnEdit = (evt) => {
  const ss = evt.source;
  const dataSheet = ss.getSheets()[0];

//Omit header row for 'data range'
  //const dataRange = dataSheet.getRange(2, dataSheet.getLastColumn());
  const contentRange = dataSheet.getDataRange();
  // const dataRange = dataSheet.getRange(2, contentRange.getColumn(), contentRange.getNumRows() - 1, contentRange.getNumColumns());
  const dataRange = util.accordianAdjustRange(contentRange, {addRows: 1});
  
  //all rows are offset to exclude header rows (and in normal 0-index scheme) - so 0 is first row
  const data = dataRange.getValues();

  data.forEach((rowFields, idx) => {
    //prolly want some mapping obj to take col nums (or could try to get header names, but meh) and map to replacement keys
    //only consider rows where not marked invoice already generated
    if(!shouldCreate(rowFields)) return;
    invoiceReplacements = buildReplacementMap(rowFields);

    //create it
    const customerIdCamel = util.camelize(rowFields[DATA_KEY_MAP.customerId.dataGridIdx]);
    
    //ok so, here we might use a transform fxn built in to the field....TODO
    const invoiceNumber = rowFields[DATA_KEY_MAP.invoiceNum.dataGridIdx]?.toString().padStart(2, '0');
    const fileName = `${customerIdCamel}Invoice_${invoiceNumber}`; 

    createFilledDoc(rowFields[DATA_KEY_MAP.templateFileId.dataGridIdx], fileName, invoiceReplacements);

    //mark it
    const generatedCell = dataRange.offset(idx, DATA_KEY_MAP.generated.dataGridIdx, 1, 1);
    generatedCell.setValue('X')
  });
};

//check for complete row that's not had generation before
const shouldCreate = (rowFields) => {
  if(util.present(rowFields[DATA_KEY_MAP.generated.dataGridIdx])) return false;
  return Object.values(DATA_KEY_MAP).every((spec) => !spec.required || util.present(rowFields[spec.dataGridIdx]));
}

const createFilledDoc = (templateFileId, newDocName, replacementMap) => {
  const newDoc = newDocFromTemplate(templateFileId, newDocName);
  const newDocBody = newDoc.getBody();

  replaceKeys(newDocBody, replacementMap);
};

const replaceKeys = (docBody, replacementMap) => {
  //replaceText will replace all instances, which is fine for current purposes
  Object.entries(replacementMap).forEach(([key, val]) => {
    docBody.replaceText(`\{\{${key}\}\}`, val);  //descends through doc tree to replace all the things. yay.
  })
};

const newDocFromTemplate = (templateFileId, newFileName) => {
  const file = DriveApp.getFileById(templateFileId).makeCopy(newFileName);  //yay, same folder (well, future state might be a subfolder for the company id?).  if file by name already exists, another of same name is created anyway.
  const newfileId = file.getId();
  
  return DocumentApp.openById(newfileId);
};


//-- general utils
const util = {
  accordianAdjustRange: (range, {addRows = 0, addColumns = 0}) => {
    return range.offset(addRows, addColumns, range.getNumRows() - addRows, range.getNumColumns() - addColumns);
  },

  camelize: (str) => str && str.split(' ').map(it => it[0].toUpperCase() + it.slice(1)).join(''),
  
  // non-null and non empty/whitespace.  Naming is iffy, but monkeying ruby/rails, php, etc..
  //TODO reexamine - does this do the same surprising things ruby/php does, w/ booleans etc?  if not, rename...
  present: (str) => !!str?.toString().match(/\S+/)
};
