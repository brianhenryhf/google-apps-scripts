/**
Script to merge a source Sheet's rows into a Doc, based on a pre-existing template doc and a simple templating language.  
At present, this is specific to an invoicing use case (though it could perhaps be generalized in the future via Apps
Scripts libraries and such).

The templating language is simple - replacement fields are enclosed like so: `{{REPLACE_THIS}}`. Note that existing
formatting of field key in the doc will be kept (e.g, if you want a bold result, embolden the replacement key in the
template).  

For fancy features like defaults for missing vals or calculated or interdependent values, functions in the source sheet
are suggested. This _does_ allow for some content transforms for visual presentation (e.g., rounding, date and currency
formats), as formatting of source values in the sheet for those purposes does not transfer to the output doc.

Note that the template file to use is specified in the source sheet itself.
*/

// some quirks and inconsistencies exist in these fields (invoice num vs id, file naming format vs invoiceid format) 
// for historical reasons.


// Each value in the below is a spec for a field:
// - dataGridIdx is 0-based column idx in data sheet.
// - replacementKey is key in template, to be replaced by a val. null replacement key means not a field involved in 
//   replacing.  
// - required is requiredness in data sheet - if all required fields are present for a row, merge doc can be generated
//   for that row
// - valTransform is optional fxn to transform val before replacing in destination doc

//TODO might be nice to also just have ctor to make a row obj from data that answers these questions...
const DATA_KEY_MAP = {
  generated: { 
    dataGridIdx: 0, 
    replacementKey: null, 
    required: false 
  },
  customerId: { 
    dataGridIdx: 1, 
    replacementKey: null,
    required: true 
  },
  invoiceNum: { 
    dataGridIdx: 2, 
    replacementKey: null, 
    required: true 
  },
  invoiceId: { 
    dataGridIdx: 3, 
    replacementKey: 'INVC_ID', 
    required: true 
  },
  invoiceDate: { 
    dataGridIdx: 4, 
    replacementKey: 'INVC_DT', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  workPeriodStart: { 
    dataGridIdx: 5, 
    replacementKey: 'PERIOD_ST_INCL', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  workPeriodEnd: { 
    dataGridIdx: 6, 
    replacementKey: 'PERIOD_END_INCL', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  totalHours: { 
    dataGridIdx: 7, 
    replacementKey: 'TOT_HRS', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { minimumFractionDigits: 1 }).format(val), 
    required: true 
  },
  hourlyRate: { 
    dataGridIdx: 8, 
    replacementKey: 'HRLY_RT', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val), 
    required: true 
  },
  totalAmtDue: { 
    dataGridIdx: 9, 
    replacementKey: 'TOT_DUE', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val), 
    required: true 
  },
  templateFileId: { 
    dataGridIdx: 10, 
    replacementKey: null, 
    required: true 
  },
  payerInfo: { 
    dataGridIdx: 11, 
    replacementKey: 'PAYER_INFO', 
    required: true 
  },
  payeeInfo: { 
    dataGridIdx: 12, 
    replacementKey: 'PAYEE_INFO', 
    required: true 
  },
  lineItemDesc: { 
    dataGridIdx: 13, 
    replacementKey: 'LI_DESC', 
    required: true 
  },
  invoiceNote: { 
    dataGridIdx: 14, 
    replacementKey: 'INV_NOTE', 
    required: false 
  },
  terms: { 
    dataGridIdx: 15, 
    replacementKey: 'TERMS', 
    required: true 
  },
  dueDate: { 
    dataGridIdx: 16, 
    replacementKey: 'DUE_DT', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  }
};

/** This is the main entrypoint - edit a sheet, face the fiery wrath of this function. */
const handleOnEdit = (evt) => {
  const ss = evt.source;
  const dataSheet = ss.getSheets()[0];

  processSheet(dataSheet);
};

const processSheet = (sheet) => {
  const contentRange = sheet.getDataRange();
  // trim off the header row, so row 0 starts the data
  const dataRange = util.accordianAdjustRange(contentRange, { addRows: 1 });
  const data = dataRange.getValues();

  data.forEach((rowFields, idx) => {
    if (!shouldCreate(rowFields)) return;
    invoiceReplacements = buildReplacementMap(rowFields);

    // historical naming convention - not ideal that we reference specific fields here and custom transform.
    // TODO a case could be made that this filename should be written to source sheet for recordkeeping
    const customerIdCamel = util.camelize(rowFields[DATA_KEY_MAP.customerId.dataGridIdx]);
    const invoiceNumber = rowFields[DATA_KEY_MAP.invoiceNum.dataGridIdx]?.toString().padStart(2, '0');
    const fileName = `${customerIdCamel}Invoice_${invoiceNumber}`;

    createFilledDoc(rowFields[DATA_KEY_MAP.templateFileId.dataGridIdx], fileName, invoiceReplacements);

    //mark it done - using a marker column instead of checking for row count changes lets us re-generate from rows for testing or whatever reason we want.
    const generatedCell = dataRange.offset(idx, DATA_KEY_MAP.generated.dataGridIdx, 1, 1);
    generatedCell.setValue('X')
  });
}


//val:  rowFields[DATA_KEY_MAP.customerId.dataGridIdx]
// the idea here was less nasty way to reference fields - `row(rowFields).generated.dataGridIdx` instead of `rowFields[DATA_KEY_MAP.generated.dataGridIdx]`
// eh, is that really much better?  oh, woudl most usages just be row(rowFields).generated to get specific value (refer to field specs for metadata)
// const row = (rowFields) => {
//   return Object.entries(DATA_KEY_MAP).reduce((agg, curr) => {
//     agg[curr[0]] = agg[curr[1]  /*..xfer keys of sub objects to this object. explicit naming is fine...  prolly some easy es5 way to merge in fields? actually, this may be it already?  */]
//   }, {})
// }

/** For a given row's array of columns, this does any needed transforms to produce presentable result */
const buildReplacementMap = (rowFields) => {
  // loop the formal field specs and collect(/transform) this row's actual field vals as appropriate
  return Object.values(DATA_KEY_MAP).reduce((agg, curr) => {
    if (curr.replacementKey === null) return agg;
    agg[curr.replacementKey] = curr.valTransform ? curr.valTransform(rowFields[curr.dataGridIdx]) : rowFields[curr.dataGridIdx];
    return agg;
  }, {});
}

/** Check for complete row (all required fields are present) that's not had generation done before */
const shouldCreate = (rowFields) => {
  if (util.isNonBlank(rowFields[DATA_KEY_MAP.generated.dataGridIdx])) return false;
  return Object.values(DATA_KEY_MAP).every((spec) => !spec.required || util.isNonBlank(rowFields[spec.dataGridIdx]));
}

const createFilledDoc = (templateFileId, newDocName, replacementMap) => {
  const newDoc = newDocFromTemplate(templateFileId, newDocName);
  const newDocBody = newDoc.getBody();

  replaceKeys(newDocBody, replacementMap);
};

const replaceKeys = (docBody, replacementMap) => {
  // replaceText will replace all instances down the doc tree, which is fine for current purposes
  Object.entries(replacementMap).forEach(([key, val]) => {
    docBody.replaceText(`\{\{${key}\}\}`, val);
  })
};

/** Create and return a new doc to fill in, based on template */
const newDocFromTemplate = (templateFileId, newFileName) => {
  // Copy lands in same folder as template file. If file by name already exists, another of same name is created anyway.
  const file = DriveApp.getFileById(templateFileId).makeCopy(newFileName);  //yay, same folder (well, future state might be a subfolder for the company id?).  
  const newfileId = file.getId();
  
  return DocumentApp.openById(newfileId);
};

/** barring a better way to do this, this primes a test fxn that behaves like an actual handler of real sheet event.  */
const TEST_SheetEditEvent = () => {
  const testSheetId = PropertiesService.getScriptProperties().getProperty('TEST_SHEET_ID')
  const evt = {
    source: SpreadsheetApp.openById(testSheetId)
  };

  handleOnEdit(evt);
}


//-- general utils
const util = {
  accordianAdjustRange: (range, {addRows = 0, addColumns = 0}) => {
    return range.offset(addRows, addColumns, range.getNumRows() - addRows, range.getNumColumns() - addColumns);
  },

  /** Camel-case input string */
  camelize: (str) => str && str.split(' ').map(it => it[0].toUpperCase() + it.slice(1)).join(''),
  
  /** Check if string is non-null and non-empty/whitespace */
  isNonBlank: (str) => !!str?.toString().match(/\S+/)
};
