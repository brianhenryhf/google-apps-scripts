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

// Each value in the below is a spec for a field in the data sheet:
// - dataColIdx is 0-based column index in data sheet.
// - replacementKey is key in template, to be replaced by a val. null replacement key means not a field involved in 
//   replacing.  
// - required is requiredness in data sheet - if all required fields are present for a row, merge doc can be generated
//   for that row
// - valTransform is optional fxn to transform val before replacing in destination doc
const FIELD_SPECS = {
  generated: { 
    dataColIdx: 0, 
    replacementKey: null, 
    required: false 
  },
  customerId: { 
    dataColIdx: 1, 
    replacementKey: null,
    required: true 
  },
  invoiceNum: { 
    dataColIdx: 2, 
    replacementKey: null, 
    required: true 
  },
  invoiceId: { 
    dataColIdx: 3, 
    replacementKey: 'INVC_ID', 
    required: true 
  },
  invoiceDate: { 
    dataColIdx: 4, 
    replacementKey: 'INVC_DT', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  workPeriodStart: { 
    dataColIdx: 5, 
    replacementKey: 'PERIOD_ST_INCL', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  workPeriodEnd: { 
    dataColIdx: 6, 
    replacementKey: 'PERIOD_END_INCL', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  },
  totalHours: { 
    dataColIdx: 7, 
    replacementKey: 'TOT_HRS', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { minimumFractionDigits: 1 }).format(val), 
    required: true 
  },
  hourlyRate: { 
    dataColIdx: 8, 
    replacementKey: 'HRLY_RT', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val), 
    required: true 
  },
  totalAmtDue: { 
    dataColIdx: 9, 
    replacementKey: 'TOT_DUE', 
    valTransform: (val) => new Intl.NumberFormat('en-US', { style: 'currency', currency: 'USD' }).format(val), 
    required: true 
  },
  templateFileId: { 
    dataColIdx: 10, 
    replacementKey: null, 
    required: true 
  },
  payerInfo: { 
    dataColIdx: 11, 
    replacementKey: 'PAYER_INFO', 
    required: true 
  },
  payeeInfo: { 
    dataColIdx: 12, 
    replacementKey: 'PAYEE_INFO', 
    required: true 
  },
  lineItemDesc: { 
    dataColIdx: 13, 
    replacementKey: 'LI_DESC', 
    required: true 
  },
  invoiceNote: { 
    dataColIdx: 14, 
    replacementKey: 'INV_NOTE', 
    required: false 
  },
  terms: { 
    dataColIdx: 15, 
    replacementKey: 'TERMS', 
    required: true 
  },
  dueDate: { 
    dataColIdx: 16, 
    replacementKey: 'DUE_DT', 
    valTransform: (val) => val.toLocaleDateString(), 
    required: true 
  }
};

/** Trigger handler for this script. */
const handleOnEdit = (evt) => {
  const ss = evt.source;
  const dataSheet = ss.getSheets()[0];

  processSheet(dataSheet);
};

/** Main workhorse of script */
const processSheet = (sheet) => {
  const contentRange = sheet.getDataRange();
  // trim off the header row, so row 0 starts the data
  const dataRange = util.accordianAdjustRange(contentRange, { addRows: 1 });
  const data = dataRange.getValues();

  data.forEach((row, idx) => {
    const rowFields = wrapRowData(FIELD_SPECS, row);

    if (!shouldCreate(rowFields)) return;
    invoiceReplacements = buildReplacementMap(rowFields);

    // historical naming convention - not ideal that we reference specific fields here and custom transform.
    // TODO a case could be made that this filename should be written to source sheet for recordkeeping. Consider this.
    const customerIdCamel = util.camelize(rowFields.valForName('customerId'));
    const invoiceNumber = rowFields.valForName('invoiceNum')?.toString().padStart(2, '0');
    const fileName = `${customerIdCamel}Invoice_${invoiceNumber}`;

    createFilledDoc(rowFields.valForName('templateFileId'), fileName, invoiceReplacements);

    //mark it done - using a marker column instead of checking for row count changes lets us re-generate from rows for testing or whatever reason we want.
    const generatedCell = dataRange.offset(idx, FIELD_SPECS.generated.dataColIdx, 1, 1);
    generatedCell.setValue('X')
  });
}

/** Wrapper to make it a bit more clean/grokkable to access field data by field name or data index */
const wrapRowData = (fieldSpecs, rowFields) => {
  return {
    valForIndex: function(idx) {
      return rowFields[idx];
    },
    valForName: function(fieldName) {
      return this.valForIndex(fieldSpecs[fieldName].dataColIdx);
    }
  }
}

/**
 * For a given row's array of columns, this does any needed transforms to produce presentable result in map of 
 * replacement key to display val. 
 */
const buildReplacementMap = (rowFields) => {
  const identityFn = (x) => { return x; };

  // loop the formal field specs and collect(/transform, if spec'd) this row's actual field vals as appropriate
  return Object.values(FIELD_SPECS).reduce((agg, curr) => {
    if (curr.replacementKey === null) return agg;
    agg[curr.replacementKey] = (curr.valTransform || identityFn)(rowFields.valForIndex(curr.dataColIdx));
    return agg;
  }, {});
}

/** Check for complete row (all required fields are present) that's not had generation done before */
const shouldCreate = (rowFields) => {
  if (util.isNonBlank(rowFields.valForName('generated'))) return false;
  return Object.values(FIELD_SPECS).every(
    (fieldSpec) => !fieldSpec.required || util.isNonBlank(rowFields.valForIndex(fieldSpec.dataColIdx))
  );
}

const createFilledDoc = (templateFileId, newDocName, replacementMap) => {
  const newDoc = newDocFromTemplate(templateFileId, newDocName);
  const newDocBody = newDoc.getBody();

  replaceKeys(newDocBody, replacementMap);
};

const replaceKeys = (docBody, replacementMap) => {
  // #replaceText will replace all instances down the doc tree, which is fine for current purposes
  Object.entries(replacementMap).forEach(([key, val]) => {
    docBody.replaceText(`\{\{${key}\}\}`, val);
  })
};

/** Create and return a new doc to fill in, based on template */
const newDocFromTemplate = (templateFileId, newFileName) => {
  // Copy lands in same folder as template file. If file by name already exists, another of same name is created.
  const file = DriveApp.getFileById(templateFileId).makeCopy(newFileName);  //yay, same folder (well, future state might be a subfolder for the company id?).  
  const newfileId = file.getId();
  
  return DocumentApp.openById(newfileId);
};

/** 
 * Barring a better way to do this, this primes a test fxn that behaves like an actual handler of real sheet event. 
 * Allows for quick iterating when debugging.
 */
const TEST_SheetEditEvent = () => {
  // set this property up for script beforehand, with spreadsheet file ID
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
