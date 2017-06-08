import { JSON } from './json';

/**
 * stdout
 *
 * @param data
 */
export function stdout(data) {
  try {
    WScript.StdOut.Write(encodeURI(JSON.stringify(data)));
  } catch (e) {
    WScript.StdOut.Write(encodeURI(JSON.stringify({
      valid: false,
      message: e.message
    })));
  }
}

/**
 * fill records array
 *
 * @param recordset
 * @returns {Array}
 */
export function fillRecords(recordset) {
  var item;
  var field;
  var records = [];
  var count = recordset.Fields.Count;

  // not empty
  if (!recordset.BOF || !recordset.EOF) {
    recordset.MoveFirst();

    while (!recordset.EOF) {
      item = {};

      for (var i = 0; i < count; i++) {
        field = recordset.Fields.Item(i);

        if (typeof(field.value) === 'date') {
          var date = new Date(field.value);
          // ADO has given us a UTC date but JScript assumes it's a local timezone date.  Correct the date.
          date = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate(), date.getHours(), date.getMinutes(), date.getSeconds(), date.getMilliseconds()));

          item[field.name] = date;
        } else {
          item[field.name] = field.value;
        }
      }

      records.push(item);
      recordset.MoveNext();
    }
  }

  // return records
  return records;
}
