import { stdout, fillRecords } from './utils';

/**
 * Free connection
 *
 * @param {Object} object
 */
function free(object) {
  object.State && object.Close();
}

var DataTypeEnum = {
  0x2000: 'AdArray',
  20:     'adBigInt',
  128:    'adBinary',
  11:     'adBoolean',
  8:      'adBSTR',
  136:    'adChapter',
  129:    'adChar',
  6:      'adCurrency',
  7:      'adDate',
  133:    'adDBDate',
  134:    'adDBTime',
  135:    'adDBTimeStamp',
  14:     'adDecimal',
  5:      'adDouble',
  0:      'adEmpty',
  10:     'adError',
  64:     'adFileTime',
  72:     'adGUID',
  9:      'adIDispatch',
  3:      'adInteger',
  13:     'adIUnknown',
  205:    'adLongVarBinary',
  201:    'adLongVarChar',
  203:    'adLongVarWChar',
  131:    'adNumeric',
  138:    'adPropVariant',
  4:      'adSingle',
  2:      'adSmallInt',
  16:     'adTinyInt',
  21:     'adUnsignedBigInt',
  19:     'adUnsignedInt',
  18:     'adUnsignedSmallInt',
  17:     'adUnsignedTinyInt',
  132:    'adUserDefined',
  204:    'adVarBinary',
  200:    'adVarChar',
  12:     'adVariant',
  139:    'adVarNumeric',
  202:    'adVarWChar',
  130:    'adWChar'
};

var FieldAttributeEnum = {
  0x1000:     'adFldCacheDeferred',
  0x10:       'adFldFixed',
  0x2000:     'adFldIsChapter',
  0x40000:    'adFldIsCollection',
  0x8000:     'adFldKeyColumn',
  0x20000:    'adFldIsDefaultStream',
  0x20:       'adFldIsNullable',
  0x10000:    'adFldIsRowURL',
  0x80:       'adFldLong',
  0x40:       'adFldMayBeNull',
  0x2:        'adFldMayDefer',
  0x4000:     'adFldNegativeScale',
  0x100:      'adFldRowID',
  0x200:      'adFldRowVersion',
  0x8:        'adFldUnknownUpdatable',
  0xFFFFFFFF: 'adFldUnspecified',
  0x4:        'adFldUpdatable'
};

function resolveEnum(oEnum, num) {
  stringValue = oEnum[num];
  return stringValue ? stringValue : num;
}

function resolveBitfieldEnum(oEnum, num) {
  var flags = {};
  for (var key in FieldAttributeEnum) {
    var value = oEnum[key];
    if (key & num) {
      flags[value] = true;
      num -= key;
    }
  }

  if (num) {
    flags[num] = true;
  }

  return flags;
}

/**
 * ADODB
 */
export var ADODB = {
  execute: function(params) {
    var valid = true;
    var message = params.sql + ' success';
    var connection = new ActiveXObject('ADODB.Connection');

    // Set CursorLocation
    connection.CursorLocation = 3;

    try {
      // Open
      connection.Open(params.connection);
      // Execute
      connection.Execute(params.sql);
    } catch (e) {
      valid = false;
      message = e.message;
    } finally {
      // Close database
      free(connection);
    }

    // Write data
    stdout({
      valid: valid,
      message: message
    });
  },
  scalar: function(params) {
    var records;
    var valid = true;
    var message = params.sql + ' / ' + params.scalar + ' success';
    var connection = new ActiveXObject('ADODB.Connection');
    var recordset = new ActiveXObject('ADODB.Recordset');

    // Set CursorLocation
    connection.CursorLocation = 3;

    try {
      // Open
      connection.Open(params.connection);
      // Execute
      connection.Execute(params.sql);
      // Scalar
      recordset.Open(params.scalar, connection, 0, 1);
      records = fillRecords(recordset);

      // Empty
      if (records.length === 0) {
        message = 'The scalar recordset is empty';
      }
    } catch (e) {
      valid = false;
      message = e.message;
    } finally {
      // Close database
      free(recordset);
      free(connection);
    }

    // Write data
    stdout({
      valid: valid,
      message: message,
      records: records
    });
  },
  query: function(params) {
    var records;
    var valid = true;
    var message = params.sql + ' success';
    var connection = new ActiveXObject('ADODB.Connection');
    var recordset = new ActiveXObject('ADODB.Recordset');

    // Set CursorLocation
    if (!params.fields) {
      connection.CursorLocation = 3;
    }

    try {
      // Open
      connection.Open(params.connection);
      // Query
      recordset.Open(params.sql, connection, 0, 1);
      records = fillRecords(recordset);

      // Fields
      var fields = undefined;

      if (params.fields) {

        fields = {};
        var count = recordset.Fields.Count;

        for (var i = 0; i < count; i++) {
          field = recordset.Fields.Item(i);

          var adoProperties = field.Properties;
          var propCount = adoProperties.Count;
          var properties = {};

          for (var j = 0; j < propCount; j++) {
            prop = adoProperties.Item(j);
            properties[prop.Name] = {
//              Attributes: prop.Attributes,
              Type: resolveEnum(DataTypeEnum, prop.Type),
              Value: prop.Value
            }
          }

          fields[field.name] = {
            Attributes: resolveBitfieldEnum(FieldAttributeEnum, field.Attributes),
            DefinedSize: field.DefinedSize,
            NumericScale: field.NumericScale,
            Precision: field.Precision,
            Properties: properties,
            Type: resolveEnum(DataTypeEnum, field.Type)
          };
        }
      }

      // Empty
      if (records.length === 0) {
        message = 'The recordset is empty';
      }
    } catch (e) {
      valid = false;
      message = e.message;
    } finally {
      // Close database
      free(recordset);
      free(connection);
    }

    // Write data
    stdout({
      valid: valid,
      message: message,
      records: records,
      fields: fields
    });
  }
};
