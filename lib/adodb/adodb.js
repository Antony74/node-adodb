import { stdout, fillRecords, fillSchema } from './utils';

/**
 * Free connection
 *
 * @param {Object} object
 */
function free(object) {
  object.State && object.Close();
}

/**
 * ADODB
 */
export var ADODB = {
  execute: function(params) {
    var recordset;
    var result = { valid: true };
    var scalarMode = !!params.scalar;
    var connection = new ActiveXObject('ADODB.Connection');

    if (scalarMode) {
      recordset = new ActiveXObject('ADODB.Recordset');
    }

    // Set CursorLocation
    connection.CursorLocation = 3;

    try {
      // Open
      connection.Open(params.connection);
      // Execute
      connection.Execute(params.sql);

      // Scalar
      if (scalarMode) {
        recordset.Open(params.scalar, connection, 0, 1);

        // fill records
        result.records = fillRecords(recordset)
      }
    } catch (e) {
      result.valid = false;
      result.message = e.message;
    }

    // Write data
    stdout(result);

    // Close database
    free(connection);

    if (scalarMode) {
      free(recordset);
    }
  },
  query: function(params) {
    var data;
    var descrMode = !!params.schema;
    var connection = new ActiveXObject('ADODB.Connection');
    var recordset = new ActiveXObject('ADODB.Recordset');
    var result = { valid: true };

    // Set CursorLocation
    connection.CursorLocation = 2; // adUseServer

    try {
      // Open
      connection.Open(params.connection);
      // Query
      recordset.Open(params.sql, connection, 0, 1);

      // fill records
      result.records = fillRecords(recordset);

      // fill schema
      if (descrMode) {
        result.schema = fillSchema(recordset);
      }
    } catch (e) {
      result.valid = false;
      result.message = e.message;
    }

    // Write data
    stdout(result);

    // Close database
    free(recordset);
    free(connection);
  },

  openSchema: function(params) {
    var data;
    var connection = new ActiveXObject('ADODB.Connection');
    var result = { valid: true };

    // Set CursorLocation
    connection.CursorLocation = 2; // adUseServer

    try {
      // Open
      connection.Open(params.connection);
      // OpenSchema
      if (params.schemaId) {
        recordset = connection.OpenSchema(params.queryType, params.criteria ? params.criteria : null, params.schemaId);
      } else {
        recordset = connection.OpenSchema(params.queryType, params.criteria ? params.criteria : null);
      }

      // fill records
      result.records = fillRecords(recordset);

    } catch (e) {
      result.valid = false;
      result.message = e.message;
    }

    // Write data
    stdout(result);

    // Close database
    free(recordset);
    free(connection);
  }
};
