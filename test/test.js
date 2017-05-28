var fs = require('fs');
var path = require('path');
var arch = require('arch');
var ADODB = require('../index');
var expect = require('expect.js');

var source = path.join(__dirname, 'node-adodb.mdb');
var mdb = fs.readFileSync(path.join(__dirname, '../examples/node-adodb.mdb'));

fs.writeFileSync(source, mdb);

// variable declaration
var x64 = arch() === 'x64';
var sysroot = process.env['systemroot'] || process.env['windir'];
var cscript = path.join(sysroot, x64 ? 'SysWOW64' : 'System32', 'cscript.exe');

if (fs.existsSync(cscript) && fs.existsSync(source)) {
  console.log('Use:', cscript);
  console.log('Database:', source);

  describe('ADODB', function() {
    // variable declaration
    var connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + source + ';');

    var query = connection.query('SELECT * FROM Users');

    // noop function
    function fn() {}

    // coveralls cover
    query
      .on('done', fn)
      .off('done', fn)
      .off('done')
      .off()
      .once('done', fn);

    query.on('custom', fn);
    query.emit('custom');
    query.emit('custom', 1);
    query.emit('custom', 1, 2);
    query.emit('custom', 1, 2, 3);
    query.emit('custom', 1, 2, 3, 4);

    it('query', function(next) {
      connection
        .query('SELECT * FROM Users')
        .on('done', function(data, message) {
          expect(data).to.eql(
            [
              {
                UserId: 1,
                UserName: "Nuintun",
                UserSex: "Male",
                UserAge: 25
              },
              {
                UserId: 2,
                UserName: "Angela",
                UserSex: "Female",
                UserAge: 23
              },
              {
                UserId: 3,
                UserName: "Newton",
                UserSex: "Male",
                UserAge: 25
              }
            ]
          );
          expect(message).to.eql('SELECT * FROM Users success');

          next();
        }).on('fail', function(error) {
          next(error);
        });
    });

    it('queryWithMetadata', function(next) {
      connection
        .query('SELECT * FROM Users', {fields: true})
        .on('done', function(data, message, extras) {
          expect(data).to.eql(
            [
              {
                UserId: 1,
                UserName: "Nuintun",
                UserSex: "Male",
                UserAge: 25
              },
              {
                UserId: 2,
                UserName: "Angela",
                UserSex: "Female",
                UserAge: 23
              },
              {
                UserId: 3,
                UserName: "Newton",
                UserSex: "Male",
                UserAge: 25
              }
            ]
          );
          expect(message).to.eql('SELECT * FROM Users success');
          expect(extras).to.eql({
            "fields": {
              "UserId": {
                "Attributes": {
                  "adFldFixed": true,
                  "adFldMayBeNull": true,
                  "adFldMayDefer": true,
                  "adFldUnknownUpdatable": true
                },
                "DefinedSize": 4,
                "NumericScale": 255,
                "Precision": 10,
                "Properties": {
                  "BASECOLUMNNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASETABLENAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASECATALOGNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASESCHEMANAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "KEYCOLUMN": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "ISAUTOINCREMENT": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "RELATIONCONDITIONS": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "CALCULATIONINFO": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "OPTIMIZE": {
                    "Type": "adBoolean",
                    "Value": false
                  }
                },
                "Type": "adInteger"
              },
              "UserName": {
                "Attributes": {
                  "adFldIsNullable": true,
                  "adFldMayBeNull": true,
                  "adFldMayDefer": true,
                  "adFldUnknownUpdatable": true
                },
                "DefinedSize": 100,
                "NumericScale": 255,
                "Precision": 255,
                "Properties": {
                  "BASECOLUMNNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASETABLENAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASECATALOGNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASESCHEMANAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "KEYCOLUMN": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "ISAUTOINCREMENT": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "RELATIONCONDITIONS": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "CALCULATIONINFO": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "OPTIMIZE": {
                    "Type": "adBoolean",
                    "Value": false
                  }
                },
                "Type": "adVarWChar"
              },
              "UserSex": {
                "Attributes": {
                  "adFldIsNullable": true,
                  "adFldMayBeNull": true,
                  "adFldMayDefer": true,
                  "adFldUnknownUpdatable": true
                },
                "DefinedSize": 255,
                "NumericScale": 255,
                "Precision": 255,
                "Properties": {
                  "BASECOLUMNNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASETABLENAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASECATALOGNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASESCHEMANAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "KEYCOLUMN": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "ISAUTOINCREMENT": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "RELATIONCONDITIONS": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "CALCULATIONINFO": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "OPTIMIZE": {
                    "Type": "adBoolean",
                    "Value": false
                  }
                },
                "Type": "adVarWChar"
              },
              "UserAge": {
                "Attributes": {
                  "adFldFixed": true,
                  "adFldIsNullable": true,
                  "adFldMayBeNull": true,
                  "adFldMayDefer": true,
                  "adFldUnknownUpdatable": true
                },
                "DefinedSize": 4,
                "NumericScale": 255,
                "Precision": 10,
                "Properties": {
                  "BASECOLUMNNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASETABLENAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASECATALOGNAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "BASESCHEMANAME": {
                    "Type": "adVarWChar",
                    "Value": null
                  },
                  "KEYCOLUMN": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "ISAUTOINCREMENT": {
                    "Type": "adBoolean",
                    "Value": false
                  },
                  "RELATIONCONDITIONS": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "CALCULATIONINFO": {
                    "Type": "adVarBinary",
                    "Value": null
                  },
                  "OPTIMIZE": {
                    "Type": "adBoolean",
                    "Value": false
                  }
                },
                "Type": "adInteger"
              }
            }
          });

          next();
        }).on('fail', function(error) {
          next(error);
        });
    });

    it('execute', function(next) {
      connection
        .execute('INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Nuintun", "Male", 25)')
        .on('done', function(data, message) {
          expect(data).to.eql([]);
          expect(message).to.eql('INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Nuintun", "Male", 25) success');

          next();
        }).on('fail', function(error) {
          next(error);
        });
    });

    it('scalar', function(next) {
      connection
        .execute('INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Alice", "Female", 25)', 'SELECT @@Identity AS id')
        .on('done', function(data, message) {
          expect(data).to.eql([{ id: 5 }]);
          expect(message).to.eql('INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Alice", "Female", 25) / SELECT @@Identity AS id success');

          next();
        }).on('fail', function(error) {
          next(error);
        });
    });
  });
} else {
  console.log('This OS not support node-adodb.');
}
