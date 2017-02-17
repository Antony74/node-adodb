node-adodb
===========
>一个用 Node.js 实现的 ADODB 协议。

>[![NPM Version][npm-image]][npm-url] [![Download Status][download-image]][npm-url] [![Dependencies][david-image]][david-url]

###安装
```
$ npm install node-adodb
```

###使用示例:
```js
var ADODB = require('node-adodb'),
  connection = ADODB.open('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=node-adodb.mdb;');

// 全局调试开关，默认关闭
ADODB.debug = true;
// 编码设定
ADODB.encoding = 'gbk';

// 不带返回的查询
connection
  .execute('INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Newton", "Male", 25)')
  .on('done', function (data){
    console.log('Result:'.green.bold, JSON.stringify(data, null, '  ').bold);
  })
  .on('fail', function (data){
    // TODO 逻辑处理
  });
  
// 带返回标识的查询
connection
  .executeScalar(
    'INSERT INTO Users(UserName, UserSex, UserAge) VALUES ("Newton", "Male", 25)',
    'SELECT @@Identity AS id'
  )
  .on('done', function (data){
    console.log('Result:'.green.bold, JSON.stringify(data, null, '  ').bold);
  })
  .on('fail', function (data){
    // TODO 逻辑处理
  });

// 带返回的查询
connection
  .query('SELECT * FROM Users')
  .on('done', function (data){
    console.log('Result:'.green.bold, JSON.stringify(data, null, '  ').bold);
  })
  .on('fail', function (data){
    // TODO 逻辑处理
  });
```

###接口文档:
`ADODB.debug`
>全局调试开关。

`ADODB.encoding`
>全局默认编码设置。

`ADODB.query(sql)`
>执行有返回值的SQL语句。

`ADODB.execute(sql)`
>执行无返回值的SQL语句。

`ADODB.executeScalar(sql, scalar)`
>执行带返回标识的SQL语句。

`ADODB.open(connection[, encoding])`
>编码设置为可选参数，可以用`ADODB.encoding`进行全局设置。

###扩展:
>该插件理论支持 Windows 平台下所有支持 ADODB 连接的数据库，只需要更改数据库连接字符串即可实现操作！

###注意:
>该插件需要系统支持 Microsoft.Jet.OLEDB.4.0，对于 Windows XP SP2 以上系统默认支持，其它需要自己升级，具体操作过程请参考：
[如何获取 Microsoft Jet 4.0 数据库引擎的最新 Service Pack](http://support.microsoft.com/default.aspx?scid=kb;zh-CN;239114)

[npm-url]: https://www.npmjs.org/package/node-adodb
[npm-image]: https://img.shields.io/npm/v/node-adodb.svg?style=flat-square
[download-image]: https://img.shields.io/npm/dm/node-adodb.svg?style=flat-square
[david-url]: https://david-dm.org/Nuintun/node-adodb
[david-image]: https://img.shields.io/david/nuintun/node-adodb.svg?style=flat-square