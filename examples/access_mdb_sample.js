var win32ole = require('win32ole');
win32ole.print('access_mdb_sample\n');
console.log(win32ole.version());
var path = require('path');
var cwd = path.join(win32ole.MODULEDIRNAME, '..');

var fs = require('fs');
var tmpdir = path.join(cwd, 'test/tmp');
if(!fs.existsSync(tmpdir)) fs.mkdirSync(tmpdir);
var outfile = path.join(tmpdir, 'access_mdb_sample.mdb');

var display_or_edit_all = function(rs, ed){
  var getRSvalue = function(rs, fieldname){
    return rs.Fields(fieldname).Value;
  }
  rs.MoveFirst();
/*
 In ParseUnaryExpression() < v8/src/parser.cc >
 v8::Object::ToBoolean() is called directly for unary operator '!'
 instead of v8::Object::valueOf()
 so NamedPropertyHandler will not be called
 Local<Boolean> ToBoolean(); // How to fake ? override v8::Value::ToBoolean
*/
//while(rs.Eof != true){ // It works. (without unary operator !)
//while(!rs.Eof){ // It does not work.
  while(!rs.Eof._){ // *** It works. oops!
    win32ole.print('id: ');
    win32ole.print(getRSvalue(rs, 'id'));
    win32ole.print(', c1: ');
    win32ole.print(getRSvalue(rs, 'c1'));
    win32ole.print(', c2: ');
    win32ole.print(getRSvalue(rs, 'c2'));
    win32ole.print(', c3: ');
    win32ole.print(getRSvalue(rs, 'c3'));
    win32ole.print('\n');
    if(ed){
      rs.Edit();
      var id = getRSvalue(rs, 'id');
      rs.Fields('c2').Value = id * 1000;
      rs.Update(); // rs.CancelUpdate();
    }
    rs.MoveNext();
  }
};

var adox_sample = function(filename){
  if(fs.existsSync(filename)) fs.unlinkSync(filename);
  console.log('creating mdb (by ADOX): "' + filename + '" ...');
  var dsn = 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=' + filename + ';';
  // dsn += ('Locale Identifier=' + CAT_LOCALE_(language));
  var db = win32ole.client.Dispatch('ADOX.Catalog');
  db.Create(dsn);
  var sql_create_table = 'create table testtbl (id autoincrement primary key,';
  sql_create_table += ' c1 varchar(255), c2 integer, c3 varchar(255));';
  var cn = db.ActiveConnection._; // ***
  cn.Execute(sql_create_table);
  for(var i = 11; i < 13; ++i){
    var sql_insert = "insert into testtbl (c1, c2, c3) values";
    sql_insert += " ('a(', " + i + ", ')z');";
    cn.Execute(sql_insert);
  }
  console.log('open RS');
  var rs = win32ole.client.Dispatch('ADODB.Recordset');
  rs.ActiveConnection = cn;
  var sql_select = 'select * from testtbl;';
  rs.Open(sql_select, cn, 1, 3); // adOpenKeyset, adLockOptimistic
  display_or_edit_all(rs, false);
  rs.Close();
  rs = null;
  console.log('RS released');
  cn.Close();
  cn = null;
  // db.Finalize(); and db.Close() db.Disconnect() db.Release() is wrong
  db = null;
  console.log('disconnected (ADOX)');
};

var ole_automation_sample = function(filename){
  console.log('open mdb (by OLE Automation): "' + filename + '" ...');
  var mdb = win32ole.client.Dispatch('Access.Application');
  mdb.Visible = true;
  // mdb.NewCurrentDatabase(filename); // to create new mdb
  mdb.OpenCurrentDatabase(filename); // to open exists mdb
  var cdb = mdb.CurrentDb();
  for(var i = 101; i < 103; ++i){
    var sql_insert = "insert into testtbl (c1, c2, c3) values";
    sql_insert += " ('p(', " + i + ", ')q');";
    cdb.Execute(sql_insert);
  }
  var sql_select = 'select * from testtbl;';
  var rs = cdb.OpenRecordset(sql_select);
  display_or_edit_all(rs, true);
  display_or_edit_all(rs, false);
  rs.Close();
  cdb.Close();
  mdb.CloseCurrentDatabase();
  mdb.Quit(1); // 1: acQuitSaveAll (default), 2: acQuitSaveNone
  cdb = null;
  mdb = null;
  console.log('disconnected (OLE Automation)');
};

try{
  adox_sample(outfile);
  win32ole.sleep(2000, true, true);
  ole_automation_sample(outfile);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
