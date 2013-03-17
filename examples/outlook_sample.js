var win32ole = require('win32ole');
win32ole.print('outlook_sample\n');

var outlook_sample = function(){
  var ol = win32ole.client.Dispatch('Outlook.Application');
  // ol.Visible = true;
  var ns = ol.GetNameSpace('MAPI');
  win32ole.print('mail:\n');
  var frcv = ns.GetDefaultFolder(6); // receive mail box tray
  var items = frcv.Items._; // ***
  var count = items.Count;
  for(var n = 1; n <= count; ++n){
    win32ole.print(' ' + n + ' : ');
    var i = items.Item(n);
    win32ole.printACP(i.Subject);
    win32ole.print('\n');
  }
  win32ole.print('schedule:\n');
  var fcal = ns.GetDefaultFolder(9); // olFolderCalendar (schedule)
  if(true){
    var apnt = ol.CreateItem(1); // olAppointmentItem
    var t = new Date();
    apnt.Start = t;
    t.setTime(t.getTime() + 30 * 60 * 1000); // + 30 minutes
    apnt.End = t;
    apnt.Subject = 'TTEESSTT';
    apnt.Body = 'bodybodybody';
    apnt.Location = 'node';
    apnt.Sensitivity = 0; // olNormal
    apnt.ReminderSet = true;
    apnt.ReminderMinutesBeforeStart = 120; // minutes
    apnt.Save();
  }
  // var apnts = fcal.Items._; // ***
  var apnts = fcal.Items.restrict('[Start] >= "26/02/2013 09:30"');
  var acnt = apnts.Count;
  for(var n = 1; n <= acnt; ++n){
    win32ole.print(' ' + n + ' : ');
    /*
    var apnt = apnts.Item(n);
    var ptn = apnt.GetRecurrencePattern();
    var i = ptn.GetOccurrence(new Date(2013, 2, 13, 13, 30, 0)); // month - 1
    */
    var i = apnts.Item(n);
    win32ole.print('DateTime: ( from ');
    win32ole.printACP(i.Start._);
    win32ole.print(' to ');
    win32ole.printACP(i.End._);
    win32ole.print(' )\n');
    win32ole.printACP('  Subject: ' + i.Subject);
    win32ole.print('\n');
    win32ole.printACP('  Body: ' + i.Body);
    win32ole.print('\n');
    win32ole.printACP('  Location: ' + i.Location);
    win32ole.print('\n');
  }
  // ol.Quit();
};

try{
  outlook_sample();
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
