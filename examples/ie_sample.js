var win32ole = require('win32ole');
win32ole.print('ie_sample\n');

var sleep = function(milliSeconds){
  var startTime = new Date().getTime();
  while(new Date().getTime() < startTime + milliSeconds);
}

var ie_sample = function(uris){
  var ie = win32ole.client.Dispatch('InternetExplorer.Application', 'C');
  ie.set('Visible', true);
  for(var i = 0; i < uris.length; ++i){
    console.log(uris[i]);
    ie.call('Navigate', [uris[i]]);
    console.log('waiting 15 seconds...');
    sleep(15000);
  }
  ie.call('Quit');
};

win32ole.client = new win32ole.Client;
try{
  ie_sample([
    'http://www.google.com/',
    'http://www.mozilla.org/',
    'http://nodejs.org']);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
win32ole.client.Finalize(); // must be called (version 0.0.x)
