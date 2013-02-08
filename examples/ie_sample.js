var win32ole = require('win32ole');
win32ole.print('ie_sample\n');

var ie_sample = function(uris){
  var ie = win32ole.client.Dispatch('InternetExplorer.Application');
  ie.set('Visible', true);
  for(var i = 0; i < uris.length; ++i){
    console.log(uris[i]);
    ie.call('Navigate', [uris[i]]);
    win32ole.sleep(15000, true, true);
  }
  ie.call('Quit');
};

try{
  ie_sample([
    'http://www.google.com/',
    'http://www.mozilla.org/',
    'http://nodejs.org/']);
}catch(e){
  console.log('*** exception cached ***\n' + e);
}
