#!/usr/local/bin/python
# -*- coding: utf-8 -*-
'''OLE_args_tester_server
How to use it on Windows 7/Vista:
1. register server
  1-1. Enable your PC's Administrator account.
  1-2. runas /user:administrator "python ole_args_tester_server.py"
2. python ole_args_tester_client.py

optional ( to unregister this server ) runas administrator
3. python ole_args_tester_server.py --unregister
'''

class OLEArgsTesterServer:
  _public_methods_ = ['Init', 'Test', 'GetSubName', 'SetSubName', 'Quit']
  _public_attrs_ = ['subname']
  _reg_progid_ = 'OLEArgsTester.Server'
  # NEVER copy the ProgID / CLSID
  # Use "print pythoncom.CreateGuid()" to make a new one.
  _reg_clsid_ = '{4866809D-FCD5-498E-93D5-FA78949CC1DA}'

  def Init(self):
    self.name = OLEArgsTesterServer._reg_progid_
    self.subname = unicode('%s(%s)' % ('subname', self.name))

  def Test(self, a0, a1=None, a2=None, a3=None):
    s = [unicode(a0), unicode(a1), unicode(a2), unicode(a3)]
    return u'test: %s' % u', '.join(s)

  def GetSubName(self, a0=None):
    return self.subname

  def SetSubName(self, sn):
    self.subname = u'%s(%s)' % (unicode(sn), unicode(self.name))
    return None

  def Quit(self):
    return None

if __name__ == '__main__':
  print 'Registering ole_args_tester_server...'
  import win32com.server.register
  win32com.server.register.UseCommandLine(OLEArgsTesterServer)
