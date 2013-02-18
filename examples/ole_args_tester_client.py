#!/usr/local/bin/python
# -*- coding: utf-8 -*-
'''OLE_args_tester_client
How to use it on Windows 7/Vista:
1. register server
  1-1. Enable your PC's Administrator account.
  1-2. runas /user:administrator "python ole_args_tester_server.py"
2. python ole_args_tester_client.py
'''

import win32com.client

def main():
  cl = win32com.client.Dispatch('OLEArgsTester.Server')
  cl.Init()
  print cl.Test('a', 'b', 'c', 'd')
  print cl.GetSubName()
  print cl.SetSubName('xyz')
  print cl.GetSubName()
  print cl.subname
  cl.subname = 'ooo'
  print cl.GetSubName()
  cl.Quit()

if __name__ == '__main__':
  main()
