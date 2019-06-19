# -*- coding: utf-8 -*-
import excelread
import requester
import os

path = os.getcwd()
if __name__=='__main__':
    print('read from '+path)
    manager = excelread.Manager('会計管理簿.xlsm')
    resp = requester.send(manager.get_json())
    print(resp)
    print(manager.get_json())