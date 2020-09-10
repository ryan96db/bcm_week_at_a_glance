#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Aug 21 11:48:53 2020

@author: ryandebose-boyd1
"""


import os
import pandas as pd
import random
from xlwt import Workbook

global length

#Creates blank excel spreadsheet with columns labeled
def blank():
    global wb
    global data
    wb = Workbook()
    data = wb.add_sheet('Data')
    data.write(0, 1, "Monday")
    data.write(0, 2, "Tuesday")
    data.write(0,3, "Wednesday")
    data.write(0, 4, "Thursday")
    data.write(0, 5, "Friday")
    
    hour = 8
    o_clock = ":00" #if index % 4 == 1
    quince = ":15" #if index % 4 == 2
    treinta = ":30" #f index % 4 == 3
    fifteen_til = ":45" #if index % == 0
    row = 1
    
    time_slot = ""
    for index in range(38):
        
        if index % 4 == 1:
            time_slot = str(hour) + o_clock
            data.write(row, 0, time_slot)
        elif index % 4 == 2:
            time_slot = str(hour) + quince
            data.write(row, 0, time_slot)
        elif index % 4 == 3:
            time_slot = str(hour) + treinta
            data.write(row, 0, time_slot)
        elif index % 4 == 0 and not(index == 0):
            time_slot = str(hour) + fifteen_til
            data.write(row, 0, time_slot)
            hour += 1
        
        
        row += 1
    
    
    wb.save("Calendar.xls")

blank()

#def submitToSpreadsheet(current_item):
    
    # data.write(x, 0, current_item.getKey())
    # data.write(x, 1, current_item.getName())
    # data.write(x, 2, current_item.getCategory())
    # data.write(x, 3, current_item.getInventory())
    # data.write(x, 4, current_item.getNote())

    