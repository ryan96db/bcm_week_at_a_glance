#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Aug 21 11:48:53 2020

@author: ryandebose-boyd1
"""
#1 hour last week
#Hours 9-9: 2
#Hours 9-10 1:40
#Hours 9-11 1:30 + 1 = 2:30
#Hours 9-13 2:30 
#Hours 9-17 1:15, :45 = 2:00
#11:40 hours so far
#Hours 9-18: 1:15

import fitz
#install?
import xlrd
#install?

import re

import os
import pandas as pd
import random
#import xlwt

import xlsxwriter
#install?

global length

global slots
slots = []

global wkbk


def createCourseObj(crs):
    
    wrkbk = xlrd.open_workbook("Catalog.xlsx")
    sheet = wrkbk.sheet_by_index(0)
    rows = sheet.nrows
    foundCourse = False
    courseObject = None
    for c in range(rows):
        current_course = sheet.cell_value(c, 0)
        current_course = current_course.strip()
        crs = crs.strip()
        if (current_course == crs):
            courseNumber = sheet.cell_value(c, 0)
            courseName = sheet.cell_value(c, 1)
            courseDays = sheet.cell_value(c, 2)
            courseTime = sheet.cell_value(c, 3)
            courseLoc = sheet.cell_value(c, 4)
            
            courseObject = Course(courseNumber, courseName,
                                  courseDays, courseTime,
                                  courseLoc)
            foundCourse = True
            break
    
    if foundCourse == True:
        return courseObject
    else:
        return courseObject
        
        
#Creates blank excel spreadsheet with columns labeled
def blank():
    global wb
    global data
    wb = xlsxwriter.Workbook('Calendar.xlsx')
    data = wb.add_worksheet()
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
    row = 0
    
    time_slot = ""
    for index in range(38):
        
        if index % 4 == 1:
            time_slot = str(hour) + o_clock
            data.write(row, 0, time_slot)
            
            slots.append(time_slot)
        elif index % 4 == 2:
            time_slot = str(hour) + quince
            data.write(row, 0, time_slot)
            
            slots.append(time_slot)
        elif index % 4 == 3:
            time_slot = str(hour) + treinta
            data.write(row, 0, time_slot)
            
            slots.append(time_slot)
        elif index % 4 == 0 and not(index == 0):
            time_slot = str(hour) + fifteen_til
            data.write(row, 0, time_slot)
            
            slots.append(time_slot)
            hour += 1
        
        if hour == 13:
            hour = 1
        
        
        row += 1
    
#Course object
class Course:
    def __init__(self, course_num, course_name, course_days, course_time, course_loc):
        self.course_num = course_num
        self.course_name = course_name
        self.course_days = course_days
        self.course_time = course_time
        self.course_loc = course_loc
        
    def getCourse_num(self):
        return self.course_num
        
    def getCourse_name(self):
        return self.course_name
        
    
    #Convert Day of week to be converted from string to number
    def getCourse_days(self):
        return self.course_days
        
    def getCourse_time(self):
        return self.course_time
        
    def getCourse_loc(self):
        return self.course_loc
        
        
#Template course object   
course1 = Course("GS-QC-6301", "Practical Introduction to Programming for Scientists",
                 "MF", "4:00-5:00 ", "N315")

#Converts meeting time to numbers for excel spreadsheet
def dayToNumber(crs):
    #When we get to parsing syllabi, we'll need to find way to convert Monday,
    #Wednesday, and Friday to string of weekdays represented as single letters,
    #So this function can work.
    dayIndexes = []
    days = crs.getCourse_days()
     
    for i in range(len(days)):
        x = days[i]
        x = x.upper()
        if x == "M":
            dayIndexes.append(1)
        elif x == "T":
            dayIndexes.append(2)
        elif x == "W":
            dayIndexes.append(3)
        elif x == "R":
            dayIndexes.append(4)
        elif x == "F":
            dayIndexes.append(5)

    return dayIndexes

#Converts begin/end meeting times to corresponding timeslots in excel
def meetToNumber(crs):
    
    t =  str(crs.getCourse_time())
    times = t.split('-')
    begin_time = times[0].strip()
    begin_time = begin_time.lstrip("0")
    
    end_time = times[1].strip()
    
    begin_excel_index = slots.index(begin_time) + 1
    end_excel_index = slots.index(end_time) + 1
    
    excelIndexes = [begin_excel_index, end_excel_index]
    
    return excelIndexes

#Takes individual course number (for now) and adds it to excel file
def courseToTimeSlots(crs):
    dayindeces = dayToNumber(crs)
    excelindeces = meetToNumber(crs)
    begin = excelindeces[0]
    end = excelindeces[1]
    
    #align: vcenter
    cell_format = wb.add_format({'text_wrap': True})
    merge_format = wb.add_format({'align': 'center', 
    'text_wrap': True
    })
    
    
    big_string = crs.getCourse_num() + "\n" + crs.getCourse_time() + "\n" + crs.getCourse_loc()
                
    for day in dayindeces:
        for v in range(begin, end+1):
            if v == begin:
                #data.write(v, day, big_string, cell_format)
                data.merge_range(begin, day, end, day, big_string, merge_format)
    
           
def find(ident, courseSeq):
    try:
        return courseSeq.index(ident)
    except:
        return -1

weekIdentifiers = ["M,W,F ", "T R", "M W F", "M,,", "T,,", "M W", "MTWR", "W ", "F ", "R "]

#ex_course = "GS-CC-6401"


    
    
    


    
   

     
    
    
    
    
    
    






        
        
        
        
        



    