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
#Hours 9-20: 1:00
#Hours 9-21: 1:15 + 1:30 
#16:40 hours so far
#Hours 9-23: 1:00 + 1:45


#install
import fitz
#install?
import xlrd
#install?

import re

import os.path
import pandas as pd
import random
#import xlwt

import xlsxwriter
#install?

global length

global slots
slots = []

global wkbk

def locationInfo(course):
    cat_building_abbrev = {"A": "Cullen", "B": "Cullen", "CNRC": "Children's Nutritional Reasearch Center",
                       "D": "Jewish Institute for Medical Research", "E": "MD Anderson Cancer Center",
                       "M": "Michael E. DeBakey Center (BCM)", "N": "Alkek Tower", "R": "Alkek Bldg for Biomedical Research (ABBR)", 
                       "S": "Smith Research Wing", "T": "Ben Taub Research Center", "TXFC": "Texas Children’s Feigin Center"}
    loc = course.getCourse_loc()
    
    loc = loc.strip()

    nri = "NRI"
    
    bcm_map_building_abbrev = {"Alkek Bldg for Biomedical Research (ABBR)": "ABBR","Alkek Tower":"ALKT","Ben Taub Research Center":"BCMT",
                               "Children's Nutritional Reasearch Center":"CNRC","Cullen":"BCMC", "Texas Children’s Feigin Center":"TXFC",
                               "NRI":"NRIB", "Jewish Institute for Medical Research":"BCMD",
                                "MD Anderson Cancer Center":"MDAC", "Michael E. DeBakey Center (BCM)":"BCMM",
                                "Rice University":"RICE", "Smith Research Wing":"BCMS"  }
    
    
    for key in cat_building_abbrev:
        if key in loc:
            if not (loc[1] == "."):
                
                bcm_map_key = cat_building_abbrev[key]
                web_map = bcm_map_building_abbrev[bcm_map_key]
            else:
                web_map = bcm_map_building_abbrev[nri]
    
    web_map_link = "https://www.bcm.edu/map/#!" + web_map
    
    return web_map_link
    
def manualCourseArray():
    cont = True
    course_array = []
    while cont == True:
        courseInp = input("Please enter course numbers (press q to quit): ")
        
        if courseInp.upper() == "Q":
            break
        
        courseObj = None
        courseObj = createCourseObj(courseInp)

        if not (courseObj == None):

            course_array.append(courseInp)
            
        else:
            print("Course name entered was not found, so course was not added.")
        
    return course_array

def syllabiArray():
    cont = True
    pdf_syl_array = []
    while cont == True:
        pdfInp = input("Please enter syllabi file names, including file extension (press q to quit): ")
        
        if pdfInp.upper() == "Q":
            break
        
        syllabus_exists = False
        
        if os.path.isfile(pdfInp):
            pdf_syl_array.append(pdfInp)
        elif syllabus_exists == False:
            print("File " + pdfInp + " not found. ")

    return pdf_syl_array
 
def courseFromPDF(pdfFileName):
    pdfObj = fitz.open(pdfFileName)
    
    pages = pdfObj.pageCount
    
    wrkbk = xlrd.open_workbook("Catalog.xlsx")
    sheet = wrkbk.sheet_by_index(0)
    
    rows = sheet.nrows
    courseNameFound = False
    
    
    for page in range(pages):
        pageObject = pdfObj.loadPage(page)
        txt = pageObject.getText("text")
        c = 0
        while c < rows and courseNameFound == False:
            current_course = sheet.cell_value(c, 0)
            current_course = current_course.strip()
            if current_course in txt:
                courseNameFound = True
            else:
                c+= 1
    
    course = createCourseObj(current_course)
    courseToTimeSlots(course)


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
    for index in range(48):
        
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

    web_link = locationInfo(crs)
    location = crs.getCourse_loc()

    big_string = crs.getCourse_num() + "\n" + crs.getCourse_time() + "\n" + crs.getCourse_loc()
                
    for day in dayindeces:
        for v in range(begin, end+1):
            if v == begin:
                #data.write(v, day, big_string, cell_format)
                data.write_url(end, day, web_link, string=location)
                data.merge_range(begin, day, end, day, big_string, merge_format)
                

def courseConflict(course_arr):
    info = []
    anyCourseConflict = False
    for b in range(len(course_arr)):
        
        h=0
        
        wk = xlrd.open_workbook("Calendar.xlsx")
        sht = wk.sheet_by_index(0)
        
        course1= createCourseObj(course_arr[b])
        
        day_array_one = dayToNumber(course1)
        time_array_one = meetToNumber(course1)
        
        conflictDay = False
        conflictTime = False
        
        if not(h == b-1):
            h+=1
        
            course2= createCourseObj(course_arr[h])
            dayConf = ""
            
            day_array_two = dayToNumber(course2)
            time_array_two = meetToNumber(course2)
            
            
            for i in range(len(day_array_one)):
                if day_array_one[i] in day_array_two:

                    if day_array_one[i] == 1:
                        dayConf = "Mon"
                    elif day_array_one[i] == 2:
                        dayConf = "Tues"
                    elif day_array_one[i] ==3:
                        dayConf = "Wed"
                    elif day_array_one[i] == 4:
                        dayConf = "Thurs"
                    else:
                        dayConf = "Fri"
                        
                    info.append(course_arr[b])
                    info.append(course_arr[h])
                    info.append(dayConf)
                    
                    conflictDay = True
            
            
            time = 0
            for j in range(len(time_array_one)):
                if time_array_one[j] in time_array_two:
                    time = time_array_one[j]
                    info.append(str(sht.cell_value(time, 0)))
                    conflictTime = True
                    
            if conflictDay and conflictTime:
                print("These courses conflict with each other: " + "\n")
                anyCourseConflict = True
                
                for x in info:
                    if "GS" in x:
                        print(x)
                print("\n" + "Your schedule will contain the conflicting course that was most recently entered. \n")
              
    return anyCourseConflict




    
    
    


    
   

     
    
    
    
    
    
    






        
        
        
        
        



    