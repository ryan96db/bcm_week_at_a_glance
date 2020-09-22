#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 10:15:43 2020

@author: ryandebose-boyd1
"""

import cp1_functions
import numpy as np
import xlrd

cp1_functions.blank()

inp = True

some_courses = ['GS-CP-6203', 'GS-CC-6202', 'GS-CC-6208']

# course_array = []
# while inp:
    
#     input_class = input("Enter course name: ")
    
#     course_array.append(input_class)
#     cont = input("Enter another course? (Y/N): ")
#     if cont.upper() == 'N':
#         inp = False
#         break
        

# print(course_array)
cp1_functions.blank()

for j in range(len(some_courses)):
    g = None
    g = cp1_functions.createCourseObj(some_courses[j])
    if not (g == None):
        cp1_functions.courseToTimeSlots(g)
        
    else:
        print("Course Name Not Found")



def courseConflict(course_arr):
    info = []
    anyCourseConflict = False
    for b in range(len(course_arr)):
        
        h=0
        
        wk = xlrd.open_workbook("Calendar.xlsx")
        sht = wk.sheet_by_index(0)
        
        twodimarray = []
        course1= cp1_functions.createCourseObj(some_courses[b])
        
        day_array_one = cp1_functions.dayToNumber(course1)
        time_array_one = cp1_functions.meetToNumber(course1)
        
        conflictDay = False
        conflictTime = False
        
        if not(h == b-1):
            h+=1
        
            course2= cp1_functions.createCourseObj(some_courses[h])
            dayConf = ""
            timeConf = ""
            
            day_array_two = cp1_functions.dayToNumber(course2)
            time_array_two = cp1_functions.meetToNumber(course2)
            
            
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
                noConflictingCourse = False
                for x in info:
                    if "GS" in x:
                        print(x)
                print("\n" + "Your schedule will contain the conflicting course that was most recently entered. \n")
              
    return anyCourseConflict
        
    
    
    
q = courseConflict(some_courses)
if q == False:
    print("Schedule Complete! ")
else:
    courseConflict(some_courses)
    

cp1_functions.wb.close()
    
    
    
  






           