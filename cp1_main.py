#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 10:15:43 2020

@author: ryandebose-boyd1
"""

import cp1_functions

cp1_functions.blank()

inp = True

some_courses = ['GS-GS-6600', 'GS-GS-6400', 'GS-NE-6112']

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
    print(j)
    g = None
    g = cp1_functions.createCourseObj(some_courses[j])
    if not (g == None):
        cp1_functions.dayToNumber(g)
        cp1_functions.meetToNumber(g)
        cp1_functions.courseToTimeSlots(g)
        
    else:
        print("Course Name Not Found")

cp1_functions.wb.close()
    
    
    
  






           