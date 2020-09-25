#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 18 10:15:43 2020

@author: ryandebose-boyd1
"""

import cp1_functions
import numpy as np
import xlrd
import PyPDF2
import os

import fitz
#install



inp = True

some_courses = ['GS-QC-6401', 'GS-GS-6400', 'GS-GS-6600', "GS-QC-5105", "GS-QC-5110"]



man = False
syl = False
val = False

while val == False:
    inp = input("Welcome to BCM ScheduleMaker! \nWould you like to enter your courses manually or with your syllabi (Choose M or S)? ")
    if inp.upper() == "M":
        val = True
        man = True
    elif inp.upper() == "S":
        val = True
        syl = True
    else:
        print("Invalid input. Please try again.")

if man == True:
    cp1_functions.blank()
    course_ar = cp1_functions.manualCourseArray()
    for curr_cors in course_ar:
        course = cp1_functions.createCourseObj(curr_cors)
        
        cp1_functions.courseToTimeSlots(course)
    conf = False
    if len(course_ar) > 1:
        conf = cp1_functions.courseConflict(course_ar)
        
    if conf == False:
        print("Schedule Complete! ")
elif syl == True:
    cp1_functions.blank()
    syl_arr = cp1_functions.syllabiArray()
    for curr_cors in syl_arr:
        cp1_functions.courseFromPDF(curr_cors)
        

cp1_functions.wb.close()
    
    
    






           