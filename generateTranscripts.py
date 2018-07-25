from xlrd import open_workbook
import xlrd
from xlutils.copy import copy
import os

def _getOutCell(outSheet, rowIndex, colIndex):
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, row, col, value):
    previousCell = _getOutCell(outSheet, col, row)
    outSheet.write(row, col, value)
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx

def GPA(grade):
	if grade is 'O':
		return 10
	elif grade is 'A':
		return 9
	elif grade is 'B':
		return 8
	elif grade is 'C':
		return 7
	elif grade is 'D':
		return 6
	elif grade is 'E':
		return 5
	else:
		return 0

def fillTable(rangeBottom, rangeTop, s, rb):
	total = 0
	sheet = rb.sheet_by_index(0)
	print "Enter the grade column"
	for i in range (rangeBottom,rangeTop):
		grade = raw_input()
		gpa = GPA(grade)
		setOutCell(s, i, 5, grade)
		setOutCell(s, i, 6, gpa)
		credit = sheet.cell_value(i, 4)
		setOutCell(s, i, 7, gpa*credit)
		total = total + (gpa*credit)
	setOutCell(s, rangeTop, 7, total)
	setOutCell(s, rangeTop+1, 7, raw_input("Pointer (SGPI): "))

def fillTableExtraLine(rangeBottom, rangeTop, s, rb):
	total = 0
	sheet = rb.sheet_by_index(0)
	print "Enter the grade column"
	for i in range (rangeBottom,rangeTop):
		grade = raw_input()
		gpa = GPA(grade)
		setOutCell(s, i, 5, grade)
		setOutCell(s, i, 6, gpa)
		credit = sheet.cell_value(i, 4)
		setOutCell(s, i, 7, gpa*credit)
		total = total + (gpa*credit)
	setOutCell(s, rangeTop+1, 7, total)
	setOutCell(s, rangeTop+2, 7, raw_input("Pointer (SGPI): "))

def semester1(s, rb, year):
	setOutCell(s, 7, 4, name)
	seatNumber = raw_input("Sem1 seat number: ")
	setOutCell(s, 9,3,seatNumber)
	fillTable(12, 24, s, rb)
	passed = "DEC " + year 
	setOutCell(s, 9, 9, passed)

def semester2(s, rb, year):
	seatNumber = raw_input("Sem2 seat number: ")
	s.write(28,3,seatNumber)
	fillTable(31, 44, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 28, 9, passed)

def semester3(s, rb, year):
	seatNumber = raw_input("Sem3 seat number: ")
	s.write(48,3,seatNumber)
	fillTable(51, 62, s, rb)
	passed = "DEC " + year 
	setOutCell(s, 48, 9, passed)

def semester4(s, rb, year):
	seatNumber = raw_input("Sem4 seat number:")
	s.write(66,3,seatNumber)
	fillTable(69, 80, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 66, 9, passed)

def semester5(s, rb, year):
	setOutCell(s2, 8, 4, name)
	seatNumber = raw_input("Sem5 seat number: ")
	s.write(11,3,seatNumber)
	fillTableExtraLine(14, 24, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 11, 9, passed)

def semester6(s, rb, year):
	seatNumber = raw_input("Sem6 seat number: ")
	s.write(29,3,seatNumber)
	fillTable(32, 42, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 29, 9, passed)

def semester7(s, rb, year):
	seatNumber = raw_input("Sem7 seat number: ")
	s.write(46,3,seatNumber)
	fillTableExtraLine(49, 59, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 46, 9, passed)

def semester8(s, rb, year):
	seatNumber = raw_input("Sem8 seat number: ")
	s.write(64,3,seatNumber)
	fillTable(67, 77, s, rb)
	passed = "MAY " + year 
	setOutCell(s, 64, 9, passed)

if not os.path.exists('GeneratedTranscripts'):
    os.makedirs('GeneratedTranscripts')

if os.path.isfile('GeneratedTranscripts/TranscriptPage1.xls'):
    os.remove('GeneratedTranscripts/TranscriptPage1.xls')

if os.path.isfile('GeneratedTranscripts/TranscriptPage2.xls'):
    os.remove('GeneratedTranscripts/TranscriptPage2.xls')

semNumber = input("Till which semester would you like your Transcripts to be?")

name = raw_input("Please enter your name: ")

if semNumber==1:
	rb = open_workbook("TranscriptTemplates/TranscriptSem1.xls", formatting_info=True)
	wb = copy(rb)
	s = wb.get_sheet(0)
	sem1Pass = raw_input("Please enter year of passing sem1: ")
	semester1(s, rb, sem1Pass)
	wb.save('GeneratedTranscripts/TranscriptPage1.xls')

if semNumber==2:
	rb = open_workbook("TranscriptTemplates/TranscriptSem2.xls", formatting_info=True)
	wb = copy(rb)
	s = wb.get_sheet(0)
	sem1Pass = raw_input("Please enter year of passing sem1: ")
	semester1(s, rb, sem1Pass)
	sem2Pass = raw_input("Please enter year of passing sem2: ")
	semester2(s, rb, sem2Pass)
	wb.save('GeneratedTranscripts/TranscriptPage1.xls')

if semNumber==3:
	rb = open_workbook("TranscriptTemplates/TranscriptSem3.xls", formatting_info=True)
	wb = copy(rb)
	s = wb.get_sheet(0)
	sem1Pass = raw_input("Please enter year of passing sem1: ")
	semester1(s, rb, sem1Pass)
	sem2Pass = raw_input("Please enter year of passing sem2: ")
	semester2(s, rb, sem2Pass)
	semester3(s, rb, sem2Pass)
	wb.save('GeneratedTranscripts/TranscriptPage1.xls')

if semNumber==4:
	rb = open_workbook("TranscriptTemplates/TranscriptSem4.xls", formatting_info=True)
	wb = copy(rb)
	s = wb.get_sheet(0)
	sem1Pass = raw_input("Please enter year of passing sem1: ")
	semester1(s, rb, sem1Pass)
	sem2Pass = raw_input("Please enter year of passing sem2: ")
	semester2(s, rb, sem2Pass)
	semester3(s, rb, sem2Pass)
	sem4Pass = raw_input("Please enter year of passing sem4: ")
	semester4(s, rb, sem4Pass)
	wb.save('GeneratedTranscripts/TranscriptPage1.xls')

if semNumber>4:
	rb = open_workbook("TranscriptTemplates/TranscriptSem4+.xls", formatting_info=True)
	wb = copy(rb)
	s = wb.get_sheet(0)
	sem1Pass = raw_input("Please enter year of passing sem1: ")
	semester1(s, rb, sem1Pass)
	sem2Pass = raw_input("Please enter year of passing sem2: ")
	semester2(s, rb, sem2Pass)
	semester3(s, rb, sem2Pass)
	sem4Pass = raw_input("Please enter year of passing sem4: ")
	semester4(s, rb, sem4Pass)
	wb.save('GeneratedTranscripts/TranscriptPage1.xls')

if semNumber == 5:
	rb2 = open_workbook("TranscriptTemplates/Transcript2Sem5.xls", formatting_info=True)
	wb2 = copy(rb2)
	s2 = wb2.get_sheet(0)
	sem5Pass = raw_input("Please enter year of passing sem5: ")
	semester5(s2, rb2, sem5Pass)
	wb2.save('GeneratedTranscripts/TranscriptPage2.xls')

if semNumber==6:
	rb2 = open_workbook("TranscriptTemplates/Transcript2Sem6.xls", formatting_info=True)
	wb2 = copy(rb2)
	s2 = wb2.get_sheet(0)
	sem5Pass = raw_input("Please enter year of passing sem5: ")
	semester5(s2, rb2, sem5Pass)
	sem6Pass = raw_input("Please enter year of passing sem6: ")
	semester6(s2, rb2, sem6Pass)
	wb2.save('GeneratedTranscripts/TranscriptPage2.xls')

if semNumber == 7:
	rb2 = open_workbook("TranscriptTemplates/Transcript2Sem7.xls", formatting_info=True)
	wb2 = copy(rb2)
	s2 = wb2.get_sheet(0)
	sem5Pass = raw_input("Please enter year of passing sem5: ")
	semester5(s2, rb2, sem5Pass)
	sem6Pass = raw_input("Please enter year of passing sem6: ")
	semester6(s2, rb2, sem6Pass)
	semester7(s2, rb2, sem6Pass)
	wb2.save('GeneratedTranscripts/TranscriptPage2.xls')

if semNumber == 8:
	rb2 = open_workbook("TranscriptTemplates/Transcript2Sem8.xls", formatting_info=True)
	wb2 = copy(rb2)
	s2 = wb2.get_sheet(0)
	sem5Pass = raw_input("Please enter year of passing sem5: ")
	semester5(s2, rb2, sem5Pass)
	sem6Pass = raw_input("Please enter year of passing sem6: ")
	semester6(s2, rb2, sem6Pass)
	semester7(s2, rb2, sem6Pass)
	sem8Pass = raw_input("Please enter year of passing sem8: ")
	semester8(s2, rb2, sem8Pass)
	wb2.save('GeneratedTranscripts/TranscriptPage2.xls')