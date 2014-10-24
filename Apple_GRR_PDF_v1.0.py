######################################################################################
# Apple_GRR_PDF_v1.0.py
# Created By Han-Jang Chen
# Oct. 22, 2014

#Notes: Please note that the coordinate system for Table in reportlab is different
######################################################################################
import xlsxwriter
import xlrd
import glob
import time
import os
import datetime

#import colorama
from colorama import init
init()
init(autoreset=True) 
# If you find yourself repeatedly sending reset sequences to turn off color changes at the end of every point,
# then init(autoreset=True) will automate that.
from colorama import Fore, Back, Style

from fastnumbers import fast_real
import math
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
import numpy as np

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, letter, inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.platypus import PageBreak

FORES = [ Fore.BLACK, Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.WHITE ]
BACKS = [ Back.BLACK, Back.RED, Back.GREEN, Back.YELLOW, Back.BLUE, Back.MAGENTA, Back.CYAN, Back.WHITE ]
col_index = [ 0,         1,          2,         3,          4,          5,            6,          7]
STYLES = [ Style.DIM, Style.NORMAL, Style.BRIGHT ]
print "\n"
print "="*70

current_time = datetime.datetime.now()
timestamp = current_time.strftime("%Y-%m-%d_%H-%M-%S")
dir_path = "/Users/hanjangchen/Desktop/Apple/0=Projects/Gage_R&R/to_be_PDFed/"
dir_file = glob.glob(dir_path + "processed*") # now it is a list with only one item!
dir_file = dir_file[0] # get the string from the list

print "Data Source:"
print (FORES[6] +"%s\n")%dir_file

newpath = "/Users/hanjangchen/Desktop/Apple/0=Projects/Gage_R&R/to_be_PDFed/%s/"%timestamp # create a new folder
if not os.path.exists(newpath): os.makedirs(newpath)

workbook1 = xlrd.open_workbook(dir_file)
sheet_LAT = workbook1.sheet_by_name("Tx_LAT")

################################### Page 1 ###################################
								# Appraiser 1
appraiser1 = []
temp1 = []
xbar1 = []
range1 = []
parts = 0
for i in range(7, 17):
	if sheet_LAT.cell_value(i, 1) != "":
		parts += 1
		# get measurements
		temp1 = []
		for j in range(1, 4):
			temp1.append(str(round(sheet_LAT.cell_value(i, j), 4)))
		appraiser1.append(temp1)
		
		# get x bar and range
		xbar1.append(str(round(sheet_LAT.cell_value(i, 4), 4)))
		range1.append(str(round(sheet_LAT.cell_value(i, 5), 4)))
	else:
		temp1 = []
		for j in range(1, 4):
			temp1.append("")
		appraiser1.append(temp1)
		xbar1.append("")
		range1.append("")

appraiser1_trial_avg = []
for i in range(0, parts):
	appraiser1_trial_avg.append(str(round(sheet_LAT.cell_value(17, 1+i), 4)))
	
X1_E18 = str(round(sheet_LAT.cell_value(17, 4), 4))	
R1_F19 = str(round(sheet_LAT.cell_value(18, 5), 4))
print (FORES[2] + STYLES[2] + "\nappraiser1 imported (%s parts)"%parts)
#print "\nappraiser 1\n", appraiser1
#print "\nxbar1\n", xbar1
#print "\nrange1\n", range1
#print "X1", X1_E18
#print "R1", R1_F19
print appraiser1
print xbar1
								# Appraiser 2
appraiser2 = []
temp1 = []
xbar2 = []
range2 = []
R2 = []
for i in range(7, 17):
	if sheet_LAT.cell_value(i, 6) != "": 
		temp1 = []
		for j in range(6, 9):
			temp1.append(str(round(sheet_LAT.cell_value(i, j), 4)))
		appraiser2.append(temp1)
		# get x bar and range
		xbar2.append(str(round(sheet_LAT.cell_value(i, 9), 4)))
		range2.append(str(round(sheet_LAT.cell_value(i, 10), 4)))
	else:
		temp1 = []
		for j in range(6, 9):
			temp1.append("")
		appraiser2.append(temp1)
		xbar2.append("")
		range2.append("")
		
appraiser2_trial_avg = []
for i in range(0, parts):
	appraiser2_trial_avg.append(str(round(sheet_LAT.cell_value(17, 6+i), 4)))
		
X2_J18 = str(round(sheet_LAT.cell_value(17, 9), 4))	
R2_K19= str(round(sheet_LAT.cell_value(18, 10), 4))		
		
print (FORES[2] + STYLES[2] + "\nappraiser2 imported (%s parts)")%parts
#print "\nappraiser 2\n", appraiser2
#print "\nxbar2\n", xbar2
#print "\nrange2\n", range2
#print "X2", X2_J18
#print "R2", R2_K19
print appraiser2
print xbar2
								# Appraiser 3
appraiser3 = []
temp1 = []
xbar3 = []
range3 = []
for i in range(7, 17):
	if sheet_LAT.cell_value(i, 6) != "": 
		temp1 = []
		for j in range(11, 14):
			temp1.append(str(round(sheet_LAT.cell_value(i, j), 4)))
		appraiser3.append(temp1)
		# get x bar and range
		xbar3.append(str(round(sheet_LAT.cell_value(i, 14), 4)))
		range3.append(str(round(sheet_LAT.cell_value(i, 15), 4)))
	else:
		temp1 = []
		for j in range(6, 9):
			temp1.append("")
		appraiser3.append(temp1)
		xbar3.append("")
		range3.append("")
		
appraiser3_trial_avg = []
for i in range(0, parts):
	appraiser3_trial_avg.append(str(round(sheet_LAT.cell_value(17, 11+i), 4)))
		
X3_O18 = str(round(sheet_LAT.cell_value(17, 14), 4))	
R3_P19= str(round(sheet_LAT.cell_value(18, 15), 4))				
		
print (FORES[2] + STYLES[2] + "\nappraiser3 imported (%s parts)")%parts
#print "\nappraiser 3\n", appraiser3
#print "\nxbar3\n", xbar3
#print "\nrange3\n", range3
#print "X3", X3_O18
#print "R3", R3_P19
print appraiser3
print xbar3




								# Evaluation
part_average_Q7 = []
for i in range(7, 17):
	if sheet_LAT.cell_value(i, 16) != "":
		part_average_Q7.append(str(round(sheet_LAT.cell_value(i, 16), 4)))
	else:
		part_average_Q7.append("")
print "\npart_average\n", part_average_Q7

Xdiff_B20 = sheet_LAT.cell_value(19, 1)
R_barbar_B21 = sheet_LAT.cell_value(20, 1)
Rpart_B22 = sheet_LAT.cell_value(21, 1)
parts_E20 = sheet_LAT.cell_value(19, 4)
trials_E21 = sheet_LAT.cell_value(20, 4)
operators_E22 = sheet_LAT.cell_value(21, 16)

EV_B26 = sheet_LAT.cell_value(25, 1)
AV_B27 = sheet_LAT.cell_value(26, 1)
R_and_R_B28 = sheet_LAT.cell_value(27, 1)
PV_B29 = sheet_LAT.cell_value(28, 1)
TV_B30 = sheet_LAT.cell_value(29, 1)
EV_percent_TV_C26 = sheet_LAT.cell_value(25, 2)
AV_percent_TV_C27 = sheet_LAT.cell_value(26, 2)
R_and_R_percent_TV_C28 = sheet_LAT.cell_value(27, 2)
PV_percent_TV_C29 = sheet_LAT.cell_value(28, 2)

plot_unstacked_xaxis = []
for i in range(5, 39):
	plot_unstacked_xaxis.append(str(sheet_LAT.cell_value(i, 19)))

plot_stacked_xaxis = []
for i in range(7, 17):
	plot_stacked_xaxis.append(str(sheet_LAT.cell_value(i, 0)))

#print "\nstacked_xaxis\n", plot_stacked_xaxis
#print "\nunstacked_xaxis\n", plot_unstacked_xaxis

doc = SimpleDocTemplate(newpath + "Table.pdf", pagesize = letter)
elements = []

data = []
temp1 = []
data.append(["Gage Repeatability and Reproducibility Data Collection Sheet"])
data.append([])
data.append(["Apparaiser\n/ Trial#", "", "PART", "", "", "", "", "", "", "", "", "", "Average", "", ])
data.append(["", "", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "", ""])

data.append(["Appraiser 1 Trial 1", "", appraiser1[0][0], appraiser1[1][0], appraiser1[2][0], appraiser1[3][0], appraiser1[4][0], appraiser1[5][0], appraiser1[6][0], appraiser1[7][0], appraiser1[8][0], appraiser1[9][0], appraiser1_trial_avg[0], ""])
data.append(["Appraiser 1 Trial 2", "", appraiser1[0][1], appraiser1[1][1], appraiser1[2][1], appraiser1[3][1], appraiser1[4][1], appraiser1[5][1], appraiser1[6][1], appraiser1[7][1], appraiser1[8][1], appraiser1[9][1], appraiser1_trial_avg[1], ""])
data.append(["Appraiser 1 Trial 3", "", appraiser1[0][2], appraiser1[1][2], appraiser1[2][2], appraiser1[3][2], appraiser1[4][2], appraiser1[5][2], appraiser1[6][2], appraiser1[7][2], appraiser1[8][2], appraiser1[9][2], appraiser1_trial_avg[2], ""])
data.append(["Average", "", xbar1[0], xbar1[1], xbar1[2], xbar1[3], xbar1[4], xbar1[5], xbar1[6], xbar1[7], xbar1[8], xbar1[9], "Xa = %s"%str(round(sheet_LAT.cell_value(17, 4), 4)), ""])
data.append(["Range", "", range1[0], range1[1], range1[2], range1[3], range1[4], range1[5], range1[6], range1[7], range1[8], range1[9], "Ra = %s"%str(round(sheet_LAT.cell_value(18, 5), 4)), ""])

data.append(["Appraiser 2 Trial 1", "", appraiser2[0][0], appraiser2[1][0], appraiser2[2][0], appraiser2[3][0], appraiser2[4][0], appraiser2[5][0], appraiser2[6][0], appraiser2[7][0], appraiser2[8][0], appraiser2[9][0], appraiser2_trial_avg[0], ""])
data.append(["Appraiser 2 Trial 2", "", appraiser2[0][1], appraiser2[1][1], appraiser2[2][1], appraiser2[3][1], appraiser2[4][1], appraiser2[5][1], appraiser2[6][1], appraiser2[7][1], appraiser2[8][1], appraiser2[9][1], appraiser2_trial_avg[1], ""])
data.append(["Appraiser 2 Trial 3", "", appraiser2[0][2], appraiser2[1][2], appraiser2[2][2], appraiser2[3][2], appraiser2[4][2], appraiser2[5][2], appraiser2[6][2], appraiser2[7][2], appraiser2[8][2], appraiser2[9][2], appraiser2_trial_avg[2], ""])
data.append(["Average", "", xbar2[0], xbar2[1], xbar2[2], xbar2[3], xbar2[4], xbar2[5], xbar2[6], xbar2[7], xbar2[8], xbar2[9], "Xb = %s"%str(round(sheet_LAT.cell_value(17, 9), 4)), ""])
data.append(["Range", "", range2[0], range2[1], range2[2], range2[3], range2[4], range2[5], range2[6], range2[7], range2[8], range2[9], "Rb = %s"%str(round(sheet_LAT.cell_value(18, 10), 4)), ""])

data.append(["Appraiser 3 Trial 1", "", appraiser3[0][0], appraiser3[1][0], appraiser3[2][0], appraiser3[3][0], appraiser3[4][0], appraiser3[5][0], appraiser3[6][0], appraiser3[7][0], appraiser3[8][0], appraiser3[9][0], appraiser3_trial_avg[0], ""])
data.append(["Appraiser 3 Trial 2", "", appraiser3[0][1], appraiser3[1][1], appraiser3[2][1], appraiser3[3][1], appraiser3[4][1], appraiser3[5][1], appraiser3[6][1], appraiser3[7][1], appraiser3[8][1], appraiser3[9][1], appraiser3_trial_avg[1], ""])
data.append(["Appraiser 3 Trial 3", "", appraiser3[0][2], appraiser3[1][2], appraiser3[2][2], appraiser3[3][2], appraiser3[4][2], appraiser3[5][2], appraiser3[6][2], appraiser3[7][2], appraiser3[8][2], appraiser3[9][2], appraiser3_trial_avg[2], ""])
data.append(["Average", "", xbar3[0], xbar3[1], xbar3[2], xbar3[3], xbar3[4], xbar3[5], xbar3[6], xbar3[7], xbar3[8], xbar3[9], "Xc = %s"%str(round(sheet_LAT.cell_value(17, 14), 4)), ""])
data.append(["Range", "", range3[0], range3[1], range3[2], range3[3], range3[4], range3[5], range3[6], range3[7], range3[8], range3[9], "Rc = %s"%str(round(sheet_LAT.cell_value(18, 15), 4)), ""])

data.append(["Part\nAverage", "", part_average_Q7[0], part_average_Q7[1], part_average_Q7[2], part_average_Q7[3], part_average_Q7[4], part_average_Q7[5], part_average_Q7[6], part_average_Q7[7], part_average_Q7[8], part_average_Q7[9], "X = %s\nRp = %s"%(round((sheet_LAT.cell_value(17, 4)+sheet_LAT.cell_value(17, 9)+sheet_LAT.cell_value(17, 14))/3, 4), round(sheet_LAT.cell_value(21, 1), 4)), ""])
data.append([])
data.append(["([Ra = %s] + [Rb = %s] + [Rc = %s]) / [# OF APPRAISERS = %s] = %s "%(str(round(sheet_LAT.cell_value(18, 5), 4)), str(round(sheet_LAT.cell_value(18, 10), 4)), str(round(sheet_LAT.cell_value(18, 15), 4)), int(sheet_LAT.cell_value(21, 4)), str(round(sheet_LAT.cell_value(20, 1), 4))), "", "", "", "", "", "", "", "", "", "", "", "R = %s"%str(round(sheet_LAT.cell_value(20, 1), 4)), ""]) 
data.append(["Xdiff = Max(Xa, Xb, Xc) - Min(Xa, Xb, Xc) = %s"%str(round(sheet_LAT.cell_value(19, 1), 4)), "", "", "", "", "", "", "", "", "", "", "", "", ""])
data.append(["UCL = [D4 = %s]*R = %s"%(str(round(sheet_LAT.cell_value(24, 7), 4)), str(round(sheet_LAT.cell_value(22, 1), 4))), "", "", "", "", "", "", "", "", "", "", "", "", ""]) 
data.append(["D4 = 3.27 for 2 trials and 2.575 for 3 trials.\n\n\nNotes:"])
data.append([])
data.append([])
data.append([])
data.append([])

t = Table(data, 13*[0.5*inch], 29*[0.25*inch])
t.setStyle(TableStyle([
("SPAN", (0, 0), (13, 1)),("ALIGN", (0, 0), (13, 1), "CENTER"), ("VALIGN", (0, 0), (13, 1), "MIDDLE") 
, ("SPAN", (0, 2), (1, 3)), ("ALIGN", (0, 2), (1, 3), "CENTER"), ("VALIGN", (0, 2), (1, 3), "MIDDLE"), ("ALIGN", (2, 3), (11, 3), "CENTER"), ("SPAN", (2, 2), (11, 2)), ("ALIGN", (2, 2), (11, 2), "CENTER"), ("VALIGN", (2, 2), (11, 2), "MIDDLE"), ("SPAN", (12, 2), (13, 3)), ("ALIGN", (12, 2), (13, 3), "CENTER"), ("VALIGN", (12, 2), (13, 3), "MIDDLE")
# Appraiser 1
, ("SPAN", (0, 4), (1, 4)), ("FONTSIZE", (0, 4), (13, 4), 7), ("ALIGN", (0, 4), (13, 4), "CENTER"), ("VALIGN", (2, 4), (13, 4), "MIDDLE"), ("SPAN", (12, 4), (13, 4))
, ("SPAN", (0, 5), (1, 5)), ("FONTSIZE", (0, 5), (13, 5), 7), ("ALIGN", (0, 5), (13, 5), "CENTER"), ("VALIGN", (2, 5), (13, 5), "MIDDLE"), ("SPAN", (12, 5), (13, 5))
, ("SPAN", (0, 6), (1, 6)), ("FONTSIZE", (0, 6), (13, 6), 7), ("ALIGN", (0, 6), (13, 6), "CENTER"), ("VALIGN", (2, 6), (13, 6), "MIDDLE"), ("SPAN", (12, 6), (13, 6))
, ("SPAN", (0, 7), (1, 7)), ("FONTSIZE", (0, 7), (13, 7), 7), ("ALIGN", (0, 7), (13, 7), "CENTER"), ("VALIGN", (2, 7), (13, 7), "MIDDLE"), ("SPAN", (12, 7), (13, 7)), ("FONTSIZE", (12, 7), (12, 7), 7)
, ("SPAN", (0, 8), (1, 8)), ("FONTSIZE", (0, 8), (13, 8), 7), ("ALIGN", (0, 8), (13, 8), "CENTER"), ("VALIGN", (2, 8), (13, 8), "MIDDLE"), ("SPAN", (12, 8), (13, 8)), ("FONTSIZE", (12, 8), (12, 8), 7)
, ("BOX", (0, 4), (13, 8), 1, colors.black)
# Appraiser 2
, ("SPAN", (0, 9), (1, 9)), ("FONTSIZE", (0, 9), (13, 9), 7), ("ALIGN", (0, 9), (13, 9), "CENTER"), ("VALIGN", (2, 9), (13, 9), "MIDDLE"), ("SPAN", (12, 9), (13, 9))
, ("SPAN", (0, 10), (1, 10)), ("FONTSIZE", (0, 10), (13, 10), 7), ("ALIGN", (0, 10), (13, 10), "CENTER"), ("VALIGN", (2, 10), (13, 10), "MIDDLE"), ("SPAN", (12, 10), (13, 10))
, ("SPAN", (0, 11), (1, 11)), ("FONTSIZE", (0, 11), (13, 11), 7), ("ALIGN", (0, 11), (13, 11), "CENTER"), ("VALIGN", (2, 11), (13, 11), "MIDDLE"), ("SPAN", (12, 11), (13, 11))
, ("SPAN", (0, 12), (1, 12)), ("FONTSIZE", (0, 12), (13, 12), 7), ("ALIGN", (0, 12), (13, 12), "CENTER"), ("VALIGN", (2, 12), (13, 12), "MIDDLE"), ("SPAN", (12, 12), (13, 12)), ("FONTSIZE", (12, 12), (12, 12), 7)
, ("SPAN", (0, 13), (1, 13)), ("FONTSIZE", (0, 13), (13, 13), 7), ("ALIGN", (0, 13), (13, 13), "CENTER"), ("VALIGN", (2, 13), (13, 13), "MIDDLE"), ("SPAN", (12, 13), (13, 13)), ("FONTSIZE", (12, 13), (12, 13), 7)
, ("BOX", (0, 9), (13, 13), 1, colors.black)
# Appraiser 3
, ("SPAN", (0, 14), (1, 14)), ("FONTSIZE", (0, 14), (13, 14), 7), ("ALIGN", (0, 14), (13, 14), "CENTER"), ("VALIGN", (2, 14), (13, 14), "MIDDLE"), ("SPAN", (12, 14), (13, 14))
, ("SPAN", (0, 15), (1, 15)), ("FONTSIZE", (0, 15), (13, 15), 7), ("ALIGN", (0, 15), (13, 15), "CENTER"), ("VALIGN", (2, 15), (13, 15), "MIDDLE"), ("SPAN", (12, 15), (13, 15))
, ("SPAN", (0, 16), (1, 16)), ("FONTSIZE", (0, 16), (13, 16), 7), ("ALIGN", (0, 16), (13, 16), "CENTER"), ("VALIGN", (2, 16), (13, 16), "MIDDLE"), ("SPAN", (12, 16), (13, 16))
, ("SPAN", (0, 17), (1, 17)), ("FONTSIZE", (0, 17), (13, 17), 7), ("ALIGN", (0, 17), (13, 17), "CENTER"), ("VALIGN", (2, 17), (13, 17), "MIDDLE"), ("SPAN", (12, 17), (13, 17)), ("FONTSIZE", (12, 17), (12, 17), 7)
, ("SPAN", (0, 18), (1, 18)), ("FONTSIZE", (0, 18), (13, 18), 7), ("ALIGN", (0, 18), (13, 18), "CENTER"), ("VALIGN", (2, 18), (13, 18), "MIDDLE"), ("SPAN", (12, 18), (13, 18)), ("FONTSIZE", (12, 18), (12, 18), 7)
, ("BOX", (0, 14), (13, 18), 1, colors.black)
#PART Average
, ("SPAN", (0, 18), (1, 18)), ("FONTSIZE", (0, 18), (0, 18), 7), ("ALIGN", (0, 18), (1, 18), "CENTER"), ("SPAN", (12, 18), (13, 18)), ("FONTSIZE", (12, 19), (12, 19), 7)
, ("SPAN", (0, 19), (1, 20)), ("FONTSIZE", (0, 19), (13, 19), 7), ("ALIGN", ( 0, 19), (1, 20), "CENTER"), ("VALIGN", (0, 19), (1, 20), "MIDDLE")
, ("SPAN", (2, 19), (2, 20)), ("SPAN", (3, 19), (3, 20)), ("SPAN", (4, 19), (4, 20)), ("SPAN", (5, 19), (5, 20)), ("SPAN", (6, 19), (6, 20)), ("SPAN", (7, 19), (7, 20)), ("SPAN", (8, 19), (8, 20)), ("SPAN", (9, 19), (9, 20)), ("SPAN", (10, 19), (10, 20)), ("SPAN", (11, 19), (11, 20)), ("VALIGN", (2, 19), (11, 19), "MIDDLE")  
, ("SPAN", (12, 19), (13, 20)), ("ALIGN", (12, 19), (12, 19), "LEFT"), ("VALIGN", (12, 19), (12, 19), "MIDDLE")

, ("SPAN", (0, 21), (11, 21)), ("SPAN", (12, 21), (13, 21)), ("FONTSIZE", (0, 21), (0, 21), 7), ("FONTSIZE", (12, 21), (12, 21), 7)
, ("SPAN", (0, 22), (11, 22)), ("SPAN", (12, 22), (13, 22)), ("FONTSIZE", (0, 22), (0, 22), 7), ("FONTSIZE", (12, 22), (12, 22), 7)
# UCL
, ("SPAN", (0, 23), (11, 23)), ("SPAN", (12, 23), (13, 23)), ("FONTSIZE", (0, 23), (0, 23), 7)
, ("SPAN", (0, 24), (-1, -1)), ("ALIGN", (0, 24), (0, 24), "LEFT"), ("VALIGN", (0, 24), (0, 24), "TOP"), ("FONTSIZE", (0, 24), (0, 24), 6)

, ("BOX", (0, 0), (-1, -1), 1, colors.black)
, ("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))


elements.append(t)
elements.append(PageBreak()) # going to next page
################################### Page 2 ###################################
data2 = []
temp1 = []
data2.append(["Gage Repeatability and Reproducibility Data Collection Sheet"])
data2.append(["Part No.&Name: \nCharacteristics: \nSpecifications: \nGage Name: \nDate: \nPerformed by: \nFrom datasheet: R = %s, Xdiff = %s, Rp = %s"%(str(round(sheet_LAT.cell_value(20, 1), 4)), str(round(sheet_LAT.cell_value(19, 1), 4)), str(round(sheet_LAT.cell_value(21, 1), 4)))])
data2.append("")
data2.append("")
data2.append("")
data2.append("")
data2.append(["Measurement Unit Analysis"])

for i in range(7, 36):
	temp1 = []
	for j in range(0, 3):
		temp1.append("")
	data2.append(temp1)
t2 = Table(data2, 3*[2*inch],36*[0.25*inch])
t2.setStyle(TableStyle([
("SPAN", (0, 0), (2, 0)), ("ALIGN", (0, 0), (2, 0), "CENTER")
, ("SPAN", (0, 1), (2, 5)), ("ALING", (0, 1), (2, 5), "LEFT"), ("VALIGN", (0, 1), (2, 5), "TOP"), ("FONTSIZE", (0, 1), (2, 5), 7)
, ("SPAN", (0, 6), (1, 6)), ("ALIGN", (0, 6), (1, 6), "CENTER"), ("FONTSIZE", (0, 6), (2, 6), 3)
, ("SPAN", (0, 7), (1, 11))
, ("GRID", (0, 0), (-1, -1), 2, colors.black)]))
elements.append(t2)


doc.build(elements)

print (FORES[5] + STYLES[2] + "\nCompleted")



