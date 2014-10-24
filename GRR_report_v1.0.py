######################################################################################
# GRR_report_v1.0.py
# Created By Han-Jang Chen
# Oct. 21, 2014

#Notes:
######################################################################################
import xlsxwriter
import xlrd
import glob
import time
import os
import datetime

import colorama
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

FORES = [ Fore.BLACK, Fore.RED, Fore.GREEN, Fore.YELLOW, Fore.BLUE, Fore.MAGENTA, Fore.CYAN, Fore.WHITE ]
BACKS = [ Back.BLACK, Back.RED, Back.GREEN, Back.YELLOW, Back.BLUE, Back.MAGENTA, Back.CYAN, Back.WHITE ]
col_index = [ 0,         1,          2,         3,          4,          5,            6,          7]
STYLES = [ Style.DIM, Style.NORMAL, Style.BRIGHT ]

##################################################################################
def get_mean(data_list):
	sum = 0.0
	count = len(data_list)
	
	# get mean
	for i in data_list:
		if i != "":
			sum += i
			
	#print sum/count
	print sum
	print count
	a = sum/count
	print a
	return a	 # calculates mean of a list of values
def get_min(data_list):
	new = []
	for i in data_list:
		if i != "":
			new.append(fast_real(i))
	min = sorted(new)
	return round(min[0], 4)	# calculates min of a list of values
def get_max(data_list):
	new = []
	for i in data_list:
		if i != "":
			new.append(fast_real(i))
	max = sorted(new)
	return round(max[len(new)-1], 4) # calculates max of a list of values
def get_std_lat_lon(data_list):
	sum = 0.0
	count = 0.0
	new = []
	# get mean
	for i in data_list:
		if i != "":
			sum += fast_real(i)
			count += 1
			new.append(fast_real(i))
	mean = sum/count
	
	a = []
	b = []
	c = 0.0
	for i in new:
		a.append(i - mean)
	for i in a:
		b.append(i*i)
	for i in b:
		c+=i
	std = math.sqrt(c/count)
	return round(std, 4)	# calculates stdev of a list of values

print "\n"
print "="*70
# sys.argv[1]
current_time = datetime.datetime.now()
timestamp = current_time.strftime("%Y-%m-%d_%H-%M-%S")
dir_path = "/Users/hanjangchen/Desktop/Apple/0=Projects/Gage_R&R/AppleTemplate/"
dir_file = glob.glob(dir_path + "FXGL*" + "/") # now it is a list with only one item!

dir_file = glob.glob(dir_file[0] + "processed*")
dir_file = dir_file[0]
#dir_file = dir_file[0] # get the string from the list

print "Data Source:"
print (FORES[6] +"%s\n")%dir_file

newpath = "/Users/hanjangchen/Desktop/Apple/0=Projects/Gage_R&R/AppleTemplate/TempReport_%s/"%timestamp # create a new folder
if not os.path.exists(newpath): os.makedirs(newpath)

workbook1 = xlrd.open_workbook(dir_file) 
sheet_Tx_LAT = workbook1.sheet_by_name("Tx_LAT")
sheet_Tx_UAT = workbook1.sheet_by_name("Tx_UAT")

# Appraiser 1
x_barA = []
range_A = []
temp1 = []
temp2 = []
for i in range(7, 17):
	temp1 = []
	if sheet_Tx_LAT.cell_value(i, 1) != "":
		for j in range(1, 4):
			temp1.append(fast_real(sheet_Tx_LAT.cell_value(i, j)))	
		print temp1	
	x_barA.append(get_mean(temp1))
	range_A.append(get_max(temp1)-get_min(temp1))
average_x_barA = get_mean(xbar_A)
average_range_A = get_mean(range_A)
	
"""# Appraiser 2		
x_barB = []
range_B = []
temp1 = []
temp2 = []
for row in range(7, 17):
	for col in range(6, 9):
		temp1 = []
		temp1.append(fast_real(sheet_Tx_LAT.cell_value(row, col)))
	
	x_barB.append(get_mean(temp1))
	range_B.append(get_max(temp1)-get_min(temp1))
average_x_barB = get_mean(xbar_B)
average_range_B = get_mean(range_B)

# Appraiser 3		
x_barC = []
range_C = []
temp1 = []
temp2 = []
for row in range(7, 17):
	for col in range(11, 14):
		temp1 = []
		temp1.append(fast_real(sheet_Tx_LAT.cell_value(row, col)))
	
	x_barC.append(get_mean(temp1))
	range_C.append(get_max(temp1)-get_min(temp1))
average_x_barC = get_mean(xbar_C)
average_range_C = get_mean(range_C)

print "\nx_barA\n", x_barA
print "\nrange_A\n", range_A

print "\nx_barB\n", x_barA
print "\nrange_B\n", range_A

print "\nx_barC\n", x_barA
print "\nrange_C\n", range_A"""










