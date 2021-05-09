# Reading an excel file using Python 
import xlrd 

#C:\Users\dndou\OneDrive\Documents\ASU\Fall 2020\EGR 402/Agency Test Names.xlsx

# Give the location of the file 
file_location = ("C:\Users\dndou\OneDrive\Documents\ASU\Z (2020-2021) Senior 2\Fall 2020\EGR 402\Agency Test Names.xlsx") 

# To open Workbook 
workbook = xlrd.open_workbook(file_location) 
inputSheet = workbook.sheet_by_index(0)

#############################################################################

import xlwt 
from xlwt import Workbook

# Workbook is created 
wb = Workbook() 

# add_sheet is used to create sheet. 
outputSheet1 = wb.add_sheet('Sheet 1')

outputSheet1.write(0, 0, 'Agency')
outputSheet1.write(0, 1, 'Batch Number')
outputSheet1.write(0, 2, 'Agency Number')
outputSheet1.write(0, 3, 'Old Serial Number')
outputSheet1.write(0, 4, 'New Serial Scan')
outputSheet1.write(0, 5, 'Hand Scan')

#############################################################################
#############################################################################

#(row,column)

i = 0
Max = inputSheet.nrows
print "Max:",Max

for i in range(inputSheet.nrows):    

    if i == Max:
        break
    
    if i != Max:
  
        if i == 0:
            print(i)
            i = 1
            k = i-1
            j = i+1
            p = i
            t = 0


        elif inputSheet.cell_value(i, 0) == inputSheet.cell_value(k, 0):

            print(i)
            print(k)

            if i > 6 and t < Max and inputSheet.cell_value(i, 0) == inputSheet.cell_value(k, 0):
                
                if m >= 6:
                    
                    if p > 6:
                        i = t
                        
                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_6 = inputSheet.cell_value(i, 1)
                    print "i6:",i
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m

                    NewAB3_6 = raw_input("Scan 6th Barcode:")
                    print "NewAB3_6:",NewAB3_6

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 6')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_6)
                    outputSheet1.write(i, 4, NewAB3_6) 

                    m = m-5
                    p = p+1
                    k = k+1
              
                    if p >= 6:
                        i = i+1
                        t = i
                    
                   
                if m == 5:
                    
                    if p >= 6:
                        i = t
                        
                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_5 = inputSheet.cell_value(i, 1)
                    print "i5:",i
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m

                    NewAB3_5 = raw_input("Scan 5th Barcode:")
                    print "NewAB3_5:",NewAB3_5

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 5')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_5)
                    outputSheet1.write(i, 4, NewAB3_5) 

                    m = m+1
                    p = p+1
                    k = k+1
                    
                    if p >= 6:
                        i = i+1
                        t = i
                    
                        
                   
                if m == 4:
                    
                    if p >= 6:
                        i = t
                        
                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_4 = inputSheet.cell_value(i, 1)
                    print "i4:",i
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m
                    
                    NewAB3_4 = raw_input("Scan 4th Barcode:")
                    print "NewAB3_4:",NewAB3_4

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 4')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_4)
                    outputSheet1.write(i, 4, NewAB3_4) 

                    m = m+1
                    p = p+1
                    k = k+1
                    
                    if p >= 6:
                        i = i+1
                        t = i
                               
                   
                if m == 3:

                    if p >= 6:
                        i = t                
                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_3 = inputSheet.cell_value(i, 1)
                    print "i3:",i
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m

                    NewAB3_3 = raw_input("Scan 3rd Barcode:")
                    print "NewAB3_3:",NewAB3_3

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 3')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_3)
                    outputSheet1.write(i, 4, NewAB3_3)

                    m = m+1
                    p = p+1
                    k = k+1
                    if p >= 6:
                        i = i+1
                        t = i
                    
                   
                if m == 2:
                    
                    if p >= 6:
                        i = t
                         
                    print "i2:",i
                    print "k:",k

                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_2 = inputSheet.cell_value(i, 1)
                    print "i2:",i
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m

                    NewAB3_2 = raw_input("Scan 2nd Barcode:")
                    print "NewAB3_2:",NewAB3_2

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 2')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_2)
                    outputSheet1.write(i, 4, NewAB3_2)
                    
                    m = m+1
                    p = p+1
                    k = k+1
                    if p >= 6:
                        i = i+1
                        t = i
               
                   
                if m == 1:
                    if p >= 8:
                        i = t
                    Agency = inputSheet.cell_value(i, 0)
                    OldAB3_1 = inputSheet.cell_value(i, 1)
                    print "i1:",i               
                    print "k:",k
                    print(Agency)
                    print "p:",p
                    print "m:",m

                    NewAB3_1 = raw_input("Scan 1st Barcode:")
                    print "NewAB3_1:",NewAB3_1

                    outputSheet1.write(i, 0, Agency)
                    outputSheet1.write(i, 1, 'NewAB3 1')
                    outputSheet1.write(i, 2, p)
                    outputSheet1.write(i, 3, OldAB3_1)
                    outputSheet1.write(i, 4, NewAB3_1)
                    
                    m = m+1
                    p = p+1
                    k = k+1
                    if p >= 6:
                        i = i+1
                        t = i



               
            if i == 6:
               Agency = inputSheet.cell_value(i, 0)
               OldAB3_6 = inputSheet.cell_value(i, 1)
               print(Agency)
               print "p:",p

               NewAB3_6 = raw_input("Scan 6th Barcode: ")
               print "NewAB3_6:",NewAB3_6

               outputSheet1.write(i, 0, Agency)
               outputSheet1.write(i, 1, 'NewAB3 6')
               outputSheet1.write(i, 2, p)
               outputSheet1.write(i, 3, OldAB3_6)
               outputSheet1.write(i, 4, NewAB3_6)
                
               i = i+1
               k = i-1
               m = 1
               p = p+1
               
            if i == 5:
               Agency = inputSheet.cell_value(i, 0)
               OldAB3_5 = inputSheet.cell_value(i, 1)
               print(Agency)

               NewAB3_5 = raw_input("Scan 5th Barcode: ")
               print "NewAB3_5:",NewAB3_5

               outputSheet1.write(i, 0, Agency)
               outputSheet1.write(i, 1, 'NewAB3 5')
               outputSheet1.write(i, 2, p)
               outputSheet1.write(i, 3, OldAB3_5)
               outputSheet1.write(i, 4, NewAB3_5)
                
               i = i+1
               k = i-1
               m = 0
               p = p+1
               
            if i == 4:
               Agency = inputSheet.cell_value(i, 0)
               OldAB3_4 = inputSheet.cell_value(i, 1)
               print(Agency)

               NewAB3_4 = raw_input("Scan 4th Barcode: ")
               print "NewAB3_4:",NewAB3_4

               outputSheet1.write(i, 0, Agency)
               outputSheet1.write(i, 1, 'NewAB3 4')
               outputSheet1.write(i, 2, p)
               outputSheet1.write(i, 3, OldAB3_4)
               outputSheet1.write(i, 4, NewAB3_4)
                
               i = i+1
               k = i-1
               m = 0
               p = p+1
               
            if i == 3:
               Agency = inputSheet.cell_value(i, 0)
               OldAB3_3 = inputSheet.cell_value(i, 1)
               print(Agency)

               NewAB3_3 = raw_input("Scan 3rd Barcode: ")
               print "NewAB3_3:",NewAB3_3

               outputSheet1.write(i, 0, Agency)
               outputSheet1.write(i, 1, 'NewAB3 3')
               outputSheet1.write(i, 2, p)
               outputSheet1.write(i, 3, OldAB3_3)
               outputSheet1.write(i, 4, NewAB3_3)

               i = i+1
               k = i-1
               m = 0
               p = p+1
               
            if i == 2:
               Agency = inputSheet.cell_value(i, 0)
               OldAB3_2 = inputSheet.cell_value(i, 1)
               print(Agency)

               NewAB3_2 = raw_input("Scan 2nd Barcode: ")
               print "NewAB3_2:",NewAB3_2

               outputSheet1.write(i, 0, Agency)
               outputSheet1.write(i, 1, 'NewAB3 2')
               outputSheet1.write(i, 2, p)
               outputSheet1.write(i, 3, OldAB3_2)
               outputSheet1.write(i, 4, NewAB3_2)
            
               i = i+1
               k = i-1
               m = 0
               p = p+1

            if i == 1:
                p = 1
                Agency = inputSheet.cell_value(i, 0)
                OldAB3_1 = inputSheet.cell_value(i, 1)        
                print(Agency)
                print "p:",p
                
                NewAB3_1 = raw_input("Scan 1st Barcode:")
                print "NewAB3_1:",NewAB3_1

                outputSheet1.write(i, 0, Agency)
                outputSheet1.write(i, 1, 'NewAB3 1')        
                outputSheet1.write(i, 2, p)
                outputSheet1.write(i, 3, OldAB3_1)
                outputSheet1.write(i, 4, NewAB3_1)
                       
                i = i+1
                k = i-1
                m = 0
                p = p+1
            
        elif inputSheet.cell_value(i, 0)!= inputSheet.cell_value(k, 0):
            print(i)
            print(k)
            p = 1
            print "p:",p
            Agency = inputSheet.cell_value(i, 0)
            OldAB3_1 = inputSheet.cell_value(i, 1)        
            print(Agency)
            
            NewAB3_1 = raw_input("Scan 1st Barcode:")
            print "NewAB3_1:",NewAB3_1

            outputSheet1.write(i, 0, Agency)
            outputSheet1.write(i, 1, 'NewAB3 1')        
            outputSheet1.write(i, 2, p)
            outputSheet1.write(i, 3, OldAB3_1)
            outputSheet1.write(i, 4, NewAB3_1)
                   
            i = i+1
            k = i-1
            m = 2
            p = p+1


#############################################################################
##################################################################################
        
wb.save('New Barcode Scans 1.xls') 
