import xlrd,xlwt
from xlutils.copy import copy
import sys


def RosterLookup(number):
    #Assuming that we can have the audition number be the row-1
    #Assuming that the first row is the title row
    position=0
    for cell in Input_Roster.row(0):
        if cell.value=='Audition No.':
            aud_col=position
        position+=1
    if number==Input_Roster.cell_value(number+1,aud_col):
        return number+1
    else:
        for index,cell in enumerate(Input_Roster.col(aud_col)):
            if cell.value==number:
                return index
    print'The following number is not in the roster: '+str(number)
    sys.exit()

result = xlwt.Workbook()
Choreographer_Book = xlrd.open_workbook('Choreographers.xlsx')
Input_Roster_Book = xlrd.open_workbook('TestRoster.xlsx')
Input_Roster = Input_Roster_Book.sheet_by_index(0)
Choreo=Choreographer_Book.sheet_by_index(0)
Audition_Roster = result.add_sheet('Audition Roster')

Input_Roster_Namevals=[]   
for x in Input_Roster.col(1):
    Input_Roster_Namevals.append(x.value)

Total_Audition=0
current_number=0
for index,x in enumerate(Input_Roster.col(0)):
    if index!=0:
        if x.ctype==1:
            break
        current_number=x.value
No_List=[]

for x in range(1,int(current_number)+1):
    No_List.append(0)
    

def OfficerLookup(name):
    result=None
    try:
        result=Input_Roster_Namevals.index(name)
    except ValueError:
        print 'The following officer doenst exist '+name
    if result==None:
        print 'The officer '+name+' doesnt exist'
        sys.exit()
    return Input_Roster_Namevals.index(name)

#manually copy all of them
#!Audition_Roster = copy(Input_Roster_Book)
for x in range(0,Input_Roster.nrows):
    for y in range(0,Input_Roster.ncols):
        Audition_Roster.write(x,y,Input_Roster.cell_value(x,y))

#Start by creating the new Sheets

Audition_Roster=result.get_sheet(0)
for x in range(0,Choreo.ncols):
    result.add_sheet(Choreo.cell_value(0,x))
    curr=result.get_sheet(x+1)
    for index,item in enumerate(Input_Roster.row(0)):
        curr.write(0,index,item.value)
    for index_choreo,number in enumerate(Choreo.col(x)):
        if index_choreo!=0:
            if number.ctype==1:
                row_number=OfficerLookup(number.value)
                for index,item in enumerate(Input_Roster.row(row_number)):
                    try:
                        curr.write(index_choreo, index, item.value)
                    except ValueError:
                        print 'the following officer: '+str(number.value)+"is not a valid officer"
            else:
                if number.value!='':
                    row_number=RosterLookup(int(number.value))
                    for index,item in enumerate(Input_Roster.row(row_number)):
                        try:
                            curr.write(index_choreo, index, item.value)
                        except ValueError:
                            print 'the following number: '+str(index)+" is not a valid audition number"
            
curr=result.get_sheet(0)
Input_Roster = Input_Roster_Book.sheet_by_index(0)
Input_Roster_Colvals=[]
Input_Roster_Namevals=[]
Input_Roster_Rowvals=[]
for x in Input_Roster.col(0):
    Input_Roster_Colvals.append(x.value)
for x in Choreo.row(0):
    Input_Roster_Rowvals.append(x.value)
for x in reversed(Input_Roster.col(1)):
    Input_Roster_Namevals.append(x.value)

for index,cell in enumerate(Choreo.row(0)):
    curr.write(0,Input_Roster.ncols+index,cell.value)
for x in range(0,Choreo.ncols):
    st = xlwt.easyxf('pattern: pattern solid;')
    st.pattern.pattern_fore_colour = 2 * x
    for index,number in enumerate(Choreo.col(x)):
        if index!=0:
            if number.value!='':
                if number.ctype==1:
                    try:
                        ins_row = Input_Roster.nrows-Input_Roster_Namevals.index(number.value)
                        ins_col = Input_Roster_Rowvals.index(Choreo.row(0)[x].value)
                        curr.write(ins_row,ins_col+Input_Roster.ncols," ",st)
                    except Exception:
                        print "Error while writing file. Please check if there is a space in one of your texts on if there is a duplicate number"
                        sys.exit()
                else:
                    try:
                        No_List[int(number.value)]+=1
                        ins_row = Input_Roster_Colvals.index(number.value)
                        ins_col = Input_Roster_Rowvals.index(Choreo.row(0)[x].value)
                        curr.write(ins_row,ins_col+Input_Roster.ncols," ",st)
                    except Exception:
                        print "Error while writing file. Please check if there is a space in one of your texts on if there is a duplicate number"
                        sys.exit()
                    

No_Sheet = result.add_sheet('No')
for index,item in enumerate(Input_Roster.row(0)):
        No_Sheet.write(0,index,item.value)

current_row=1
for index,x in enumerate(No_List):
    if x==0:
        try:
            ins_row = Input_Roster_Colvals.index(index)
            for no_index,no_item in enumerate(Input_Roster.row(ins_row)):
                No_Sheet.write(current_row,no_index,no_item.value)
            current_row+=1
        except ValueError:
            print 'the following number: '+str(index)+" is not a valid audition number"
        
    
        
result.save('final.xls')



            
            

