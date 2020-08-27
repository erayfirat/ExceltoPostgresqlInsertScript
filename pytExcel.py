from openpyxl import Workbook,load_workbook

wb = load_workbook("excelFile.xlsx",  read_only=True)

ws = wb.sheetnames
file= open("postgres.txt","w+")
columnNamesArr=[]
columnNames=''
for name in ws:
    header_cells_generator=wb[name].iter_rows(max_row=1)
    for header_cells_tuple in header_cells_generator:
        for i in range(len(header_cells_tuple)):
            columnNamesArr.append(header_cells_tuple[i].value)            
    columnNames=', '.join(columnNamesArr)    
   
    for row in range(2,wb[name].max_row+1):
        insertValuesArr=[]
        insertValues=''
        for column in range(1,wb[name].max_column+1):       
            if type(wb[name].cell(row,column).value)==str:
                insertValuesArr.append("'"+wb[name].cell(row,column).value+"'")
            else:
                insertValuesArr.append(wb[name].cell(row,column).value)
        
        insertValues=', '.join([str(elem) for elem in insertValuesArr])        
        insertCommand="INSERT INTO "+ name +"("+ columnNames +") VALUES("+insertValues+");"
        file.write(insertCommand+"\n" )
        
print("success")
file.close() 
                
  
            
    
