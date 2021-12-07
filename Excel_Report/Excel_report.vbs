Dim oExcelApp, oWorkbook
Dim Name

Push_Data()

Sub Open_Excel()
	Set oExcelApp = CreateObject("Excel.Application")
	oExcelApp.Visible=false
	Set oWorkbook = oExcelApp.Workbooks.Open("Reference_Excel\Book1.xlsx")
End Sub

Sub Close_Excel()
		output = "output\" & Name & ".xlsx"
		oWorkbook.SaveAs output
		oExcelApp.Quit 
		set oExcelApp = nothing
		MsgBox("Excel Report Exported for " & Name)
End Sub

Sub Push_Data()

	 Dim a : a = 16
	 For i = 8 to a Step 2 'i is the counter variable and it is incremented by 2
	 
		Open_Excel()
		Set oSheet = oExcelApp.Sheets(1)
		Set oSheet2 = oExcelApp.Sheets(2)
		
		Name = oSheet2.Cells( i , 1).Value
		
		oSheet.Cells( 7 , 6).Value = Name
		oSheet.Cells( 7 , 11).Value = oSheet2.Cells( 2 , 6).Value
		oSheet.Cells( 9 , 6).Value = oSheet2.Cells( 4 , 6).Value
		oSheet.Cells( 9 , 11).Value = oSheet2.Cells( 3 , 6).Value
		oSheet.Cells( 23 , 6).Value = oSheet2.Cells( 5 , 6).Value
		
		
		oSheet.Cells( 14 , 7).Value = oSheet2.Cells( i , 4).Value
		oSheet.Cells( 14 , 8).Value = oSheet2.Cells( i+1 , 4).Value
		
		oSheet.Cells( 15 , 7).Value = oSheet2.Cells( i , 5).Value
		oSheet.Cells( 15 , 8).Value = oSheet2.Cells( i+1 , 5).Value
		
		oSheet.Cells( 16 , 7).Value = oSheet2.Cells( i , 6).Value
		oSheet.Cells( 16 , 8).Value = oSheet2.Cells( i+1 , 6).Value
		
		oSheet.Cells( 17 , 7).Value = oSheet2.Cells( i , 7).Value
		oSheet.Cells( 17 , 8).Value = oSheet2.Cells( i+1 , 7).Value
		
		
		oSheet.Cells( 18 , 7).Value = oSheet2.Cells( i , 8).Value
		oSheet.Cells( 18 , 8).Value = oSheet2.Cells( i+1 , 8).Value
		
		oSheet.Cells( 19 , 7).Value = oSheet2.Cells( i , 9).Value
		oSheet.Cells( 19 , 8).Value = oSheet2.Cells( i+1 , 9).Value
		
		oSheet.Cells( 23 , 9).Value = oSheet2.Cells( i , 2).Value
		
		Close_Excel()
		
	 Next 
	
	
End Sub

