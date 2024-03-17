Sub Tri()

Dim Datearray As Range
Dim Accountarray As Range
Dim Instrumentarray As Range
Dim List_date As New Collection
Dim List_account As New Collection
Dim List_instrument As New Collection

Dim Interarray As Range


Dim date_loop As Variant
Dim account_loop As Variant
Dim instrument_loop As Variant

Dim Date_value As String
Dim Stock_value As String
Dim Volume_value As Variant
Dim Price_value As Variant
Dim Currency_value As String
Dim Client_value As String
Dim test As String
Dim singlebool As Boolean
Dim firstloop As Boolean


Dim final_range As Variant

Dim firstvalue As Range

    
firstloop = True


Worksheets("Sheet").Range("B:B,C:C,D:D,F:F,G:G,J:J,L:L,N:N,O:O,P:P,Q:Q,R:R,S:S,T:T").Delete



Set Datearray = Range(Range("A2"), Range("A2").End(xlDown))


Set List_date = GetUniqueValues(Datearray.Value)



For Each date_loop In List_date
    
    Worksheets("Sheet").Range("A1").AutoFilter Field:=2
    Worksheets("Sheet").Range("A1").AutoFilter Field:=1, Criteria1:=date_loop
    
    Set Accountarray = Range(Range("B2"), Range("B2").End(xlDown)).SpecialCells(xlCellTypeVisible)
    Set List_account = GetUniqueValues(Accountarray.Value)
    
    If Accountarray.Cells.Count > 5000 Then
        List_account.Add Accountarray.Cells(1, 1).Value
    Else
        Set List_account = GetUniqueValues(Accountarray.Value)
            
        
    End If
        
    
    For Each account_loop In List_account
        
        Worksheets("Sheet").Range("A1").AutoFilter Field:=4
        Worksheets("Sheet").Range("A1").AutoFilter Field:=2, Criteria1:=account_loop
        
        Set Instrumentarray = Range(Range("D2"), Range("D2").End(xlDown)).SpecialCells(xlCellTypeVisible)
        
        
        
        If Instrumentarray.Cells.Count > 5000 Then
            List_instrument.Add Instrumentarray.Cells(1, 1).Value
            
        Else
            Set List_instrument = GetUniqueValues(Instrumentarray.Value)
            
        
        End If
        
        

        
        For Each instrument_loop In List_instrument
            Worksheets("Sheet").Range("A1").AutoFilter Field:=4, Criteria1:=instrument_loop
            
            Set firstvalue = Range(Range("D1"), Range("D1").End(xlDown)).Offset(1, 0).SpecialCells(xlCellTypeVisible).Areas(1).Rows(1)


            
            Set Interarray = Range(firstvalue, firstvalue.End(xlDown)).SpecialCells(xlCellTypeVisible)
            
            If Interarray.Cells.Count > 5000 Then
                
                singlebool = True
            Else
                
                singlebool = False
        
            End If
            
            Date_value = firstvalue.Offset(0, -3)
            Stock_value = firstvalue
            Currency_value = firstvalue.Offset(0, 2)
            Client_value = firstvalue.Offset(0, -2)
            Volume_value = Application.WorksheetFunction.Sum(Range(firstvalue.Offset(0, -1), firstvalue.Offset(0, -1).End(xlDown)).SpecialCells(xlCellTypeVisible))
            
            If singlebool = True Then
            Price_value = Application.WorksheetFunction.Sum(Range(firstvalue.Offset(0, 1), firstvalue.Offset(0, 1).End(xlDown)).SpecialCells(xlCellTypeVisible))
            
            Else
            
            Price_value = Application.WorksheetFunction.SumProduct(Range(firstvalue.Offset(0, -1), firstvalue.Offset(0, -1).End(xlDown)).SpecialCells(xlCellTypeVisible), Range(firstvalue.Offset(0, 1), firstvalue.Offset(0, 1).End(xlDown)).SpecialCells(xlCellTypeVisible)) / Volume_value
            End If
            
        
            Range("N2") = Date_value
            Range("O2") = Stock_value
            Range("P2") = Volume_value
            Range("Q2") = Price_value
            Range("R2") = Currency_value
            Range("S2") = Client_value
            
            If firstloop = True Then
                Range("N2:S2").Copy Worksheets("final").Range("A2")
                firstloop = False
            Else
                Range("N2:S2").Copy Worksheets("final").Range("A1").End(xlDown).Offset(1, 0)
            
            End If
            
            
            
            
            
            

            
            
        Next instrument_loop
        
        
    Next account_loop
    
    
Next date_loop


End Sub



Public Function GetUniqueValues(ByVal values As Variant) As Collection
    Dim result As Collection
    Dim cellValue As Variant
    Dim cellValueTrimmed As String

    Set result = New Collection
    Set GetUniqueValues = result

    On Error Resume Next

    For Each cellValue In values
        cellValueTrimmed = Trim(cellValue)
        If cellValueTrimmed = "" Then GoTo NextValue
        result.Add cellValueTrimmed, cellValueTrimmed
NextValue:
    Next cellValue

    On Error GoTo 0
End Function
