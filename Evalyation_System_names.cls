VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Sub Gading_Sheet_Click()

''''''''''''''''''''''''''''''''''''''''''
'Code name: Training Evaluation Template '
'Author :   Mathew Bidinlib              '
'Created:      30th April, 2019          '
' Modified:    2nd May, 2019             '
''''''''''''''''''''''''''''''''''''''''''



' 1 Copy Criterial sheet

If Sheets("instructions").Range("C28").Value > 0 Then

    
    Sheets("criteria_template").Copy after:=Sheets("instructions")
    Sheets("criteria_template (2)").Visible = True
    
    'updating Column B and G
    Sheets("criteria_template (2)").Range("B2").Value = Sheets("instructions").Range("B13").Value
    Sheets("criteria_template (2)").Range("B4").Value = Sheets("instructions").Range("B14").Value
    Sheets("criteria_template (2)").Range("G2").Value = Sheets("instructions").Range("D13").Value
    Sheets("criteria_template (2)").Range("G4").Value = Sheets("instructions").Range("D14").Value
    Sheets("criteria_template (2)").Range("H2").Value = Sheets("instructions").Range("E13").Value
    Sheets("criteria_template (2)").Range("H4").Value = Sheets("instructions").Range("E14").Value
    
    For i = 1 To 5
        Sheets("criteria_template (2)").Range("B" & i * 10 + 4).Value = Sheets("instructions").Range("B" & i + 14).Value
        Sheets("criteria_template (2)").Range("G" & i * 10 + 4).Value = Sheets("instructions").Range("D" & i + 14).Value
        Sheets("criteria_template (2)").Range("H" & i * 10 + 4).Value = Sheets("instructions").Range("E" & i + 14).Value
    Next i
    
    For i = 1 To 4
        Sheets("criteria_template (2)").Range("B" & i + 54).Value = Sheets("instructions").Range("B" & i + 18).Value
        Sheets("criteria_template (2)").Range("G" & i + 54).Value = Sheets("instructions").Range("D" & i + 18).Value
        Sheets("criteria_template (2)").Range("H" & i + 54).Value = Sheets("instructions").Range("E" & i + 18).Value
    Next i
    
    Sheets("criteria_template (2)").Name = "Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value

End If


'''''''''''''''''''''''''''''''
'2. Copying the Matrix Sheets '
'''''''''''''''''''''''''''''''

' Declaring a loop through all assessors

Dim rng As Range: Set rng = Application.Range("instructions!C30:C34")
  

For k = 1 To rng.Rows.Count
    
    Dim names As String
    names = rng.Cells(RowIndex:=k, ColumnIndex:=1).Value
    Dim f_name As String
    f_name = rng.Cells(RowIndex:=1, ColumnIndex:=1).Value

 ' MsgBox rng.Cells(RowIndex:=i, ColumnIndex:=1).Value
    If names <> "" Then
    
                 
        For i = 1 To Sheets("instructions").Range("C28").Value
        
            Sheets("matrix_template").Copy after:=Sheets("instructions")
            
            Sheets("matrix_template (" & i + 1 & ")").Visible = True
        
            'Sheets("matrix_template (" & i + 1 & ")").Visible = True
            
        
            Sheets("matrix_template (" & i + 1 & ")").Range("B1").Value = Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value
            
            'Updating the details of the project
            
            Sheets("matrix_template (" & i + 1 & ")").Range("C2").Value = Sheets("instructions").Range("C3").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("G2").Value = Sheets("instructions").Range("C5").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("J2").Value = Sheets("instructions").Range("C6").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("O2").Value = Sheets("instructions").Range("C26").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("Q2").Value = Sheets("instructions").Range("C27").Value
        
            ' Updating the various criteria
        
            Sheets("matrix_template (" & i + 1 & ")").Range("C3").Value = Sheets("instructions").Range("B13").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("E3").Value = Sheets("instructions").Range("B14").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("G3").Value = Sheets("instructions").Range("B15").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("I3").Value = Sheets("instructions").Range("B16").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("K3").Value = Sheets("instructions").Range("B17").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("M3").Value = Sheets("instructions").Range("B18").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("O3").Value = Sheets("instructions").Range("B19").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("Q3").Value = Sheets("instructions").Range("B20").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("S3").Value = Sheets("instructions").Range("B21").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("U3").Value = Sheets("instructions").Range("B22").Value
            Sheets("matrix_template (" & i + 1 & ")").Range("W3").Value = Sheets("instructions").Range("B23").Value
        
            
            
            'Adding the Serial Numbers
           
            
            For j = 1 To Sheets("instructions").Range("C10").Value
              
              If Sheets("instructions").Range("C10").Value <> "" Then
                  Sheets("matrix_template (" & i + 1 & ")").Range("A" & j + 4).Value = j
               End If
               
            
            Next j
        
            
                    
            'Adding and autofilling  formular to calculate the wieghted score
                 
                        
            Dim rg As Double
            rg = Sheets("instructions").Range("C10").Value + 4
            
            If Sheets("instructions").Range("C10").Value <> "" Then
                  Sheets("matrix_template (" & i + 1 & ")").Range("D5").Formula = "= C5 * (instructions!$E$13/100)"
                                            
                  Sheets("matrix_template (" & i + 1 & ")").Range("D5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("D5:D" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("F5").Formula = "= E5 * (instructions!$E$14/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("F5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("F5:F" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("H5").Formula = "= G5 * (instructions!$E$15/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("H5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("H5:H" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("J5").Formula = "= I5 * (instructions!$E$16/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("J5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("J5:J" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("L5").Formula = "= K5 * (instructions!$E$17/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("L5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("L5:L" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("N5").Formula = "= M5 * (instructions!$E$18/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("N5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("N5:N" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("P5").Formula = "= O5 * (instructions!$E$19/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("P5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("P5:P" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("R5").Formula = "= Q5 * (instructions!$E$20/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("R5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("R5:R" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("T5").Formula = "= S5 * (instructions!$E$21/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("T5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("T5:T" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("V5").Formula = "= U5 * (instructions!$E$22/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("V5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("V5:V" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("X5").Formula = "= W5 * (instructions!$E$23/100)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("X5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("X5:X" & rg)
            
            'Adding the Overall Socre and ranking for all STS
                  Sheets("matrix_template (" & i + 1 & ")").Range("Y5").Formula = "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
                  Sheets("matrix_template (" & i + 1 & ")").Range("Y5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("Y5:Y" & rg)
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("Z5").Formula = "=RANK(Y5,Y:Y)"
                  Sheets("matrix_template (" & i + 1 & ")").Range("Z5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("Z5:Z" & rg)
                  
                  
                  Sheets("matrix_template (" & i + 1 & ")").Range("AA5").Formula = "=IF(Y5 >= instructions!$D$25,1,0 )"
                  Sheets("matrix_template (" & i + 1 & ")").Range("AA5").AutoFill Destination:=Sheets("matrix_template (" & i + 1 & ")").Range("AA5:AA" & rg)
            
            
            End If
            
            
        Next i
                      
        
        'Naming all the the Matrix Sheets
        
             
        For i = 2 To Sheets("instructions").Range("C28").Value + 1
                
            If 1 >= 3 Then
                For j = 5 To Sheets("instructions").Range("C10").Value + 4
                       
                    Sheets("matrix_template (" & i + 1 & ")").Range("B5").Formula = _
                    "=" & "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & f_name & "_" & "Day_" & 1 & "!B" & j
                    
                    '"=matrix_template(2)!B1"
                
               Next j
        
          End If
                
            Sheets("matrix_template (" & i & ")").Name = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names & "_" & "Day_" & i - 1
             
                       
           
        Next i
        
        
                   'coping the names of the surveyors to the rest of the days
            
            For i = 2 To Sheets("instructions").Range("C28").Value
                
        
                Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names & "_" & "Day_" & i).Range("B5").Formula = _
                   "=" & "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names & "_" & "Day_" & 1 & "!B" & 5
                  
               Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names & "_" & "Day_" & i).Range("B5").AutoFill Destination:= _
               Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names & "_" & "Day_" & i).Range("B5:B" & rg)

 
           Next i
         
    ' end the fist if command
    End If


Next k


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3 Creating an Overal Sheet to cumulate gradings for multiple grading days''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Declaring the loop for all the sheets


For w = 1 To rng.Rows.Count
    
    Dim names_n As String
    names_n = rng.Cells(RowIndex:=w, ColumnIndex:=1).Value
    Dim f_name_n As String
    f_name = rng.Cells(RowIndex:=1, ColumnIndex:=1).Value
    
    If names_n <> "" Then



            ' Filling preliminary information
            
            Sheets("overall_template").Copy after:=Sheets("instructions")
            Sheets("overall_template (2)").Visible = True
            
            Sheets("overall_template (2)").Range("B1").Value = Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value
            
            Sheets("overall_template (2)").Range("C2").Value = Sheets("instructions").Range("C3").Value
            Sheets("overall_template (2)").Range("G2").Value = Sheets("instructions").Range("C5").Value
            Sheets("overall_template (2)").Range("J2").Value = Sheets("instructions").Range("C6").Value
            Sheets("overall_template (2)").Range("O2").Value = Sheets("instructions").Range("C26").Value
            Sheets("overall_template (2)").Range("Q2").Value = Sheets("instructions").Range("C27").Value
            
            Sheets("overall_template (2)").Range("A4").Value = "SN"
            
            For i = 1 To Sheets("instructions").Range("C10").Value
              
                If Sheets("instructions").Range("C10").Value <> "" Then
                    Sheets("overall_template (2)").Range("A" & i + 4).Value = i
                 End If
                
            Next i





' storing sheet names for use (locals)
        
        Dim sht1 As String
        
        sht1 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 1
        sht2 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 2
        sht3 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 3
        sht4 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 4
        sht5 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 5
        sht6 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 6
        sht7 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 7
        sht8 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 8
        sht9 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 9
        sht10 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 10



        'Cumulate all raw scores into overall grading score
        
        'Dim y As Double
        
        y = Sheets("instructions").Range("C28").Value
          
        For k = 5 To Sheets("instructions").Range("C10").Value + 4
                   
              Sheets("overall_template (2)").Range("B" & k).Formula = "=" & sht1 & "!B" & k
                
               
          ' Cummulating results based on the number of sheets
            If y = 10 Then
            
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & "+" & (sht6 & "!C" & k) & "+" & (sht7 & "!C" & k) & "+" & (sht8 & "!C" & k) & "+" & (sht9 & "!C" & k) & "+" & (sht10 & "!C" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & "+" & (sht6 & "!E" & k) & "+" & (sht7 & "!E" & k) & "+" & (sht8 & "!E" & k) & "+" & (sht9 & "!E" & k) & "+" & (sht10 & "!E" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & "+" & (sht6 & "!G" & k) & "+" & (sht7 & "!G" & k) & "+" & (sht8 & "!G" & k) & "+" & (sht9 & "!G" & k) & "+" & (sht10 & "!G" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & "+" & (sht6 & "!I" & k) & "+" & (sht7 & "!I" & k) & "+" & (sht8 & "!I" & k) & "+" & (sht9 & "!I" & k) & "+" & (sht10 & "!I" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & "+" & (sht6 & "!K" & k) & "+" & (sht7 & "!K" & k) & "+" & (sht8 & "!K" & k) & "+" & (sht9 & "!K" & k) & "+" & (sht10 & "!K" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & "+" & (sht6 & "!M" & k) & "+" & (sht7 & "!M" & k) & "+" & (sht8 & "!M" & k) & "+" & (sht9 & "!M" & k) & "+" & (sht10 & "!M" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & "+" & (sht6 & "!O" & k) & "+" & (sht7 & "!O" & k) & "+" & (sht8 & "!O" & k) & "+" & (sht9 & "!O" & k) & "+" & (sht10 & "!O" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & "+" & (sht6 & "!Q" & k) & "+" & (sht7 & "!Q" & k) & "+" & (sht8 & "!Q" & k) & "+" & (sht9 & "!Q" & k) & "+" & (sht10 & "!Q" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & "+" & (sht6 & "!S" & k) & "+" & (sht7 & "!S" & k) & "+" & (sht8 & "!S" & k) & "+" & (sht9 & "!S" & k) & "+" & (sht10 & "!S" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & "+" & (sht6 & "!U" & k) & "+" & (sht7 & "!U" & k) & "+" & (sht8 & "!U" & k) & "+" & (sht9 & "!U" & k) & "+" & (sht10 & "!U" & k) & ")" & "/" & 10
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & "+" & (sht6 & "!W" & k) & "+" & (sht7 & "!W" & k) & "+" & (sht8 & "!W" & k) & "+" & (sht9 & "!W" & k) & "+" & (sht10 & "!W" & k) & ")" & "/" & 10
                
                ElseIf y = 9 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & "+" & (sht6 & "!C" & k) & "+" & (sht7 & "!C" & k) & "+" & (sht8 & "!C" & k) & "+" & (sht9 & "!C" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & "+" & (sht6 & "!E" & k) & "+" & (sht7 & "!E" & k) & "+" & (sht8 & "!E" & k) & "+" & (sht9 & "!E" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & "+" & (sht6 & "!G" & k) & "+" & (sht7 & "!G" & k) & "+" & (sht8 & "!G" & k) & "+" & (sht9 & "!G" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & "+" & (sht6 & "!I" & k) & "+" & (sht7 & "!I" & k) & "+" & (sht8 & "!I" & k) & "+" & (sht9 & "!I" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & "+" & (sht6 & "!K" & k) & "+" & (sht7 & "!K" & k) & "+" & (sht8 & "!K" & k) & "+" & (sht9 & "!K" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & "+" & (sht6 & "!M" & k) & "+" & (sht7 & "!M" & k) & "+" & (sht8 & "!M" & k) & "+" & (sht9 & "!M" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & "+" & (sht6 & "!O" & k) & "+" & (sht7 & "!O" & k) & "+" & (sht8 & "!O" & k) & "+" & (sht9 & "!O" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & "+" & (sht6 & "!Q" & k) & "+" & (sht7 & "!Q" & k) & "+" & (sht8 & "!Q" & k) & "+" & (sht9 & "!Q" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & "+" & (sht6 & "!S" & k) & "+" & (sht7 & "!S" & k) & "+" & (sht8 & "!S" & k) & "+" & (sht9 & "!S" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & "+" & (sht6 & "!U" & k) & "+" & (sht7 & "!U" & k) & "+" & (sht8 & "!U" & k) & "+" & (sht9 & "!U" & k) & ")" & "/" & 9
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & "+" & (sht6 & "!W" & k) & "+" & (sht7 & "!W" & k) & "+" & (sht8 & "!W" & k) & "+" & (sht9 & "!W" & k) & ")" & "/" & 9
            
                ElseIf y = 8 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & "+" & (sht6 & "!C" & k) & "+" & (sht7 & "!C" & k) & "+" & (sht8 & "!C" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & "+" & (sht6 & "!E" & k) & "+" & (sht7 & "!E" & k) & "+" & (sht8 & "!E" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & "+" & (sht6 & "!G" & k) & "+" & (sht7 & "!G" & k) & "+" & (sht8 & "!G" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & "+" & (sht6 & "!I" & k) & "+" & (sht7 & "!I" & k) & "+" & (sht8 & "!I" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & "+" & (sht6 & "!K" & k) & "+" & (sht7 & "!K" & k) & "+" & (sht8 & "!K" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & "+" & (sht6 & "!M" & k) & "+" & (sht7 & "!M" & k) & "+" & (sht8 & "!M" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & "+" & (sht6 & "!O" & k) & "+" & (sht7 & "!O" & k) & "+" & (sht8 & "!O" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & "+" & (sht6 & "!Q" & k) & "+" & (sht7 & "!Q" & k) & "+" & (sht8 & "!Q" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & "+" & (sht6 & "!S" & k) & "+" & (sht7 & "!S" & k) & "+" & (sht8 & "!S" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & "+" & (sht6 & "!U" & k) & "+" & (sht7 & "!U" & k) & "+" & (sht8 & "!U" & k) & ")" & "/" & 8
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & "+" & (sht6 & "!W" & k) & "+" & (sht7 & "!W" & k) & "+" & (sht8 & "!W" & k) & ")" & "/" & 8
                
                ElseIf y = 7 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & "+" & (sht6 & "!C" & k) & "+" & (sht7 & "!C" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & "+" & (sht6 & "!E" & k) & "+" & (sht7 & "!E" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & "+" & (sht6 & "!G" & k) & "+" & (sht7 & "!G" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & "+" & (sht6 & "!I" & k) & "+" & (sht7 & "!I" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & "+" & (sht6 & "!K" & k) & "+" & (sht7 & "!K" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & "+" & (sht6 & "!M" & k) & "+" & (sht7 & "!M" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & "+" & (sht6 & "!O" & k) & "+" & (sht7 & "!O" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & "+" & (sht6 & "!Q" & k) & "+" & (sht7 & "!Q" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & "+" & (sht6 & "!S" & k) & "+" & (sht7 & "!S" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & "+" & (sht6 & "!U" & k) & "+" & (sht7 & "!U" & k) & ")" & "/" & 7
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & "+" & (sht6 & "!W" & k) & "+" & (sht7 & "!W" & k) & ")" & "/" & 7
                
                ElseIf y = 6 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & "+" & (sht6 & "!C" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & "+" & (sht6 & "!E" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & "+" & (sht6 & "!G" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & "+" & (sht6 & "!I" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & "+" & (sht6 & "!K" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & "+" & (sht6 & "!M" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & "+" & (sht6 & "!O" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & "+" & (sht6 & "!Q" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & "+" & (sht6 & "!S" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & "+" & (sht6 & "!U" & k) & ")" & "/" & 6
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & "+" & (sht6 & "!W" & k) & ")" & "/" & 6
                
                ElseIf y = 5 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & ")" & "/" & 5
                
                ElseIf y = 4 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & ")" & "/" & 4
                
                ElseIf y = 3 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & ")" & "/" & 3
                
                ElseIf y = 2 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("E" & k).Formula = "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("G" & k).Formula = "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("I" & k).Formula = "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("K" & k).Formula = "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("M" & k).Formula = "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("O" & k).Formula = "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("Q" & k).Formula = "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("S" & k).Formula = "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("U" & k).Formula = "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("W" & k).Formula = "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & ")" & "/" & 2
                
            End If
        
            
        Next k
            
            
            'Adding and autofilling  formular to calculate the wieghted score
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             
           
            Dim rg_1 As Double
            
            
            rg_1 = Sheets("instructions").Range("C10") + 4
            
              Sheets("overall_template (2)").Range("D5").Formula = "= C5 * (instructions!$E$13/100)"
              Sheets("overall_template (2)").Range("D5").AutoFill Destination:=Sheets("overall_template (2)").Range("D5:D" & rg_1)
              
              Sheets("overall_template (2)").Range("F5").Formula = "= E5* (instructions!$E$14/100)"
              Sheets("overall_template (2)").Range("F5").AutoFill Destination:=Sheets("overall_template (2)").Range("F5:F" & rg_1)
              
              Sheets("overall_template (2)").Range("H5").Formula = "= G5* (instructions!$E$15/100)"
              Sheets("overall_template (2)").Range("H5").AutoFill Destination:=Sheets("overall_template (2)").Range("H5:H" & rg_1)
              
              Sheets("overall_template (2)").Range("J5").Formula = "= I5 * (instructions!$E$16/100)"
              Sheets("overall_template (2)").Range("J5").AutoFill Destination:=Sheets("overall_template (2)").Range("J5:J" & rg_1)
              
              Sheets("overall_template (2)").Range("L5").Formula = "= K5  * (instructions!$E$17/100)"
              Sheets("overall_template (2)").Range("L5").AutoFill Destination:=Sheets("overall_template (2)").Range("L5:L" & rg_1)
              
              Sheets("overall_template (2)").Range("N5").Formula = "= M5  * (instructions!$E$18/100)"
              Sheets("overall_template (2)").Range("N5").AutoFill Destination:=Sheets("overall_template (2)").Range("N5:N" & rg_1)
              
              Sheets("overall_template (2)").Range("P5").Formula = "= O5  * (instructions!$E$19/100)"
              Sheets("overall_template (2)").Range("P5").AutoFill Destination:=Sheets("overall_template (2)").Range("P5:P" & rg_1)
              
              Sheets("overall_template (2)").Range("R5").Formula = "= Q5  * (instructions!$E$20/100)"
              Sheets("overall_template (2)").Range("R5").AutoFill Destination:=Sheets("overall_template (2)").Range("R5:R" & rg_1)
              
              Sheets("overall_template (2)").Range("T5").Formula = "= S5  * (instructions!$E$21/100)"
              Sheets("overall_template (2)").Range("T5").AutoFill Destination:=Sheets("overall_template (2)").Range("T5:T" & rg_1)
              
              Sheets("overall_template (2)").Range("V5").Formula = "= U5  * (instructions!$E$22/100)"
              Sheets("overall_template (2)").Range("V5").AutoFill Destination:=Sheets("overall_template (2)").Range("V5:V" & rg_1)
              
              Sheets("overall_template (2)").Range("X5").Formula = "= W5  * (instructions!$E$23/100)"
              Sheets("overall_template (2)").Range("X5").AutoFill Destination:=Sheets("overall_template (2)").Range("X5:X" & rg_1)
        
             'Adding the Overall Socre and ranking for all STS
        
              Sheets("overall_template (2)").Range("Y5").Formula = "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
              Sheets("overall_template (2)").Range("Y5").AutoFill Destination:=Sheets("overall_template (2)").Range("Y5:Y" & rg_1)
              
              Sheets("overall_template (2)").Range("Z5").Formula = "=RANK(Y5,Y:Y)"
              Sheets("overall_template (2)").Range("Z5").AutoFill Destination:=Sheets("overall_template (2)").Range("Z5:Z" & rg_1)
              
            
              Sheets("overall_template (2)").Range("AA5").Formula = "=IF(Y5 >= instructions!$D$25,1,0)"
              Sheets("overall_template (2)").Range("AA5").AutoFill Destination:=Sheets("overall_template (2)").Range("AA5:AA" & rg_1)
                    
        
            ' Add overall sheet names
        Sheets("overall_template (2)").Name = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n


      End If


Next w
       
       
       
    ' Overall Master
       

    Sheets("overall_template").Copy after:=Sheets("instructions")
    Sheets("overall_template (2)").Visible = True
    
    Sheets("overall_template (2)").Range("B1").Value = Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value
    
    Sheets("overall_template (2)").Range("C2").Value = Sheets("instructions").Range("C3").Value
    Sheets("overall_template (2)").Range("G2").Value = Sheets("instructions").Range("C5").Value
    Sheets("overall_template (2)").Range("J2").Value = Sheets("instructions").Range("C6").Value
    Sheets("overall_template (2)").Range("O2").Value = Sheets("instructions").Range("C26").Value
    Sheets("overall_template (2)").Range("Q2").Value = Sheets("instructions").Range("C27").Value
    
    Sheets("overall_template (2)").Range("A4").Value = "SN"
    
    For i = 1 To Sheets("instructions").Range("C10").Value
      
        If Sheets("instructions").Range("C10").Value <> "" Then
            Sheets("overall_template (2)").Range("A" & i + 4).Value = i
         End If
        
    Next i


     Dim n1 As String
     Dim n2 As String
     Dim n3 As String
     Dim n4 As String
     Dim n5 As String

    Dim n1_c As Integer
    Dim n2_c As Integer
    Dim n3_c As Integer
    Dim n4_c As Integer
    Dim n5_c As Integer
    


' storing sheet names for use (locals)

     Dim n_st1 As String
     Dim n_sht2 As String
     Dim n_sht3 As String
     Dim n_sht4 As String
     Dim n_sht5 As String
     
     
Dim rng_n As Range: Set rng_n = Application.Range("instructions!C30:C34")

    n1 = rng_n.Cells(RowIndex:=1, ColumnIndex:=1).Value
    n2 = rng_n.Cells(RowIndex:=2, ColumnIndex:=1).Value
    n3 = rng_n.Cells(RowIndex:=3, ColumnIndex:=1).Value
    n4 = rng_n.Cells(RowIndex:=4, ColumnIndex:=1).Value
    n5 = rng_n.Cells(RowIndex:=5, ColumnIndex:=1).Value


'Verify if all five cells for assessors have values and set sheet names of their overals as locals

If n1 <> "" Then
     n_sht1 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=1, ColumnIndex:=1).Value
     n1_c = 1 ' if there's an assesor in the sheet
Else:
n_sht1 = "no_inp"
n1_c = 0
End If

 If n2 <> "" Then
     n_sht2 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=2, ColumnIndex:=1).Value
     n2_c = 1
Else:
    n_sht2 = "no_inp"
    n2_c = 0
End If

 If n3 <> "" Then
     n_sht3 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=3, ColumnIndex:=1).Value
     n3_c = 1 ' if there's an assesor in the sheet
Else:
    n_sht3 = "no_inp"
    n3_c = 0
End If

 If n4 <> "" Then
     n_sht4 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=4, ColumnIndex:=1).Value
     n4_c = 1 ' if there's an assesor in the sheet
Else:
    n_sht4 = "no_inp"
    n4_c = 0
End If

 If n5 <> "" Then
     n_sht5 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=5, ColumnIndex:=1).Value
     n5_c = 1
Else:
    n_sht5 = "no_inp"
    n5_c = 0
End If


' Declare the variable that counts the total number of Assessors

Dim n_count As Integer
    n_count = n1_c + n2_c + n3_c + n4_c + n5_c


'Cumulate all raw scores into overall grading score

'Dim y As Double
  
For k = 5 To Sheets("instructions").Range("C10").Value + 4
           
      Sheets("overall_template (2)").Range("B" & k).Formula = "=" & sht1 & "!B" & k
        
  ' Adding the overall sheets
        Sheets("overall_template (2)").Range("C" & k).Formula = _
        "=" & "(" & (n_sht1 & "!C" & k) & "+" & (n_sht2 & "!C" & k) & "+" & (n_sht3 & "!C" & k) & "+" & (n_sht4 & "!C" & k) & "+" & (n_sht5 & "!C" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("E" & k).Formula = _
        "=" & "(" & (n_sht1 & "!E" & k) & "+" & (n_sht2 & "!E" & k) & "+" & (n_sht3 & "!E" & k) & "+" & (n_sht4 & "!E" & k) & "+" & (n_sht5 & "!E" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("G" & k).Formula = _
         "=" & "(" & (n_sht1 & "!G" & k) & "+" & (n_sht2 & "!G" & k) & "+" & (n_sht3 & "!G" & k) & "+" & (n_sht4 & "!G" & k) & "+" & (n_sht5 & "!G" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("I" & k).Formula = _
        "=" & "(" & (n_sht1 & "!I" & k) & "+" & (n_sht2 & "!I" & k) & "+" & (n_sht3 & "!I" & k) & "+" & (n_sht4 & "!I" & k) & "+" & (n_sht5 & "!I" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("K" & k).Formula = _
         "=" & "(" & (n_sht1 & "!K" & k) & "+" & (n_sht2 & "!K" & k) & "+" & (n_sht3 & "!K" & k) & "+" & (n_sht4 & "!K" & k) & "+" & (n_sht5 & "!K" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("M" & k).Formula = _
        "=" & "(" & (n_sht1 & "!M" & k) & "+" & (n_sht2 & "!M" & k) & "+" & (n_sht3 & "!M" & k) & "+" & (n_sht4 & "!M" & k) & "+" & (n_sht5 & "!M" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("O" & k).Formula = _
        "=" & "(" & (n_sht1 & "!O" & k) & "+" & (n_sht2 & "!O" & k) & "+" & (n_sht3 & "!O" & k) & "+" & (n_sht4 & "!O" & k) & "+" & (n_sht5 & "!O" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("Q" & k).Formula = _
        "=" & "(" & (n_sht1 & "!Q" & k) & "+" & (n_sht2 & "!Q" & k) & "+" & (n_sht3 & "!Q" & k) & "+" & (n_sht4 & "!Q" & k) & "+" & (n_sht5 & "!Q" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("S" & k).Formula = _
        "=" & "(" & (n_sht1 & "!S" & k) & "+" & (n_sht2 & "!S" & k) & "+" & (n_sht3 & "!S" & k) & "+" & (n_sht4 & "!S" & k) & "+" & (n_sht5 & "!S" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("U" & k).Formula = _
        "=" & "(" & (n_sht1 & "!U" & k) & "+" & (n_sht2 & "!U" & k) & "+" & (n_sht3 & "!U" & k) & "+" & (n_sht4 & "!U" & k) & "+" & (n_sht5 & "!U" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("W" & k).Formula = _
        "=" & "(" & (n_sht1 & "!W" & k) & "+" & (n_sht2 & "!W" & k) & "+" & (n_sht3 & "!W" & k) & "+" & (n_sht4 & "!W" & k) & "+" & (n_sht5 & "!W" & k) & ")" & "/" & n_count
 Next k
 
 
 Dim rg_n As Double
            
            
            rg_n = Sheets("instructions").Range("C10") + 4
            
              Sheets("overall_template (2)").Range("D5").Formula = "= C5 * (instructions!$E$13/100)"
              Sheets("overall_template (2)").Range("D5").AutoFill Destination:=Sheets("overall_template (2)").Range("D5:D" & rg_n)
              
              Sheets("overall_template (2)").Range("F5").Formula = "= E5* (instructions!$E$14/100)"
              Sheets("overall_template (2)").Range("F5").AutoFill Destination:=Sheets("overall_template (2)").Range("F5:F" & rg_n)
              
              Sheets("overall_template (2)").Range("H5").Formula = "= G5* (instructions!$E$15/100)"
              Sheets("overall_template (2)").Range("H5").AutoFill Destination:=Sheets("overall_template (2)").Range("H5:H" & rg_n)
              
              Sheets("overall_template (2)").Range("J5").Formula = "= I5 * (instructions!$E$16/100)"
              Sheets("overall_template (2)").Range("J5").AutoFill Destination:=Sheets("overall_template (2)").Range("J5:J" & rg_n)
              
              Sheets("overall_template (2)").Range("L5").Formula = "= K5  * (instructions!$E$17/100)"
              Sheets("overall_template (2)").Range("L5").AutoFill Destination:=Sheets("overall_template (2)").Range("L5:L" & rg_n)
              
              Sheets("overall_template (2)").Range("N5").Formula = "= M5  * (instructions!$E$18/100)"
              Sheets("overall_template (2)").Range("N5").AutoFill Destination:=Sheets("overall_template (2)").Range("N5:N" & rg_n)
              
              Sheets("overall_template (2)").Range("P5").Formula = "= O5  * (instructions!$E$19/100)"
              Sheets("overall_template (2)").Range("P5").AutoFill Destination:=Sheets("overall_template (2)").Range("P5:P" & rg_n)
              
              Sheets("overall_template (2)").Range("R5").Formula = "= Q5  * (instructions!$E$20/100)"
              Sheets("overall_template (2)").Range("R5").AutoFill Destination:=Sheets("overall_template (2)").Range("R5:R" & rg_n)
              
              Sheets("overall_template (2)").Range("T5").Formula = "= S5  * (instructions!$E$21/100)"
              Sheets("overall_template (2)").Range("T5").AutoFill Destination:=Sheets("overall_template (2)").Range("T5:T" & rg_n)
              
              Sheets("overall_template (2)").Range("V5").Formula = "= U5  * (instructions!$E$22/100)"
              Sheets("overall_template (2)").Range("V5").AutoFill Destination:=Sheets("overall_template (2)").Range("V5:V" & rg_n)
              
              Sheets("overall_template (2)").Range("X5").Formula = "= W5  * (instructions!$E$23/100)"
              Sheets("overall_template (2)").Range("X5").AutoFill Destination:=Sheets("overall_template (2)").Range("X5:X" & rg_n)
        
             'Adding the Overall Socre and ranking for all STS
        
              Sheets("overall_template (2)").Range("Y5").Formula = "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
              Sheets("overall_template (2)").Range("Y5").AutoFill Destination:=Sheets("overall_template (2)").Range("Y5:Y" & rg_n)
              
              Sheets("overall_template (2)").Range("Z5").Formula = "=RANK(Y5,Y:Y)"
              Sheets("overall_template (2)").Range("Z5").AutoFill Destination:=Sheets("overall_template (2)").Range("Z5:Z" & rg_n)
              
            
              Sheets("overall_template (2)").Range("AA5").Formula = "=IF(Y5 >= instructions!$D$25,1,0)"
              Sheets("overall_template (2)").Range("AA5").AutoFill Destination:=Sheets("overall_template (2)").Range("AA5:AA" & rg_n)
            
 
 

' Naming the Master overall

Sheets("overall_template (2)").Name = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & "all"

      
       

'Store the number of assaments in code woeksheet
    Sheets("codes").Range("A1").Value = Sheets("instructions").Range("C28").Value
      
    ' MsgBox n_sht5
    
    
End Sub





Private Sub Update_Overall_Click()



' 1 Update Criterial sheet
   
    'updating Column B and G
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("B2").Value = _
    Sheets("instructions").Range("B13").Value
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("G4").Value = _
    Sheets("instructions").Range("B14").Value
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("G2").Value = _
    Sheets("instructions").Range("D13").Value
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("G4").Value = _
    Sheets("instructions").Range("D14").Value
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("H2").Value = _
    Sheets("instructions").Range("E13").Value
    Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("H4").Value = _
    Sheets("instructions").Range("E14").Value
    
    For i = 1 To 5
        Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("B" & i * 10 + 4).Value = _
        Sheets("instructions").Range("B" & i + 14).Value
        Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("G" & i * 10 + 4).Value = _
        Sheets("instructions").Range("D" & i + 14).Value
        Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("H" & i * 10 + 4).Value = _
        Sheets("instructions").Range("E" & i + 14).Value
    Next i
    
    For i = 1 To 4
        Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("G" & i + 54).Value = _
        Sheets("instructions").Range("D" & i + 18).Value
        Sheets("Criteria_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value).Range("H" & i + 54).Value = _
        Sheets("instructions").Range("E" & i + 18).Value
    Next i
    



'''''''''''''''''''''''''''''''
'2. Copying the Matrix Sheets '
'''''''''''''''''''''''''''''''

    '2a: Verify the number of assesment days and delete or add extra sheets
    
Dim rng_2 As Range: Set rng_2 = Application.Range("instructions!C30:C34")
  

For y = 1 To rng_2.Rows.Count
    
    Dim names_2 As String
    names_2 = rng_2.Cells(RowIndex:=y, ColumnIndex:=1).Value

 ' MsgBox rng.Cells(RowIndex:=i, ColumnIndex:=1).Value
    If names_2 <> "" Then
   
        ' Deleting extra days of evaluation if evaluation days have reduced
        
        If Sheets("codes").Range("A1").Value > Sheets("instructions").Range("C28").Value Then
         
            Dim z As Double
            z = Sheets("codes").Range("A1").Value - Sheets("instructions").Range("C28").Value
            
            For i = Sheets("instructions").Range("C28").Value + 1 To Sheets("codes").Range("A1").Value
                Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Delete
            Next i
            
            ' Adding extra days of evaluation if evaluation days have increased
            
         ElseIf Sheets("codes").Range("A1").Value < Sheets("instructions").Range("C28").Value Then
            
                For i = 1 To (Sheets("instructions").Range("C28").Value - Sheets("codes").Range("A1").Value)
            
                    Sheets("matrix_template").Copy after:=Sheets("instructions")
                    Sheets("matrix_template (" & i + 1 & ")").Visible = True
                    
               Next i
                  
                  For w = 1 To (Sheets("instructions").Range("C28").Value - Sheets("codes").Range("A1").Value)
            
                    Sheets("matrix_template (" & w + 1 & ")").Name = _
                    "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & (w + Sheets("codes").Range("A1").Value)
               Next w
         
           
        End If
        
    End If

Next y
        


For p = 1 To rng_2.Rows.Count
    

 ' MsgBox rng.Cells(RowIndex:=i, ColumnIndex:=1).Value
    If names_2 <> "" Then

        For i = 1 To Sheets("instructions").Range("C28").Value
        
           
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Visible = True
        
        
        
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("B1").Value = _
            Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value
            
            'updating the title of the matrix sheets
            
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("C2").Value = _
            Sheets("instructions").Range("C3").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("G2").Value = _
            Sheets("instructions").Range("C5").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("J2").Value = _
            Sheets("instructions").Range("C6").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("O2").Value = _
            Sheets("instructions").Range("C26").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Q2").Value = _
            Sheets("instructions").Range("C27").Value
            
            ' Updating the criteria names
            
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("C3").Value = _
            Sheets("instructions").Range("B13").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("E3").Value = _
            Sheets("instructions").Range("B14").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("G3").Value = _
            Sheets("instructions").Range("B15").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("I3").Value = _
            Sheets("instructions").Range("B16").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("K3").Value = _
            Sheets("instructions").Range("B17").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("M3").Value = _
            Sheets("instructions").Range("B18").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("O3").Value = _
            Sheets("instructions").Range("B19").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Q3").Value = _
            Sheets("instructions").Range("B20").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("S3").Value = _
            Sheets("instructions").Range("B21").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("U3").Value = _
            Sheets("instructions").Range("B22").Value
            Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("W3").Value = _
            Sheets("instructions").Range("B23").Value
         
            
            For j = 1 To Sheets("instructions").Range("C10").Value
              
              If Sheets("instructions").Range("C10").Value <> "" Then
                  Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("A" & j + 4).Value = j
               End If
                
            Next j
                    
            'Adding and autofilling  formular to calculate the wieghted score
             
               
                        
            Dim rg As Double
            
            rg = Sheets("instructions").Range("C10").Value + 4
            
            If Sheets("instructions").Range("C10").Value <> "" Then
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("D5").Formula = _
              "= C5 * (instructions!$E$13/100)"
                                        
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("D5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("D5:D" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("F5").Formula = _
              "= E5 * (instructions!$E$14/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("F5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("F5:F" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("H5").Formula = _
              "= G5 * (instructions!$E$15/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("H5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("H5:H" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("J5").Formula = _
              "= I5 * (instructions!$E$16/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("J5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("J5:J" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("L5").Formula = _
              "= K5 * (instructions!$E$17/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("L5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("L5:L" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("N5").Formula = _
              "= M5 * (instructions!$E$18/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("N5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("N5:N" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("P5").Formula = _
              "= O5 * (instructions!$E$19/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("P5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("P5:P" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("R5").Formula = _
              "= Q5 * (instructions!$E$20/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("R5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("R5:R" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("T5").Formula = _
              "= S5 * (instructions!$E$21/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("T5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("T5:T" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("V5").Formula = _
              "= U5 * (instructions!$E$22/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("V5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("V5:V" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("X5").Formula = _
              "= W5 * (instructions!$E$23/100)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("X5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("X5:X" & rg)
        
        'Adding the Overall Socre and ranking for all STS
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Y5").Formula = _
              "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Y5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Y5:Y" & rg)
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Z5").Formula = _
              "=RANK(Y5,Y:Y)"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Z5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("Z5:Z" & rg)
              
              
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("AA5").Formula = _
              "=IF(Y5 >= instructions!$D$25,1,0 )"
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("AA5").AutoFill Destination:= _
              Sheets("Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_2 & "_" & "Day_" & i).Range("AA5:AA" & rg)
            
            
            End If
            
        Next i
    
    End If

Next p
                    




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'3 Updating Overal Sheet to cumulate gradings for multiple grading days''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


' Declaring the loop for all the sheets

Dim rng As Range: Set rng = Application.Range("instructions!C30:C34")

For w = 1 To rng.Rows.Count
    
    Dim names_n As String
    names_n = rng.Cells(RowIndex:=w, ColumnIndex:=1).Value
    Dim f_name_n As String
    f_name = rng.Cells(RowIndex:=1, ColumnIndex:=1).Value
    
    If names_n <> "" Then

    ' Deleating old sheets
    
    Sheets("Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n).Delete

            ' Filling preliminary information
            
            Sheets("overall_template").Copy after:=Sheets("instructions")
            Sheets("overall_template (2)").Visible = True
            
            Sheets("overall_template (2)").Range("B1").Value = Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value
            
            Sheets("overall_template (2)").Range("C2").Value = Sheets("instructions").Range("C3").Value
            Sheets("overall_template (2)").Range("G2").Value = Sheets("instructions").Range("C5").Value
            Sheets("overall_template (2)").Range("J2").Value = Sheets("instructions").Range("C6").Value
            Sheets("overall_template (2)").Range("O2").Value = Sheets("instructions").Range("C26").Value
            Sheets("overall_template (2)").Range("Q2").Value = Sheets("instructions").Range("C27").Value
            
            Sheets("overall_template (2)").Range("A4").Value = "SN"
            
            For i = 1 To Sheets("instructions").Range("C10").Value
              
                If Sheets("instructions").Range("C10").Value <> "" Then
                    Sheets("overall_template (2)").Range("A" & i + 4).Value = i
                 End If
                
            Next i





' storing sheet names for use (locals)
        
        Dim sht1 As String
        
        sht1 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 1
        sht2 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 2
        sht3 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 3
        sht4 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 4
        sht5 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 5
        sht6 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 6
        sht7 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 7
        sht8 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 8
        sht9 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 9
        sht10 = "Matrix_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n & "_" & "Day_" & 10



        'Cumulate all raw scores into overall grading score
        
        'Dim y As Double
        
        y = Sheets("instructions").Range("C28").Value
          
        For k = 5 To Sheets("instructions").Range("C10").Value + 4
                   
              Sheets("overall_template (2)").Range("B" & k).Formula = "=" & sht1 & "!B" & k
                
               
          ' Cummulating results based on the number of sheets
                
            If y = 5 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & "+" & (sht5 & "!C" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & "+" & (sht5 & "!E" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & "+" & (sht5 & "!G" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & "+" & (sht5 & "!I" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & "+" & (sht5 & "!K" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & "+" & (sht5 & "!M" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & "+" & (sht5 & "!O" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & "+" & (sht5 & "!Q" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & "+" & (sht5 & "!S" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & "+" & (sht5 & "!U" & k) & ")" & "/" & 5
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & "+" & (sht5 & "!W" & k) & ")" & "/" & 5
                
                ElseIf y = 4 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & "+" & (sht4 & "!C" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & "+" & (sht4 & "!E" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & "+" & (sht4 & "!G" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & "+" & (sht4 & "!I" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & "+" & (sht4 & "!K" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & "+" & (sht4 & "!M" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & "+" & (sht4 & "!O" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & "+" & (sht4 & "!Q" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & "+" & (sht4 & "!S" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & "+" & (sht4 & "!U" & k) & ")" & "/" & 4
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & "+" & (sht4 & "!W" & k) & ")" & "/" & 4
                
                ElseIf y = 3 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = _
                "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & "+" & (sht3 & "!C" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("E" & k).Formula = _
                "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & "+" & (sht3 & "!E" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("G" & k).Formula = _
                "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & "+" & (sht3 & "!G" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("I" & k).Formula = _
                "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & "+" & (sht3 & "!I" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("K" & k).Formula = _
                "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & "+" & (sht3 & "!K" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("M" & k).Formula = _
                "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & "+" & (sht3 & "!M" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("O" & k).Formula = _
                "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & "+" & (sht3 & "!O" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("Q" & k).Formula = _
                "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & "+" & (sht3 & "!Q" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("S" & k).Formula = _
                "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & "+" & (sht3 & "!S" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("U" & k).Formula = _
                "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & "+" & (sht3 & "!U" & k) & ")" & "/" & 3
                Sheets("overall_template (2)").Range("W" & k).Formula = _
                "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & "+" & (sht3 & "!W" & k) & ")" & "/" & 3
                
                ElseIf y = 2 Then
                
                Sheets("overall_template (2)").Range("C" & k).Formula = "=" & "(" & (sht1 & "!C" & k) & "+" & (sht2 & "!C" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("E" & k).Formula = "=" & "(" & (sht1 & "!E" & k) & "+" & (sht2 & "!E" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("G" & k).Formula = "=" & "(" & (sht1 & "!G" & k) & "+" & (sht2 & "!G" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("I" & k).Formula = "=" & "(" & (sht1 & "!I" & k) & "+" & (sht2 & "!I" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("K" & k).Formula = "=" & "(" & (sht1 & "!K" & k) & "+" & (sht2 & "!K" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("M" & k).Formula = "=" & "(" & (sht1 & "!M" & k) & "+" & (sht2 & "!M" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("O" & k).Formula = "=" & "(" & (sht1 & "!O" & k) & "+" & (sht2 & "!O" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("Q" & k).Formula = "=" & "(" & (sht1 & "!Q" & k) & "+" & (sht2 & "!Q" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("S" & k).Formula = "=" & "(" & (sht1 & "!S" & k) & "+" & (sht2 & "!S" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("U" & k).Formula = "=" & "(" & (sht1 & "!U" & k) & "+" & (sht2 & "!U" & k) & ")" & "/" & 2
                Sheets("overall_template (2)").Range("W" & k).Formula = "=" & "(" & (sht1 & "!W" & k) & "+" & (sht2 & "!W" & k) & ")" & "/" & 2
                
            End If
        
            
        Next k
            
            
            'Adding and autofilling  formular to calculate the wieghted score
             ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
             
           
            Dim rg_1 As Double
            
            
            rg_1 = Sheets("instructions").Range("C10") + 4
            
              Sheets("overall_template (2)").Range("D5").Formula = "= C5 * (instructions!$E$13/100)"
              Sheets("overall_template (2)").Range("D5").AutoFill Destination:=Sheets("overall_template (2)").Range("D5:D" & rg_1)
              
              Sheets("overall_template (2)").Range("F5").Formula = "= E5* (instructions!$E$14/100)"
              Sheets("overall_template (2)").Range("F5").AutoFill Destination:=Sheets("overall_template (2)").Range("F5:F" & rg_1)
              
              Sheets("overall_template (2)").Range("H5").Formula = "= G5* (instructions!$E$15/100)"
              Sheets("overall_template (2)").Range("H5").AutoFill Destination:=Sheets("overall_template (2)").Range("H5:H" & rg_1)
              
              Sheets("overall_template (2)").Range("J5").Formula = "= I5 * (instructions!$E$16/100)"
              Sheets("overall_template (2)").Range("J5").AutoFill Destination:=Sheets("overall_template (2)").Range("J5:J" & rg_1)
              
              Sheets("overall_template (2)").Range("L5").Formula = "= K5  * (instructions!$E$17/100)"
              Sheets("overall_template (2)").Range("L5").AutoFill Destination:=Sheets("overall_template (2)").Range("L5:L" & rg_1)
              
              Sheets("overall_template (2)").Range("N5").Formula = "= M5  * (instructions!$E$18/100)"
              Sheets("overall_template (2)").Range("N5").AutoFill Destination:=Sheets("overall_template (2)").Range("N5:N" & rg_1)
              
              Sheets("overall_template (2)").Range("P5").Formula = "= O5  * (instructions!$E$19/100)"
              Sheets("overall_template (2)").Range("P5").AutoFill Destination:=Sheets("overall_template (2)").Range("P5:P" & rg_1)
              
              Sheets("overall_template (2)").Range("R5").Formula = "= Q5  * (instructions!$E$20/100)"
              Sheets("overall_template (2)").Range("R5").AutoFill Destination:=Sheets("overall_template (2)").Range("R5:R" & rg_1)
              
              Sheets("overall_template (2)").Range("T5").Formula = "= S5  * (instructions!$E$21/100)"
              Sheets("overall_template (2)").Range("T5").AutoFill Destination:=Sheets("overall_template (2)").Range("T5:T" & rg_1)
              
              Sheets("overall_template (2)").Range("V5").Formula = "= U5  * (instructions!$E$22/100)"
              Sheets("overall_template (2)").Range("V5").AutoFill Destination:=Sheets("overall_template (2)").Range("V5:V" & rg_1)
              
              Sheets("overall_template (2)").Range("X5").Formula = "= W5  * (instructions!$E$23/100)"
              Sheets("overall_template (2)").Range("X5").AutoFill Destination:=Sheets("overall_template (2)").Range("X5:X" & rg_1)
        
             'Adding the Overall Socre and ranking for all STS
        
              Sheets("overall_template (2)").Range("Y5").Formula = "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
              Sheets("overall_template (2)").Range("Y5").AutoFill Destination:=Sheets("overall_template (2)").Range("Y5:Y" & rg_1)
              
              Sheets("overall_template (2)").Range("Z5").Formula = "=RANK(Y5,Y:Y)"
              Sheets("overall_template (2)").Range("Z5").AutoFill Destination:=Sheets("overall_template (2)").Range("Z5:Z" & rg_1)
              
            
              Sheets("overall_template (2)").Range("AA5").Formula = "=IF(Y5 >= instructions!$D$25,1,0)"
              Sheets("overall_template (2)").Range("AA5").AutoFill Destination:=Sheets("overall_template (2)").Range("AA5:AA" & rg_1)
                    
        
            ' Add overall sheet names
        Sheets("overall_template (2)").Name = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & names_n


      End If


Next w
       
       
    ' Overall Master
    ' Delete previous

    Sheets("Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & "all").Delete

    ' Create new

    Sheets("overall_template").Copy after:=Sheets("instructions")
    Sheets("overall_template (2)").Visible = True

    Sheets("overall_template (2)").Range("B1").Value = Sheets("instructions").Range("C3").Value & "_" & "Training Evaluation Matrix" & "_" & Sheets("instructions").Range("C9").Value

    Sheets("overall_template (2)").Range("C2").Value = Sheets("instructions").Range("C3").Value
    Sheets("overall_template (2)").Range("G2").Value = Sheets("instructions").Range("C5").Value
    Sheets("overall_template (2)").Range("J2").Value = Sheets("instructions").Range("C6").Value
    Sheets("overall_template (2)").Range("O2").Value = Sheets("instructions").Range("C26").Value
    Sheets("overall_template (2)").Range("Q2").Value = Sheets("instructions").Range("C27").Value

    Sheets("overall_template (2)").Range("A4").Value = "SN"

    For i = 1 To Sheets("instructions").Range("C10").Value

        If Sheets("instructions").Range("C10").Value <> "" Then
            Sheets("overall_template (2)").Range("A" & i + 4).Value = i
         End If

    Next i


     Dim n1 As String
     Dim n2 As String
     Dim n3 As String
     Dim n4 As String
     Dim n5 As String

    Dim n1_c As Integer
    Dim n2_c As Integer
    Dim n3_c As Integer
    Dim n4_c As Integer
    Dim n5_c As Integer



' storing sheet names for use (locals)

     Dim n_st1 As String
     Dim n_sht2 As String
     Dim n_sht3 As String
     Dim n_sht4 As String
     Dim n_sht5 As String


Dim rng_n As Range: Set rng_n = Application.Range("instructions!C30:C34")

    n1 = rng_n.Cells(RowIndex:=1, ColumnIndex:=1).Value
    n2 = rng_n.Cells(RowIndex:=2, ColumnIndex:=1).Value
    n3 = rng_n.Cells(RowIndex:=3, ColumnIndex:=1).Value
    n4 = rng_n.Cells(RowIndex:=4, ColumnIndex:=1).Value
    n5 = rng_n.Cells(RowIndex:=5, ColumnIndex:=1).Value


'Verify if all five cells for assessors have values and set sheet names of their overals as locals

If n1 <> "" Then
     n_sht1 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=1, ColumnIndex:=1).Value
     n1_c = 1 ' if there's an assesor in the sheet
Else:
n_sht1 = "no_inp"
n1_c = 0
End If

 If n2 <> "" Then
     n_sht2 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=2, ColumnIndex:=1).Value
     n2_c = 1
Else:
    n_sht2 = "no_inp"
    n2_c = 0
End If

 If n3 <> "" Then
     n_sht3 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=3, ColumnIndex:=1).Value
     n3_c = 1 ' if there's an assesor in the sheet
Else:
    n_sht3 = "no_inp"
    n3_c = 0
End If

 If n4 <> "" Then
     n_sht4 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=4, ColumnIndex:=1).Value
     n4_c = 1 ' if there's an assesor in the sheet
Else:
    n_sht4 = "no_inp"
    n4_c = 0
End If

 If n5 <> "" Then
     n_sht5 = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & rng_n.Cells(RowIndex:=5, ColumnIndex:=1).Value
     n5_c = 1
Else:
    n_sht5 = "no_inp"
    n5_c = 0
End If


' Declare the variable that counts the total number of Assessors

Dim n_count As Integer
    n_count = n1_c + n2_c + n3_c + n4_c + n5_c


'Cumulate all raw scores into overall grading score

'Dim y As Double

For k = 5 To Sheets("instructions").Range("C10").Value + 4

      Sheets("overall_template (2)").Range("B" & k).Formula = "=" & sht1 & "!B" & k

  ' Adding the overall sheets
        Sheets("overall_template (2)").Range("C" & k).Formula = _
        "=" & "(" & (n_sht1 & "!C" & k) & "+" & (n_sht2 & "!C" & k) & "+" & (n_sht3 & "!C" & k) & "+" & (n_sht4 & "!C" & k) & "+" & (n_sht5 & "!C" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("E" & k).Formula = _
        "=" & "(" & (n_sht1 & "!E" & k) & "+" & (n_sht2 & "!E" & k) & "+" & (n_sht3 & "!E" & k) & "+" & (n_sht4 & "!E" & k) & "+" & (n_sht5 & "!E" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("G" & k).Formula = _
         "=" & "(" & (n_sht1 & "!G" & k) & "+" & (n_sht2 & "!G" & k) & "+" & (n_sht3 & "!G" & k) & "+" & (n_sht4 & "!G" & k) & "+" & (n_sht5 & "!G" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("I" & k).Formula = _
        "=" & "(" & (n_sht1 & "!I" & k) & "+" & (n_sht2 & "!I" & k) & "+" & (n_sht3 & "!I" & k) & "+" & (n_sht4 & "!I" & k) & "+" & (n_sht5 & "!I" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("K" & k).Formula = _
         "=" & "(" & (n_sht1 & "!K" & k) & "+" & (n_sht2 & "!K" & k) & "+" & (n_sht3 & "!K" & k) & "+" & (n_sht4 & "!K" & k) & "+" & (n_sht5 & "!K" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("M" & k).Formula = _
        "=" & "(" & (n_sht1 & "!M" & k) & "+" & (n_sht2 & "!M" & k) & "+" & (n_sht3 & "!M" & k) & "+" & (n_sht4 & "!M" & k) & "+" & (n_sht5 & "!M" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("O" & k).Formula = _
        "=" & "(" & (n_sht1 & "!O" & k) & "+" & (n_sht2 & "!O" & k) & "+" & (n_sht3 & "!O" & k) & "+" & (n_sht4 & "!O" & k) & "+" & (n_sht5 & "!O" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("Q" & k).Formula = _
        "=" & "(" & (n_sht1 & "!Q" & k) & "+" & (n_sht2 & "!Q" & k) & "+" & (n_sht3 & "!Q" & k) & "+" & (n_sht4 & "!Q" & k) & "+" & (n_sht5 & "!Q" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("S" & k).Formula = _
        "=" & "(" & (n_sht1 & "!S" & k) & "+" & (n_sht2 & "!S" & k) & "+" & (n_sht3 & "!S" & k) & "+" & (n_sht4 & "!S" & k) & "+" & (n_sht5 & "!S" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("U" & k).Formula = _
        "=" & "(" & (n_sht1 & "!U" & k) & "+" & (n_sht2 & "!U" & k) & "+" & (n_sht3 & "!U" & k) & "+" & (n_sht4 & "!U" & k) & "+" & (n_sht5 & "!U" & k) & ")" & "/" & n_count
        Sheets("overall_template (2)").Range("W" & k).Formula = _
        "=" & "(" & (n_sht1 & "!W" & k) & "+" & (n_sht2 & "!W" & k) & "+" & (n_sht3 & "!W" & k) & "+" & (n_sht4 & "!W" & k) & "+" & (n_sht5 & "!W" & k) & ")" & "/" & n_count
 Next k


 Dim rg_n As Double


            rg_n = Sheets("instructions").Range("C10") + 4

              Sheets("overall_template (2)").Range("D5").Formula = "= C5 * (instructions!$E$13/100)"
              Sheets("overall_template (2)").Range("D5").AutoFill Destination:=Sheets("overall_template (2)").Range("D5:D" & rg_n)

              Sheets("overall_template (2)").Range("F5").Formula = "= E5* (instructions!$E$14/100)"
              Sheets("overall_template (2)").Range("F5").AutoFill Destination:=Sheets("overall_template (2)").Range("F5:F" & rg_n)

              Sheets("overall_template (2)").Range("H5").Formula = "= G5* (instructions!$E$15/100)"
              Sheets("overall_template (2)").Range("H5").AutoFill Destination:=Sheets("overall_template (2)").Range("H5:H" & rg_n)

              Sheets("overall_template (2)").Range("J5").Formula = "= I5 * (instructions!$E$16/100)"
              Sheets("overall_template (2)").Range("J5").AutoFill Destination:=Sheets("overall_template (2)").Range("J5:J" & rg_n)

              Sheets("overall_template (2)").Range("L5").Formula = "= K5  * (instructions!$E$17/100)"
              Sheets("overall_template (2)").Range("L5").AutoFill Destination:=Sheets("overall_template (2)").Range("L5:L" & rg_n)

              Sheets("overall_template (2)").Range("N5").Formula = "= M5  * (instructions!$E$18/100)"
              Sheets("overall_template (2)").Range("N5").AutoFill Destination:=Sheets("overall_template (2)").Range("N5:N" & rg_n)

              Sheets("overall_template (2)").Range("P5").Formula = "= O5  * (instructions!$E$19/100)"
              Sheets("overall_template (2)").Range("P5").AutoFill Destination:=Sheets("overall_template (2)").Range("P5:P" & rg_n)

              Sheets("overall_template (2)").Range("R5").Formula = "= Q5  * (instructions!$E$20/100)"
              Sheets("overall_template (2)").Range("R5").AutoFill Destination:=Sheets("overall_template (2)").Range("R5:R" & rg_n)

              Sheets("overall_template (2)").Range("T5").Formula = "= S5  * (instructions!$E$21/100)"
              Sheets("overall_template (2)").Range("T5").AutoFill Destination:=Sheets("overall_template (2)").Range("T5:T" & rg_n)

              Sheets("overall_template (2)").Range("V5").Formula = "= U5  * (instructions!$E$22/100)"
              Sheets("overall_template (2)").Range("V5").AutoFill Destination:=Sheets("overall_template (2)").Range("V5:V" & rg_n)

              Sheets("overall_template (2)").Range("X5").Formula = "= W5  * (instructions!$E$23/100)"
              Sheets("overall_template (2)").Range("X5").AutoFill Destination:=Sheets("overall_template (2)").Range("X5:X" & rg_n)

             'Adding the Overall Socre and ranking for all STS

              Sheets("overall_template (2)").Range("Y5").Formula = "=(D5+F5+H5+J5+L5+N5+P5+R5+T5+V5+X5)*10"
              Sheets("overall_template (2)").Range("Y5").AutoFill Destination:=Sheets("overall_template (2)").Range("Y5:Y" & rg_n)

              Sheets("overall_template (2)").Range("Z5").Formula = "=RANK(Y5,Y:Y)"
              Sheets("overall_template (2)").Range("Z5").AutoFill Destination:=Sheets("overall_template (2)").Range("Z5:Z" & rg_n)


              Sheets("overall_template (2)").Range("AA5").Formula = "=IF(Y5 >= instructions!$D$25,1,0)"
              Sheets("overall_template (2)").Range("AA5").AutoFill Destination:=Sheets("overall_template (2)").Range("AA5:AA" & rg_n)




' Naming the Master overall

Sheets("overall_template (2)").Name = "Overall_" & Sheets("instructions").Range("C4").Value & "_" & Sheets("instructions").Range("C9").Value & "_" & "all"

      
       

'Store the number of assaments in code woeksheet
    Sheets("codes").Range("A1").Value = Sheets("instructions").Range("C28").Value
      
    ' MsgBox n_sht5


End Sub

