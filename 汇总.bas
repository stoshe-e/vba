Attribute VB_Name = "模块2"
Sub 复制()
Attribute 复制.VB_ProcData.VB_Invoke_Func = "w\n14"
'
    Dim i As Integer
        
    For i = 1 To 5
   
     
    Sheets(i & "月").Select
    Range("A5:C22").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A" & (i - 1) * 90 + 1).Select
    ActiveSheet.Paste
    
    Sheets(i & "月").Select
    Range("F5:H22").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A" & (i - 1) * 90 + 19).Select
    ActiveSheet.Paste
    
    Sheets(i & "月").Select
    Range("A31:C48").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A" & (i - 1) * 90 + 37).Select
    ActiveSheet.Paste

    Sheets(i & "月").Select
    Range("F31:H48").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A" & (i - 1) * 90 + 55).Select
    ActiveSheet.Paste
  
    Sheets(i & "月").Select
    Range("A57:c74").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A" & (i - 1) * 90 + 73).Select
    ActiveSheet.Paste
    
    Next i
                                                                                
   Range(Range("A" & i), Range("A10")).Select
   Selection.Formula = "1月"
    
    Application.SendKeys ("^{DOWN}")
    Application.SendKeys ("{DOWN}")                                                                 
    
End Sub
