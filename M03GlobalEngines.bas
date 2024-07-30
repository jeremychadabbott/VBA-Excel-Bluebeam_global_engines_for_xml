Attribute VB_Name = "M03GlobalEngines"
'Option Explicit
Declare PtrSafe Function ShowWindow _
    Lib "User32" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) _
As Long '
Declare PtrSafe Function SetForegroundWindow _
    Lib "User32" ( _
    ByVal hwnd As Long) _
As Long
Public Const SW_MAXIMIZE As Long = 3&        'Show window Maximised
'Declare mouse events
Public Declare PtrSafe Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare PtrSafe Sub mouse_event Lib "User32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Public Const MOUSEEVENTF_RIGHTUP As Long = &H10
Declare PtrSafe Function GetSystemMetrics32 Lib "User32" _
Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Declare Function Beep Lib "kernel32" _
 (ByVal dwFreq As Long, _
 ByVal dwDuration As Long) As Long

Sub Modle3()

    'shelved for later use
    
End Sub
' Mouse Movement Library
' Double Click
Sub DoubleClick()
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
    Application.Wait (Now + TimeValue("00:00:01")) / 4
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    Application.Wait (Now + TimeValue("00:00:01")) / 4
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
    Application.Wait (Now + TimeValue("00:00:01")) / 4
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
End Sub

' Mouse Movement Library
' TASK BAR
' Click Sage
Sub ClickOnSage()
    Application.Wait (Now + TimeValue("00:00:01"))
    SetCursorPos 1240, 63
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:03"))
    'hit escape 6 times to make sure we're in the right "ready state"
    Application.SendKeys "{Esc}", True
    Application.Wait (Now + TimeValue("00:00:01")) 'must be in format 00:00:00
    Application.SendKeys "{Esc}", True
    Application.Wait (Now + TimeValue("00:00:01")) 'must be in format 00:00:00
    Application.SendKeys "{Esc}", True
    Application.Wait (Now + TimeValue("00:00:01")) 'must be in format 00:00:00
    Application.SendKeys "{Esc}", True
    For Repeat = 1 To 200
        Sleep 50
    Next Repeat
    Application.Wait (Now + TimeValue("00:00:01")) 'must be in format 00:00:00
End Sub

' Mouse Movement Library
' SAGE HOME SCREEN
' Click 6-Project Management
Sub Sage_6_ProjectManagement()
    Application.Wait (Now + TimeValue("00:00:01"))
    SetCursorPos 22, 272
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
End Sub

' Mouse Movement Library
' SAGE HOME SCREEN
' Click 6_2Budgets
Sub Sage_6_2Budgets()
    Application.Wait (Now + TimeValue("00:00:01"))
    SetCursorPos 76, 315
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.SendKeys "~", True
    'Application.Wait (Now + TimeValue("00:00:03"))
    'Application.SendKeys "%{ }", True
    'Application.Wait (Now + TimeValue("00:00:01"))
    'Application.SendKeys "x", True
    'Application.Wait (Now + TimeValue("00:00:01"))
End Sub

' Mouse Movement Library
' SAGE REPORT
' Click Send Report to Excel
Sub Click_send_report_data_to_excel() 'Assumes Sage Maximized
    W = GetSystemMetrics(SM_CXSCREEN) 'width in points
    h = GetSystemMetrics(SM_CYSCREEN) 'height in points
    Application.Wait (Now + TimeValue("00:00:02"))
    SetCursorPos 1149, 61
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Sleep 150
    Call Mouse_left_button_Letgo
End Sub

' CHROME Webpage
' Load Webpage
Sub Load_Chrome_Page(webpage)
    Application.Wait (Now + TimeValue("00:00:01"))
    SetCursorPos 1072, 50 'set mouse into chrome navigation field
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.SendKeys "{del}", True
    For t = 1 To 50
        Application.SendKeys "{BS}"
    Next t
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys webpage, True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Call verify_Chrome_Page(webpage)
End Sub

' CHROME Webpage
' Verify Landing Page
Sub verify_Chrome_Page(webpage)
    SetCursorPos 1072, 50 'set mouse into chrome navigation field
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "^c", True
    Application.Wait (Now + TimeValue("00:00:01"))
    ThisWorkbook.Sheets("Temp").Paste Destination:=Sheets("Temp").Range("A1")
    test = ThisWorkbook.Sheets("Temp").Range("A1")
    If webpage <> test Then
        MsgBox "Failed to land on target web page while navigating chrome!"
    Else
        Application.SendKeys "{Esc}", True
        Application.Wait (Now + TimeValue("00:00:01"))
        'do this
    End If
End Sub

' Mouse Movement Library
' eBuilder OPEN Chrome, Login
' Click Send Report to Excel
Sub Open_Chrome_eBuilder_Login() 'Assumes Sage Maximized
    chromePath = """C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"""
    Shell (chromePath & " -url https://app.e-builder.net/auth/www/index.aspx?ReturnUrl=/index.aspx")
    Application.Wait (Now + TimeValue("00:00:04"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))
    webpage = "https://app.e-builder.net/da2/Home/index.aspx"
    'check if landing page correct
    Call verify_Chrome_Page(webpage)
End Sub

' Mouse Movement Library
' MOUSE CLICK
' Left Button Down
Sub Mouse_left_button_press()
    Sleep (250) 'must be in format 00:00:00
    Call mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0) 'click left mouse
End Sub

' Mouse Movement Library
' MOUSE CLICK
' Left Button UP
Sub Mouse_left_button_Letgo()
    Sleep (250) 'must be in format 00:00:00
    Call mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0) 'release left mouse
End Sub

' Mouse Movement Library
' MOUSE MOVE
' Set Cursor Position
Sub mouse_position(x, y)
    Dim Mx, My As Integer
    SetCursorPos (x), (y)
    Application.Wait (Now + TimeValue("00:00:01")) 'must be in format 00:00:00
End Sub
' Excel workbook actions
' SAVE Open File as Temp
Sub Save_Open_Excel_File_as_Temp(fname, fpath)

    ' Assumptions
    ' Resolution at 1280x768
    ' Excel workbook is open and in full screen

    ' Wait before starting operations
    Application.Wait (Now + TimeValue("00:00:02"))

    ' Maximize Excel window
    Application.SendKeys "%{ }"  ' Open window menu
    Application.Wait (Now + TimeValue("00:00:02"))
    Application.SendKeys "x"     ' Maximize window
    Application.Wait (Now + TimeValue("00:00:02"))

    ' Navigate to "Save As" dialog
    SetCursorPos 27, 48           ' Excel, Click "File"
    Sleep 250
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo

    SetCursorPos 67, 408           ' Excel, Click "Save As"
    Sleep 250
    Call Mouse_left_button_press
    Sleep 250
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))

    SetCursorPos 310, 587          ' Excel, Click "Browse"
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Sleep 250
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:04"))

    ' Maximize "Save As" dialog
    Application.SendKeys "%{ }"
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "x"
    Application.Wait (Now + TimeValue("00:00:01"))

    ' Enter file path in "Save As" dialog
    SetCursorPos 543, 51
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))

    fpath = ThisWorkbook.path
    fname = "temp"

    Application.SendKeys fpath, True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:01"))

    ' Enter file name in "Save As" dialog
    SetCursorPos 543, 605
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys fname, True
    Application.Wait (Now + TimeValue("00:00:01"))

    ' Save the file
    SetCursorPos 1045, 735
    Application.Wait (Now + TimeValue("00:00:01"))
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    Application.Wait (Now + TimeValue("00:00:01"))

    ' Handle overwrite prompt if needed
    SetCursorPos 1045, 735
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "{Tab}", True
    Application.Wait (Now + TimeValue("00:00:01"))
    Application.SendKeys "~", True
    Application.Wait (Now + TimeValue("00:00:04"))

    ' Close Excel
    Application.SendKeys "^{F4}"
    Application.Wait (Now + TimeValue("00:00:03"))

    ' Finalize file path and name
    fpath = fpath & "\temp.xlsx"
    fname = "Temp.xlsx"

End Sub

Sub FormatTempSheet()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets("Temp")

'clear temp sheet of any existing images or merged cells
ws.Range("BA2:CA300").UnMerge
ws.Range("BA2:CA300").ClearContents



'Clear out worksheet
ws.UsedRange.Clear
    
'Write macro type headers to sheet
    ws.Range("A1") = "DECO Order#"
    ws.Range("B1") = "DECO Order Date"
    ws.Range("C1") = "DECO Vendor"
    ws.Range("D1") = "DECO Description"
    ws.Range("E1") = "DECO Job"
    ws.Range("F1") = "DECO Phase"
    ws.Range("G1") = "DECO Ordered by"
    ws.Range("H1") = "VNDR Invoice#"
    ws.Range("I1") = "VNDR ProcessDate"
    ws.Range("J1") = "VNDR InvoiceDate"
    ws.Range("K1") = "VNDR PaidStatus"
    ws.Range("L1") = "VNDR Ttl Disc"
    ws.Range("M1") = "VNDR Freight"
    ws.Range("N1") = "VNDR TotalInvoice"
    ws.Range("O1") = "VNDR Item"
' Write line item type headers to sheet
    ws.Range("P1") = "Item Description"
    ws.Range("Q1") = "Unit"
    ws.Range("R1") = "Quantity"
    ws.Range("S1") = "Price"
    ws.Range("T1") = "Total"
    ws.Range("U1") = "Shipped"
    ws.Range("V1") = "Current"
    ws.Range("W1") = "Cancelled"
    ws.Range("X1") = "CostCode"
    ws.Range("Y1") = "Cost Type"
    ws.Range("Z1") = "Account"
    ws.Range("AA1") = "Doc Type"
'Write error tracking or note headers to sheet
    ws.Range("AB1") = "PDF Exist"
    ws.Range("AC1") = "Entered in Sage"
    ws.Range("AD1") = "Due Date"
    ws.Range("AE1") = "Discount Date"
    ws.Range("AF1") = "Discount"
    ws.Range("AG1") = "Back Ordr"
    ws.Range("AH1") = "Tax"
End Sub

Sub CheckPONumber(TargetPO, Found)
' Found = 0 TargetPO does not conform, refuse to process
' Found = 1 TargetPO is OK to Process
' Found = 2 TargetPO is Subcontract, Send PDF to Fax File
' Found = 3 TargetPO is SHOP, Send PDF to Fax File

' Special Adjustments
    TargetPO = Replace(TargetPO, "2457", "EC2457") '24 05 13
    TargetPO = Replace(TargetPO, "2109DH05292024", "2109-DH-05292024") '24 06 03
    
' Check modifications sheet for adjustments logged previously
    Dim modSheet As Worksheet
    Set modSheet = ThisWorkbook.Sheets("PO Modifications Log")
    
    Dim lastRow As Long
    lastRow = modSheet.Cells(modSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        If modSheet.Cells(i, "A").Value = TargetPO Then
            TargetPO = modSheet.Cells(i, "B").Value
            Exit For
        End If
    Next i
 
' Assign Vendor
Vendor = ThisWorkbook.Sheets("Temp").Range("C2")

'Subcontracts cross reference (Do not enter invoices which are against subcontracts)
    For x = 0 To 5000
        If TargetPO = " " & ThisWorkbook.Sheets("Subcontract list").Range("B1").Offset(x, 0) & " " Or _
        TargetPO = ThisWorkbook.Sheets("Subcontract list").Range("B1").Offset(x, 0) & " " Or _
        TargetPO = " " & ThisWorkbook.Sheets("Subcontract list").Range("B1").Offset(x, 0) Or _
        TargetPO = ThisWorkbook.Sheets("Subcontract list").Range("B1").Offset(x, 0) Or _
            UCase(Right(TargetPO, 2)) = "SC" Then
            'UCase(TargetPO) Like "*-70-*" Or _
            'UCase(TargetPO) Like "*-80-*" Then
                'MsgBox "Identified PO that is a contract number->" & TargetPO
                Found = 2 'Target PO matches a subcontract number
            Exit Sub
        End If
    Next x

'SHOP Invoices
    If UCase(TargetPO) Like "*SHOP*" Then
        'MsgBox "Found SHOP in PO number, send to Fax file..." & TargetPO
        Found = 3 'Send to Fax File
        Exit Sub
    End If

' Set Found variable to 0 if no "issues" are found
    Found = 0

' PO Conformity Check

' Everett Clinic
    If UCase(TargetPO) Like "EC[0-9][0-9][0-9][0-9]*-[A-Z][A-Z]*-[0-9]*" Then Found = 1

' Everett Clinic Thermal Imaging
    If UCase(TargetPO) Like "ECTI[0-9][0-9]*-[A-Z][A-Z]*-[0-9]*" Then Found = 1
    
' Big Job Standard
    If UCase(TargetPO) Like "[0-9][0-9][0-9][0-9]-[A-Z]*-[0-9][0-9][0-9][0-9]*" Then Found = 1
    
'Big Job Change Orders (middle digits assign change order number)
    If UCase(TargetPO) Like "[0-9][0-9][0-9][0-9]-C*[0-9]*-[0-9][0-9][0-9][0-9]*" Then Found = 1
   
'King County 2304 special cases
    If UCase(TargetPO) Like "2304*[0-9][0-9]-*-[0-9][0-9][0-9][0-9]*" Then
        ' Remove the 5th character if it is a hyphen
        If Mid(TargetPO, 5, 1) = "-" Then
            TargetPO = Left(TargetPO, 4) & Mid(TargetPO, 6)
        End If
        Found = 1
        'MsgBox "Watch out, here is a 2304XX invoice" & Chr(13) & TargetPO
    End If
    
'Big Job assign cost code (middle 2 or 3 digits are cost code)
    If UCase(TargetPO) Like "[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then Found = 1
    If UCase(TargetPO) Like "[0-9][0-9][0-9][0-9]-[0-9][0-9][0-9]-[0-9][0-9][0-9][0-9]*" Then Found = 1
    
'Small Job Standard
    If UCase(TargetPO) Like "[0-9][0-9][A-Z][A-Z]-[A-Z][A-Z]*-[0-9][0-9][0-9][0-9]*" Then Found = 1
'If UCase(TargetPO) Like "[0-9][0-9][A-Z][A-Z]-[A-Z][A-Z]*-[0-9][0-9][0-9][0-9][0-9][0-9]" Then Found = 1
    
If Found = 0 Then
    Dim userInput As String
    userInput = InputBox("Processing " & Vendor & " invoice/PO and found non-conforming PO:" _
        & Chr(13) & TargetPO & Chr(13) & "Enter a PO to change it to, or type 'No' to exit and send the PDF to fax file.", "Non-Conforming PO")
    
    If LCase(userInput) = "no" Then
        ' User chose to exit
        Exit Sub
    ElseIf userInput <> "" Then
        ' User entered a new PO
        Set modSheet = ThisWorkbook.Sheets("PO Modifications Log")
        lastRow = modSheet.Cells(modSheet.Rows.Count, "A").End(xlUp).Row + 1
        
        modSheet.Cells(lastRow, "A").Value = TargetPO
        modSheet.Cells(lastRow, "B").Value = userInput
        
        TargetPO = userInput
        
        Found = 1
    Else
        ' User clicked Cancel or left the input blank
        MsgBox "No change was made to PO:" & TargetPO
    End If
End If

' Modifications temp sheet if PO was changed by this modules code
' Determine new Job Number
    Job = ""

' Filter job number from PO
    For Repeat = 1 To 8
        If Mid(TargetPO, Repeat, 1) = "-" Then Exit For
        Job = Job & Mid(TargetPO, Repeat, 1)
    Next Repeat
    ThisWorkbook.Sheets("Temp").Range("E2").Offset(x, 0) = Job
        
' If PO was changed, Write new PO to temp sheet (if needed)
    For x = 0 To 200
        If ThisWorkbook.Sheets("Temp").Range("A2").Offset(x, 0) <> TargetPO And _
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0) <> "" Or _
            ThisWorkbook.Sheets("Temp").Range("E2").Offset(x, 0) = "" And _
            ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0) <> "" Then
            ThisWorkbook.Sheets("Temp").Range("A2").Offset(x, 0) = TargetPO
            ThisWorkbook.Sheets("Temp").Range("E2").Offset(x, 0) = Job
        End If
    Next x

    
End Sub

Sub LocateDirectory(GeneralPath, TargetPath, fname, fpath)

TargetPath = GeneralPath
                '--------------------------------------------------------------Locate Job Subfolder
                Found = 0 'found indicator
                Set objFSO = CreateObject("Scripting.FileSystemObject")
                Set objFolder = objFSO.GetFolder(TargetPath)
                For Each objSubFolder In objFolder.subfolders
                    fname = objSubFolder.Name
                    fpath = objSubFolder.path
                    'MsgBox "Primary Folder " & JobName
                    If jobname = TargetPath Then
                        Found = 1
                        Exit For
                    End If
                        'Else go to Next Tier
                            Set aobjFSO = CreateObject("Scripting.FileSystemObject")
                            Set aobjFolder = aobjFSO.GetFolder(fpath)
                            For Each aobjSubFolder In aobjFolder.subfolders
                                fname = aobjSubFolder.Name
                                fpath = aobjSubFolder.path
                                'MsgBox "Secondary Folder " & fname
                                If Left(fname, Len(TargetJob)) = TargetJob Then
                                    Found = 1
                                    Exit For
                                Else: End If
                                'Else go to Next Tier
                                            Set bobjFSO = CreateObject("Scripting.FileSystemObject")
                                            Set bobjFolder = bobjFSO.GetFolder(fpath)
                                            For Each bobjSubFolder In bobjFolder.subfolders
                                                fname = bobjSubFolder.Name
                                                fpath = bobjSubFolder.path
                                                'MsgBox "Tertiary Folder " & fname
                                                If Left(fname, Len(TargetJob)) = TargetJob Then
                                                    Found = 1
                                                    Exit For
                                                End If
                                            Next bobjSubFolder 'Exit Secondary folder tier
                                        'Exit next Tier Loop
                                If Left(fname, Len(TargetJob)) = TargetJob Then
                                    Found = 1
                                    Exit For
                                End If
                                Next aobjSubFolder 'Exit Secondary folder tier
                    'Exit next Tier Loop
                    If Left(fname, Len(TargetJob)) = TargetJob Then
                        Found = 1
                        Exit For
                    End If 'If Left(fname, Len(TargetJob)) = TargetJob Then
                Next objSubFolder 'Exit primary folder tier
                '------------------------------------------------------------------finish locate folder
                If Found = 1 Then 'COntinue on to find timecard workbook
                'filter fpath
                'For x = 0 To 100
                '   If Right(fpath, 1) <> "\" Then fpath = Left(fpath, Len(fpath) - 1)
                'Next x
Next xoffset
If Found = 0 Then MsgBox "Failed to locate target directory"
End Sub

Sub Move_data_to_Sage_Xfer_Sheet()

'clear existing data
ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A2:AA500").ClearContents


For x = 0 To 500
    'check if there is more to move
        If ThisWorkbook.Sheets("Temp").Range("P1").Offset(x, 0) = "" Then Exit For
        
        ' transpose Total for later use
        For y = 0 To 35
            ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A1").Offset(x, y) = ThisWorkbook.Sheets("Temp").Range("A1").Offset(x, y)
        Next y

    
    'Move Item Description
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0)
    'Move Unit
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("C2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("Q2").Offset(x, 0)
    'Move Qty
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("D2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("R2").Offset(x, 0)
    'Move Price
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("E2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("S2").Offset(x, 0)
    'Move Total
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("F2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("T2").Offset(x, 0)
    'Move Cost Code
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("J2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("X2").Offset(x, 0)
    'Move Cost Type
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("K2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("Y2").Offset(x, 0)
    'Move Account
        'ThisWorkbook.Sheets("Sage Xfer Sheet").Range("L2").Offset(x, 0) = ThisWorkbook.Sheets("Temp").Range("Z2").Offset(x, 0)
Next x

End Sub
Sub Move_data_to_Sage_Temp_Sheet()

    ' Clear existing data in the Temp sheet
    ThisWorkbook.Sheets("Temp").Range("A1:Z500").ClearContents
    
    For x = 0 To 500
        ' Check if there is more data to move
        If ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A1").Offset(x, 0) = "" Then Exit For
        
        ' Move Vendor Invoice Number
        ' transpose header data
        For y = 0 To 35
            ThisWorkbook.Sheets("Temp").Range("A1").Offset(x, y) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A1").Offset(x, y)
        Next y
        
        ' Move Item Description
        'ThisWorkbook.Sheets("Temp").Range("P2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("A2").Offset(x, 0)
        
        ' Move Unit
        'ThisWorkbook.Sheets("Temp").Range("Q2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("C2").Offset(x, 0)
        
        ' Move Qty
        'ThisWorkbook.Sheets("Temp").Range("R2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("D2").Offset(x, 0)
        
        ' Move Price
        'ThisWorkbook.Sheets("Temp").Range("S2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("E2").Offset(x, 0)
        
        ' Move Line Total
        'ThisWorkbook.Sheets("Temp").Range("T2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("F2").Offset(x, 0)
        
        ' Add invoice total cell N2
        'ThisWorkbook.Sheets("Temp").Range("N2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("P2")
        
        ' Move Cost Code
        'ThisWorkbook.Sheets("Temp").Range("X2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("J2").Offset(x, 0)
        
        ' Move Cost Type
        'ThisWorkbook.Sheets("Temp").Range("Y2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("K2").Offset(x, 0)
        
        ' Move Account
        'ThisWorkbook.Sheets("Temp").Range("Z2").Offset(x, 0) = ThisWorkbook.Sheets("Sage Xfer Sheet").Range("L2").Offset(x, 0)
    Next x
End Sub

Sub CleanString(StringToClean)

    CleanedString = StringToClean
    CleanedString = Replace(StringToClean, "™", "")
    CleanedString = Replace(CleanedString, "#", "")
    CleanedString = Replace(CleanedString, "/", "")
    CleanedString = Replace(CleanedString, ",", "")
    CleanedString = Replace(CleanedString, "'", "")
    CleanedString = Replace(CleanedString, "-", "")
    CleanedString = Replace(CleanedString, " ", "")
    StringToClean = CleanedString

End Sub

Sub check_for_Sage_Errors(ErrorCode)
    'Click on Sage header
    SetCursorPos 613, 33
    Application.Wait (Now + TimeValue("00:00:01")) / 2
    Call Mouse_left_button_press
    Call Mouse_left_button_Letgo
    
    'Create a DataObject to access the clipboard
    Dim clipData As New DataObject
    clipData.GetFromClipboard
    
    'Copy the content
    Application.SendKeys "^a"
    Application.Wait (Now + TimeValue("00:00:01")) / 2
    Application.SendKeys "^c"
    Application.Wait (Now + TimeValue("00:00:01"))
    
    'Assign the clipboard text to the variable SageScreen
    Dim SageScreen As String
    SageScreen = clipData.GetText
    
    'Analyze clipboard
    If InStr(SageScreen, "Error") > 0 Then
        'Message to user "Program halted, detected Sage error"
        MsgBox "Program halted, detected Sage error", vbCritical, "Sage Error"
        End
    End If
End Sub
