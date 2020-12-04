Attribute VB_Name = "Main_"
Dim IE As SHDocVw.InternetExplorerMedium ' reference Microsoft Internet Controls
Dim HTMLDoc As MSHTML.HTMLDocument ' reference Microsoft HTML Object Library
Dim HTMlInput As MSHTML.IHTMLElement ' reference Microsoft HTML Object Library
'"C:\Windows\SysWOW64\UIAutomationCore.dll" copy it from location add it to documents folder and add the reference manualy
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
Sub Main()
    
    Application.DisplayAlerts = False
            Application.ScreenUpdating = False
    Application.AskToUpdateLinks = False

Dim StartTime As Double
Dim MinutesElapsed As String
    StartTime = Timer
    
    BExWebLink = ThisWorkbook.Worksheets(1).Cells.Range("B1").Value ' load the BEx Web URL
    Application.StatusBar = "Downloading QC file"
Call biDownload(BExWebLink) ' IE automation
        Application.StatusBar = "Searching for latestet downloaded file"
        Application.Wait (Now + TimeValue("00:00:02"))
Call BExFileOpen ' find donwloaded file and open it
    LatestFile = ThisWorkbook.Worksheets(1).Cells.Range("b4").Value
    Workbooks(LatestFile).Activate
    
    ''''''''''''''''''
    'Data manipulation
    ''''''''''''''''''
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.AskToUpdateLinks = True
    MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
    MsgBox "File generated in " & MinutesElapsed & " minutes", vbInformation
    Application.StatusBar = ""

End Sub

Sub biDownload(ByVal link As String)
    Dim el As Object
Set IE = New InternetExplorerMedium
    IE.Visible = True
    IE.navigate link
    
'loop while IE loads
Do While IE.Busy = True Or IE.readyState <> 4

Loop


'loop until frame is visible
Do
    Set el = Nothing
    On Error Resume Next
    Set el = IE.document.frames("iframe_Roundtrip_9223372036563636042").document.getElementById("DLG_VARIABLE_vsc_cvl_VAR_1_INPUT_inp")
    On Error GoTo 0
    DoEvents
Loop While el Is Nothing

    Sleep 1000
Set frame_main = IE.document.getElementById("iframe_Roundtrip_9223372036563636042") 'read the iframe
    Sleep 1000
Set frame_document = frame_main.contentWindow.document ' get all the fields in the frame
    
    frame_document.getElementById("DLG_VARIABLE_dlgBase_BTNOK").Click ' click on OK button
  
'loop until a specific button on next page is visibile
Do
    Set el = Nothing
    On Error Resume Next
    Set el = IE.document.frames("iframe_Roundtrip_9223372036563636042").document.getElementById("BUTTON_TOOLBAR_2_btn8_acButton")
    On Error GoTo 0
    DoEvents
Loop While el Is Nothing

    Sleep 3000
Set frame_document = frame_main.contentWindow.document ' get the new frame after data's are generated
    frame_document.getElementById("BUTTON_TOOLBAR_2_btn8_acButton").Click 'click export

'loop while IE loads
Do While IE.Busy = True Or IE.readyState <> 4
    
Loop
   '''''''
   'script to press the ok button, provide by the UIAutomationcore.dll
Dim o As IUIAutomation
Dim e As IUIAutomationElement
Set o = New CUIAutomation
Dim h As LongPtr
    h = IE.Hwnd
    h = FindWindowEx(h, 0, "Frame Notification Bar", vbNullString)
    ' loop until Notification bar is visible
    Do While h = 0
        h = IE.Hwnd
        h = FindWindowEx(h, 0, "Frame Notification Bar", vbNullString)
    Loop
Set e = o.ElementFromHandle(ByVal h)
Dim iCnd As IUIAutomationCondition
Dim Button As IUIAutomationElement
Dim InvokePattern As IUIAutomationInvokePattern
    Sleep 1000
Set iCnd = o.CreatePropertyCondition(UIA_NamePropertyId, "Save") 'search for Save button in the Notification Bar
    Sleep 1000
Set Button = e.FindFirst(TreeScope_Subtree, iCnd) ' Read button ID
Set InvokePattern = Button.GetCurrentPattern(UIA_InvokePatternId)
    InvokePattern.Invoke ' Button click
    Sleep 1000
    IE.Quit 'Close IE
Set IE = Nothing

End Sub

Sub BExFileOpen()
Dim MyPath As String
start:
    Dim MyFile As String
    Dim LatestFile As String
    Dim LatestDate As Date
    Dim LMD As Date
    
    'Specify the path to the folder
    MyPath = "C:\Users\" & Environ("UserName") & "\Downloads\"
    
    'Make sure that the path ends in a backslash
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    
    'Get the first Excel file from the folder
    MyFile = Dir(MyPath & "*.xls", vbNormal)
    
    'If no files were found, exit the sub
    If Len(MyFile) = 0 Then
        MsgBox "No files were found...", vbExclamation
        Exit Sub
    End If
    
    'Loop through each Excel file in the folder
    Do While Len(MyFile) > 0
    
        'Assign the date/time of the current file to a variable
        LMD = FileDateTime(MyPath & MyFile)
        
        'If the date/time of the current file is greater than the latest
        'recorded date, assign its filename and date/time to variables
        If LMD > LatestDate Then
            LatestFile = MyFile
            LatestDate = LMD
        End If
        
        'Get the next Excel file from the folder
        MyFile = Dir
       
        
    Loop
    ThisWorkbook.Worksheets(1).Cells.Range("b4").Value = LatestFile
    If Not ThisWorkbook.Worksheets(1).Cells.Range("b4").Value Like "ZANALYSIS*" Then GoTo start: ' in case the file still dident finish downlooading and dose not match the name pattern it will loop over again.
    
    'Open the latest file
    Workbooks.Open MyPath & LatestFile
       
End Sub



