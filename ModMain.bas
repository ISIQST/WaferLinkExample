Attribute VB_Name = "ModMain"
Option Explicit

Public Tsk As Object ' TSKWafer.Application
Public TskProber As Object 'New TSKWafer.ClsTSKProberUF300
Public QstTester As Object 'New TSKWafer.clsTester
Public TSKWaferDesc As Object 'TSKWafer.clsWaferDesc
Global gNetHostCallBack As Object
Global gFrmMain As New FrmMain
Global qst As Object

Sub Main()
Set Tsk = New TSKWafer.Application
Dim k As Long
Dim blockID As String
Dim currentBlock As Long        'Current block the prober thinks it is at
Dim currentRow As Long          'Current Row
Dim CurrentCol As Long          'Current Column
Dim movesuccess As Boolean
Dim curxloc As Single, curyloc As Single

Set TskProber = Tsk.TskProber
Set QstTester = Tsk.Tester
Set TSKWaferDesc = Tsk.WaferDesc

    Call TskProber.Connect(QstTester.GPIBaddress)      'Will connect to the GPIB address in memory

'''Moving by row and Col numbers as specified on wafer map.  This is if wafer map has been created
'------------------------------------------------------------------
    'Looking at wafermap, the block is counted along rows. so first row in map will be 1,2,3 etc..
    'rows are from top to bottom
    'cols are from left to right
    
    
    
    'Can get information on the current block, row and col the prober thinks its on.
    currentBlock = TskProber.currentBlock
    CurrentCol = TskProber.CurrentDie
    currentRow = TskProber.currentRow
    blockID = TSKWaferDesc.Block(6).BlockName
    
    'If you want to move to any block in parent block. i.e block 6 in this case then call.
    Call TskProber.MoveToBlock(6)

    'Will move to the row and col specified in the current block: In Your case allow you to move to any module
    'if moved correctly will return true to movesucess
    movesuccess = TskProber.MoveToSliderXYinBlockDualChan(5, 5, False)
    
    'if you want to move to a different row and col in a different block all in one move then
    blockID = TSKWaferDesc.Block(8).BlockName   'Get the block id of block you want to move to.
    TskProber.currentBlock = 8  ' Set to block you want to move to: block 8 for example.
    movesuccess = TskProber.MoveToSliderXYinBlockDualChan(5, 5, False)  'will move to new block and selected row and col
'------------------------------------------------------------------



    
'''Moving to any position on the wafer.
'------------------------------------------------------------------
    Call TskProber.GetXYCoordinates(TskProber.UD, curxloc, curyloc) 'Get Current position you are on the wafer
    Call TskProber.GoToAbsolutePosition(1000, 1000)    ' Will move from current position by amounts specified in um
    Call TskProber.GetXYCoordinates(TskProber.UD, curxloc, curyloc) 'Get new position you are on
    'Debug.Print objWaferDesc.NumberofBlocks
'------------------------------------------------------------------

    
End Sub

Function gUserMSGBox(Prompt As String, Optional style As VbMsgBoxStyle = vbOKOnly, Optional title As String = "TSKWafer") As VbMsgBoxResult
   If qst Is Nothing Then
      gUserMSGBox = gUserMSGBox(Prompt, style, title)
   Else
      On Error Resume Next
      gUserMSGBox = qst.UserMsgBox(Prompt, style, title, 0)
      Err.Clear
   End If
End Function

Sub gShowForm(ByRef frPtr As Object)
   If gNetHostCallBack Is Nothing Then
      frPtr.Show
   Else
      Call gNetHostCallBack.ShowForm(frPtr)
   End If
End Sub

Function errorhandler%(FuncName$)
    Dim Stat%
        
    If Err = -1 Then
        errorhandler = vbAbort
        Exit Function
    End If
    If Err = 401 Then 'trying to show non-modal form when modal is displayed
      errorhandler = vbIgnore
      Exit Function
    End If
    
    If Err <> 0 Then
        errorhandler = gUserMSGBox(VBA.CStr(Err) + " [" & FuncName & "] : " + Error, vbAbortRetryIgnore, Err.Source)
    'Else
    '    errorhandler = vbAbort
    End If
    
End Function
