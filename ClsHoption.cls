VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsHoption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'clsHOption class is responsible for interaction with the application, that is going to control it
Option Explicit

Private Notifier As New Collection                  'all the objects in this collection should be notified about user feedback and others - whatever this application deems necessary
Private Const ErrModId As String = "clsHOption: "

Private Enum HDWOptStat
   Missing = 1          'missing hardware or software. objhandle may still be set
   DeadUnrec = 2        'dead due to unrecoverable failure. If found something missing during initialization for example
   Suspended = 4        'recoverable failure, no calls to that module until recovery is turned on (all the tests and higher level calls must be aborted)
   SkipCheckHealth = 8  'debugging mode, reduces communication to the control to bare minimum
   AppBusy = &H10    'hardware is missing but simulation is enabled
End Enum

Public Property Set NetHostCallBack(ByRef obj As Object)
   Set gNetHostCallBack = obj
End Property

Public Sub AddNotifier(ByRef n As Object)
Dim i%
   For i = 1 To Notifier.Count
      If Notifier(i) Is n Then Exit Sub
   Next i
   Notifier.Add n
End Sub

Public Sub RemoveNotifier(ByRef n As Object)
Dim i%
   For i = 1 To Notifier.Count
      If Notifier(i) Is n Then
         Notifier.Remove i
         Exit Sub
      End If
   Next i
End Sub

Public Sub GetNewProperties(ByRef colobjects As Object, ByRef propDetails() As String)
   'propdetails has 1+colobjects.count number of rows and 4 columns.
   On Error Resume Next    'some devices will not support the follwing method
'   Call curHDW.GetNewProperties(colobjects, propDetails)
   Err.Clear
End Sub

Sub UpdateUserFeedback(lastUserFeeback As String, prg!)
   On Error GoTo errorhandler
   Dim s As Object
   For Each s In Notifier
      s.Message = lastUserFeeback
      s.Progress = prg
   Next
errorhandler:
   Select Case errorhandler(ErrModId & "UpdateUserFeedback")
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
End Sub
    
Public Property Get EventInterests() As Integer
    EventInterests = 0 'eEventInt.eCheckHealth Or eEventInt.eRecover
End Property

Public Sub Initialize(Optional strInstance As String = "*")
   Call Initialize2(strInstance, CreateObject("Quasi97.application"))
End Sub

Public Sub Initialize3(strInstance$, ByRef qs As Object)
'''    Call HardwareInitialize
On Error GoTo errorhandler
Dim TempCol As Object
Dim k As Integer
   Call UpdateUserFeedback("Connecting to QST", 0)
   Set qst = qs
'   Set qst = CreateObject("Quasi97.application")
'   Set curHDW = New ClsExtHardware
'   Set objDevice = New ClsGPIBInterface
'   Call objDevice.InitDevice(strInstance)
'   Call SetAnalyzerToTestDefault
   Set Tsk = Nothing
   
   If qst.hoptionmanager.GetPointerByFunction("MechDriver") Is Nothing Then
      If gUserMSGBox("TSK Link module was not found in the hardware configuration table - would you like to add it?", vbQuestion Or vbYesNoCancel) = vbYes Then
         Call qst.hoptionmanager.AddModule("TSKWafer.Driver", "*", "MechDriver", "")
         Set Tsk = qst.hoptionmanager.GetPointerByFunction("MechDriver")
      End If
   Else
      Set Tsk = qst.hoptionmanager.GetPointerByFunction("MechDriver")
   End If
   
   
   If Not Tsk Is Nothing Then
      If Tsk.classname <> "TSKWafer.Driver" Then
         Set TempCol = qst.hoptionmanager.GetAddInsbyFunction("MechDriver")
         For k = 1 To TempCol.Count
            If TempCol.Item(k).classname = "TSKWafer.Driver" Then
               If Not qst.hoptionmanager.GetPointerByName("TSKWafer.Driver", TempCol.Item(k).instancename) Is Nothing Then
                  Set Tsk = qst.hoptionmanager.GetPointerByName("TSKWafer.Driver", TempCol.Item(k).instancename)
               End If
            End If
         Next k
      End If
   End If
   

   Set TskProber = Tsk.gethandle.TSKCls.TskProber
   Set QstTester = Tsk.gethandle.TSKCls.Tester
   Set TSKWaferDesc = Tsk.gethandle.TSKCls.WaferDesc
   
errorhandler:
   If Err = 457 Then Err.Clear: Resume Next
   Select Case errorhandler("Driver: Initialize")
       Case vbRetry
           Resume
       Case vbIgnore
           Resume Next
   End Select
End Sub

Public Function Status() As Integer
'   If llDriver Is Nothing Then Status = Status Or HDWOptStat.Missing
'   If Not objCont.Present Then Status = Status Or HDWOptStat.DeadUnrec
'   If GSuspendAll Then Status = HDWOptStat.Suspended
End Function

Public Sub DetectAllNew(ByRef DevAr() As String)
'the items in the collection must be easily identifiable for "Initialize" function later on
Dim i%
Dim colDevs As New Collection
On Error GoTo errorhandler
        
   'read serial number from the board and add it to the collection
   colDevs.Add "Main"
   ReDim DevAr(colDevs.Count)
   For i = 1 To colDevs.Count
      DevAr(i) = colDevs(i)
   Next i

errorhandler:
   Select Case errorhandler(ErrModId & "DetectAllNew")
      Case vbRetry
         Resume
      Case vbIgnore
         Resume Next
   End Select
End Sub

Public Sub NotifyOptionsUpdated()
   'here need to check if user disabled lldriver, and remove all ties to it
   
End Sub

Function CheckHealth(ByRef usrdescr As String, PartLoadedState As Byte) As Integer
'the module will use this call as an opportunity to check itself
'if fails this would return non-zero value and change the status. partloadedstate non-zero indicates that the head is probing (perhaps some checks can be done only while not probing the head)
   CheckHealth = 0
End Function

Function Recover(ByRef usrdescr$) As Integer
   'If GSuspendAll Then Recover = objMotion.RecoverFromDriverFault
End Function

Sub ShowUserMenu()
   Call gShowForm(gfrmUSer)
   gfrmUSer.ZOrder
End Sub

Sub ShowDiagnostics()
'shows technician menu
   Call gShowForm(gFrmDiag)
   gFrmDiag.ZOrder
End Sub

Public Sub Terminate()
   Unload gFrmMain
'   objDevice.CloseDevice
   Set qst = Nothing
End Sub

''returns pointer to the main hardware, to be used by fmrapp, ria, sttpulser etc
'Public Function Hardware() As ClsExtHardware
'   Set Hardware = curHDW
'End Function

Private Sub Class_Terminate()
   If Not qst Is Nothing Then Call Terminate
End Sub

