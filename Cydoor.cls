VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCydoor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' This ready made class provide you with an easy interface to
' Cydoors' SDK just add this declration to one of your modules
' "Dim Cydoor As New clsCydoor" - for a class module
' "Global Cydoor As New clsCydoor" - for a normal module
'
' You may by any mean copy this class to your software as is
' and use/manipulate it any way you find suitable to your needs
'

Option Explicit

Private Declare Function pShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hWnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Declare Function pServiceShow& Lib "cd_clint.dll" Alias "ServiceShow" _
  (ByVal AdwrCode As Long, ByVal LoctNum As Long, _
   ByVal LoctIndx As Long, ByVal hWnd As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal LenX As Long, ByVal LenY As Long, _
   ByVal Mode As Long, _
   ByVal General1 As Long, ByVal General2 As Long)
   
' MOST IMPORTANT!! you have to perform ServiceClose function
' in From_Unload event only
Private Declare Function pServiceClose& Lib "cd_clint.dll" Alias "ServiceClose" _
  (ByVal LoctIndx As Long, _
   ByVal hWnd As Long, _
   ByVal General2 As Long)
   
Private Declare Sub pChannelRead Lib "cd_clint.dll" Alias "ChannelRead" _
  (ByVal AdwrCode As Long, _
   ByVal ChannelIn As String, _
   ByVal Resv1 As Long, _
   ByVal Resv2 As Long)

Private Declare Sub pChannelWrite Lib "cd_clint.dll" Alias "ChannelWrite" _
  (ByVal AdwrCode As Long, _
   ByVal ChannelOut As String, _
   ByVal Resv1 As Long, _
   ByVal Resv2 As Long)
   
' -------------------- public section --------------------------------

Public Function ShellExecute(ByVal hWnd As Long, ByVal lpOperation As String, _
   ByVal lpFile As String, ByVal lpParameters As String, _
   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
   
   ShellExecute = pShellExecute(hWnd, lpOperation, lpFile, lpParameters, lpDirectory, nShowCmd)
End Function
Public Function ServiceShow(ByVal AdwrCode As Long, ByVal LoctNum As Long, _
   ByVal LoctIndx As Long, ByVal hWnd As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal LenX As Long, ByVal LenY As Long, _
   ByVal Mode As Long, _
   ByVal General1 As Long, ByVal General2 As Long) As Boolean

   ServiceShow = pServiceShow(AdwrCode, LoctNum, LoctIndx, hWnd, X, Y, LenX, LenY, Mode, General1, General2)
End Function

Public Function ServiceClose(ByVal LoctIndx As Long, ByVal hWnd As Long) As Boolean
   ServiceClose = pServiceClose(LoctIndx, hWnd, 0)
End Function
   
Public Sub ChannelRead(ByVal AdwrCode As Long, _
   ByVal ChannelIn As String)
   
   Call pChannelRead(AdwrCode, ChannelIn, 0, 0)
End Sub

Public Sub ChannelWrite(ByVal AdwrCode As Long, _
   ByVal ChannelOut As String)

   Call pChannelWrite(AdwrCode, ChannelOut, 0, 0)
End Sub
