VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Rhino COM Tester"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2655
   Icon            =   "TestVBApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Platform"
      Height          =   855
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   2175
      Begin VB.OptionButton rdoWin64 
         Caption         =   "64-bit"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton rdoWin32 
         Caption         =   "32-Bit"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rhino Version"
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton rdoRhino5 
         Caption         =   "Rhino 5.0"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton rdoRhino4 
         Caption         =   "Rhino 4.0"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnOpen 
      Caption         =   "Open..."
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CheckBox chkRelease 
      Caption         =   "Release without closing "
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CommandButton btnUnload 
      Caption         =   "Unload Rhino"
      Height          =   372
      Left            =   240
      TabIndex        =   5
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton btnLoad 
      Caption         =   "Load Rhino"
      Height          =   372
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Object Type"
      Height          =   852
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   2175
      Begin VB.OptionButton rdoInterface 
         Caption         =   "Interface"
         Height          =   252
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton rdoApplication 
         Caption         =   "Application"
         Height          =   252
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton btnHide 
      Caption         =   "Hide Rhino"
      Height          =   372
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton btnShow 
      Caption         =   "Show Rhino"
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' Form data members
Dim m_objRhino As Object
Dim m_bRhino5 As Boolean

Private Sub Form_Load()
  Set m_objRhino = Nothing
  m_bRhino5 = False
  rdoRhino5 = True
  rdoWin32 = True
  rdoApplication = True
  chkRelease = 0
  chkRelease.Enabled = True
End Sub

Private Sub rdoRhino4_Click()
  chkRelease.Enabled = False
End Sub

Private Sub rdoRhino5_Click()
  chkRelease.Enabled = True
End Sub

Private Sub btnLoad_Click()
  Dim strObject As String
  If (rdoRhino4 = True) Then
    m_bRhino5 = False
    If (rdoApplication = True) Then
      strObject = "Rhino4.Application"
    Else
      strObject = "Rhino4.Interface"
    End If
  Else
    m_bRhino5 = True
    If (rdoWin32 = True) Then
      If (rdoApplication = True) Then
        strObject = "Rhino5.Application"
      Else
        strObject = "Rhino5.Interface"
      End If
    Else
      If (rdoApplication = True) Then
        strObject = "Rhino5x64.Application"
      Else
        strObject = "Rhino5x64.Interface"
      End If
    End If
  End If
  
  If (m_objRhino Is Nothing) Then
    On Error Resume Next
    Set m_objRhino = CreateObject(strObject)
    If (Err.Number <> 0) Then
      Call MsgBox(Err.Description, vbCritical + vbOKOnly, "Rhino COM Tester")
      Exit Sub
    End If
    If (m_objRhino Is Nothing) Then
      Call MsgBox("Unable to create " & strObject & " object.", vbCritical + vbOKOnly, "Rhino COM Tester")
      Exit Sub
    End If
  End If
 
End Sub

Private Sub btnShow_Click()
  Dim hr As Long: hr = 0
  If Not (m_objRhino Is Nothing) Then
    ' Enables the COM server process called to take focus away from the client application
    hr = CoAllowSetForegroundWindow(m_objRhino, 0)
    'If hr >= 0 Then
    '  MsgBox "CoAllowSetForegroundWindow = SUCCESS"
    'Else
    '  MsgBox "CoAllowSetForegroundWindow = FAILED"
    'End If
    m_objRhino.Visible = True
    m_objRhino.BringToTop
  End If
End Sub

Private Sub btnOpen_Click()
  ' This only works for Rhino5 and Rhino5x64 objects
  If Not (m_objRhino Is Nothing) Then
    Call m_objRhino.RunScript("_Open", 0)
  End If
End Sub

Private Sub btnHide_Click()
  If Not (m_objRhino Is Nothing) Then
    m_objRhino.Visible = False
  End If
End Sub

Private Sub btnUnload_Click()
  If Not (m_objRhino Is Nothing) Then
    If (m_bRhino5 = True) Then
      If (chkRelease = 1) Then
        m_objRhino.ReleaseWithoutClosing = 1
      End If
    End If
  End If
  Set m_objRhino = Nothing
End Sub

