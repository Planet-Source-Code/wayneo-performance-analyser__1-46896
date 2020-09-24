VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Performance Tester"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   7740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIterations 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Text            =   "1000"
      Top             =   1250
      Width           =   975
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Begin Testing"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1250
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvwReport 
      Height          =   1935
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Test Number"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Total Time"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Avg Call Time"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Best Call Time"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Slowest Call Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Relative Call Time"
         Object.Width           =   2734
      EndProperty
   End
   Begin VB.CheckBox chkEnableTest 
      Caption         =   "Enable Test (4)"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkEnableTest 
      Caption         =   "Enable Test (3)"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkEnableTest 
      Caption         =   "Enable Test (2)"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CheckBox chkEnableTest 
      Caption         =   "Enable Test (1)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status: Idle"
      ForeColor       =   &H8000000D&
      Height          =   280
      Left            =   4320
      TabIndex        =   13
      Top             =   1250
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   975
      Left            =   1800
      TabIndex        =   12
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Iterations of each:"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1270
      Width           =   1335
   End
   Begin VB.Label lblPerformanceGraph 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   3
      Left            =   2280
      TabIndex        =   9
      Top             =   900
      Width           =   5295
   End
   Begin VB.Label lblPerformanceGraph 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   660
      Width           =   5295
   End
   Begin VB.Label lblPerformanceGraph 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   1
      Left            =   2280
      TabIndex        =   7
      Top             =   420
      Width           =   5295
   End
   Begin VB.Label lblPerformanceGraph 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   0
      Left            =   2295
      TabIndex        =   6
      Top             =   180
      Width           =   5280
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PerformanceDetails
    lBestTime As Long
    lWorstTime As Long
    lTotalTime As Long
    lAverageTime As Long
End Type
 
'------------------------------------------------------
'To test your code, open the (modTestFunctions) module
'------------------------------------------------------


Private Sub cmdExecute_Click()
    Dim Times(3) As PerformanceDetails
    
    Dim timer As New PrecisionTimer
    
    Dim lBestTime As Long
    Dim lWorstTime As Long
    Dim lTotalTime As Long
    
    Dim lMaxIterations As Long
    lMaxIterations = CLng(txtIterations.Text)
    
    Dim lTimeElapsed As Long
    Dim sTempElapsed As String

    
    '---Clean The Current Data / Graph---
    Dim lvwitem As ListItem
    lvwReport.ListItems.Clear
    '------------------------------------
    
    Dim i As Long
    
    
    If chkEnableTest(0).Value = Checked Then
        lblStatus = "Status: Performing Function 1"
        
        lBestTime = 20000000                        '-should- always get set lower
        lWorstTime = 0                              'It will always get set higher
        lTotalTime = 0                              'Always reset to 0
        
        For i = 1 To lMaxIterations
            timer.ResetTimer    'Begin Timing
            
                Test1
            
            timer.StopTimer     'Finish Timing
            
            sTempElapsed = timer.Elapsed                                'Value needs to be sent to a variable
            lTimeElapsed = CLng(sTempElapsed)
            
            'lTimeElapsed = Left(sTempElapsed, lDecPoint)        'remove the .00 things and the microsecond time thing
            If lTimeElapsed > lWorstTime Then lWorstTime = lTimeElapsed
            If lTimeElapsed < lBestTime Then lBestTime = lTimeElapsed
            lTotalTime = lTotalTime + lTimeElapsed
            
            DoEvents
        Next
            
        With Times(0)
            .lAverageTime = lTotalTime / lMaxIterations
            .lBestTime = lBestTime
            .lWorstTime = lWorstTime
            .lTotalTime = lTotalTime
        End With

    End If
    
    
    
    
    If chkEnableTest(1).Value = Checked Then
        lBestTime = 20000000                        '-should- always get set lower
        lWorstTime = 0                              'It will always get set higher
        lTotalTime = 0                              'Always reset to 0
        
        lblStatus = "Status: Performing Function 2"
        
        For i = 1 To lMaxIterations
            timer.ResetTimer    'Begin Timing
            
                Test2
            
            timer.StopTimer     'Finish Timing
            
            sTempElapsed = timer.Elapsed                                'Value needs to be sent to a variable
            lDecPoint = InStr(1, sTempElapsed, ".", vbBinaryCompare)    'find the dec point
            
            lTimeElapsed = Left(sTempElapsed, lDecPoint)        'remove the .00 things and the microsecond time thing
            If lTimeElapsed > lWorstTime Then lWorstTime = lTimeElapsed
            If lTimeElapsed < lBestTime Then lBestTime = lTimeElapsed
            lTotalTime = lTotalTime + lTimeElapsed
            
            DoEvents
        Next
            
        With Times(1)
            .lAverageTime = lTotalTime / lMaxIterations
            .lBestTime = lBestTime
            .lWorstTime = lWorstTime
            .lTotalTime = lTotalTime
        End With

    End If
    
    
    If chkEnableTest(2).Value = Checked Then
        lBestTime = 20000000                        '-should- always get set lower
        lWorstTime = 0                              'It will always get set higher
        lTotalTime = 0                              'Always reset to 0
        
        lblStatus = "Status: Performing Function 3"
        
        For i = 1 To lMaxIterations
            timer.ResetTimer    'Begin Timing
            
                Test3
            
            timer.StopTimer     'Finish Timing
            
            sTempElapsed = timer.Elapsed                                'Value needs to be sent to a variable
            lDecPoint = InStr(1, sTempElapsed, ".", vbBinaryCompare)    'find the dec point
            
            lTimeElapsed = Left(sTempElapsed, lDecPoint)        'remove the .00 things and the microsecond time thing
            If lTimeElapsed > lWorstTime Then lWorstTime = lTimeElapsed
            If lTimeElapsed < lBestTime Then lBestTime = lTimeElapsed
            lTotalTime = lTotalTime + lTimeElapsed
            
            DoEvents
        Next
        
        With Times(2)
            .lAverageTime = lTotalTime / lMaxIterations
            .lBestTime = lBestTime
            .lWorstTime = lWorstTime
            .lTotalTime = lTotalTime
        End With

    End If
    
    
    If chkEnableTest(3).Value = Checked Then
        lBestTime = 20000000                        '-should- always get set lower
        lWorstTime = 0                              'It will always get set higher
        lTotalTime = 0                              'Always reset to 0
        
        lblStatus = "Status: Performing Function 4"
        
        For i = 1 To lMaxIterations
            timer.ResetTimer    'Begin Timing
            
                Test4
            
            timer.StopTimer     'Finish Timing
            
            sTempElapsed = timer.Elapsed                                'Value needs to be sent to a variable
            lDecPoint = InStr(1, sTempElapsed, ".", vbBinaryCompare)    'find the dec point
            
            lTimeElapsed = Left(sTempElapsed, lDecPoint)        'remove the .00 things and the microsecond time thing
            If lTimeElapsed > lWorstTime Then lWorstTime = lTimeElapsed
            If lTimeElapsed < lBestTime Then lBestTime = lTimeElapsed
            lTotalTime = lTotalTime + lTimeElapsed
            
            DoEvents
        Next
            
        With Times(3)
            .lAverageTime = lTotalTime / lMaxIterations
            .lBestTime = lBestTime
            .lWorstTime = lWorstTime
            .lTotalTime = lTotalTime
        End With

    End If
    
    lblStatus = "Status: Compiling Results"
    
    Dim lLongestCallTime As Long
    lLongestCallTime = 0
    
    'find the highest time for execution
    For i = 0 To 3
        If chkEnableTest(i).Value = Checked Then
            If Times(i).lTotalTime > lLongestCallTime Then lLongestCallTime = Times(i).lTotalTime
        Else
            lblPerformanceGraph(i).Width = 0
        End If
    Next
    
    
    
    Dim lRelative As Long
    
    For i = 0 To 3
        If chkEnableTest(i).Value = Checked Then
            Set lvwitem = lvwReport.ListItems.Add
            lvwitem.Text = "Function " & i + 1
            lvwitem.SubItems(1) = Times(i).lTotalTime & " µs"
            lvwitem.SubItems(2) = Times(i).lTotalTime / lMaxIterations & " µs"
            lvwitem.SubItems(3) = Times(i).lBestTime & " µs"
            lvwitem.SubItems(4) = Times(i).lWorstTime & " µs"
            
            lRelative = (Times(i).lTotalTime / lLongestCallTime) * 100
            
            lblPerformanceGraph(i).Width = lRelative * 50
            
            lvwitem.SubItems(5) = lRelative & "%"
            If lRelative > 90 Then
                lblPerformanceGraph(i).BackColor = &HFF&
            ElseIf lRelative > 80 Then
                lblPerformanceGraph(i).BackColor = &H80&
            ElseIf lRelative > 70 Then
                lblPerformanceGraph(i).BackColor = &HC0C0&
            ElseIf lRelative > 50 Then
                lblPerformanceGraph(i).BackColor = &H80FF&
            ElseIf lRelative <= 50 Then
                lblPerformanceGraph(i).BackColor = &HC000&
            ElseIf lRelative <= 40 Then
                lblPerformanceGraph(i).BackColor = &HFF00&
            End If
        End If
    Next
    
    lblStatus = "Status: Idle"
        
End Sub


Private Sub Form_Load()
    For i = 0 To 3
        lblPerformanceGraph(i).Width = 0
    Next
End Sub

