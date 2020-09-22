VERSION 5.00
Begin VB.Form frmSortNatural 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   6480
      Width           =   2775
   End
   Begin VB.ListBox List3 
      Height          =   5910
      Left            =   5640
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Height          =   5910
      Left            =   2880
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sort It"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   6300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Normal Sort (ASCII)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Natural Sort (Alphabetically)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmSortNatural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const sortNaturalSmaller As Long = -1
Const sortNaturalEqual As Long = 0
Const sortNaturalLarger As Long = 1

'This is the function that does the natural sort
Public Function sortNatural(s1 As String, s2 As String) As Long
    sortNatural = 1
    If s1 = s2 Then
        sortNatural = 0
    ElseIf IsNumeric(s1) Then
        sortNatural = -1
        If IsNumeric(s2) Then If Val(s1) > Val(s2) Then sortNatural = 1
    ElseIf IsNumeric(s2) Then
        sortNatural = 1
        If IsNumeric(s1) Then If Val(s1) < Val(s2) Then sortNatural = -1
    ElseIf s1 Like "*[0-9]*" Then 'There's a non-number + number
        Dim sStr1 As String, lNum1 As Long, lLen1 As Long
        Dim sStr2 As String, lNum2 As Long, lLen2 As Long
        Dim lX As Long, lY As Long
        lLen1 = Len(s1)
        lLen2 = Len(s2)
        For lX = 1 To lLen1
            If IsNumeric(Mid(s1, lX, 1)) Then
                sStr1 = Strings.Mid$(s1, 1, lX - 1)
                lNum1 = Val(Strings.Mid$(s1, lX))
                Exit For
            End If
        Next lX
        For lX = 1 To lLen2
            If IsNumeric(Mid(s2, lX, 1)) Then
                sStr2 = Strings.Mid$(s2, 1, lX - 1)
                lNum2 = Val(Strings.Mid$(s2, lX))
                Exit For
            End If
        Next lX
        If sStr1 = sStr2 Then
            If lNum1 < lNum2 Then sortNatural = -1
        Else
            If sStr1 < sStr2 Then sortNatural = -1
        End If
    Else 'Just strings
        sortNatural = -1
        If s1 > s2 Then sortNatural = 1
    End If
End Function

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim asData(300) As String, alIndex() As Long, lX As Long
    Dim sngT1 As Single, sngT2 As Single
    Randomize Timer
    For lX = 0 To 100
        asData(lX) = "aaaa " & CLng(Rnd * 20 + 1)
    Next lX
    For lX = 101 To 200
        asData(lX) = "bbbb " & CLng(Rnd * 20 + 1)
    Next lX
    For lX = 200 To 210 'This data will show that there are still some bugs to fix in the routine
        asData(lX) = "1 " & Chr$(64 + CLng(Rnd * 20 + 1))
    Next lX
    For lX = 210 To 300
        asData(lX) = CLng(Rnd * 20 + 1)
    Next lX
    
    'Fill in unsorted data
    For lX = 0 To 300
        List1.AddItem asData(lX)
    Next lX
    
    'Sort naturally
    pvArraySortString asData, alIndex
    For lX = 0 To 300
        List2.AddItem asData(alIndex(lX))
    Next lX
    
    'Sort normally
    pvArraySortStringOri asData, alIndex
    For lX = 0 To 300
        List3.AddItem asData(alIndex(lX))
    Next lX
End Sub

Private Sub pvArraySortString(ByRef asSortArray() As String, ByRef alSortedIndex() As Long, Optional ByVal IgnoreCase As Boolean = True)
    Dim sVal1 As String, sVal2 As String
    Dim lX As Long, lRow As Long, lMaxRow As Long, lMinRow As Long
    Dim lSwitch As Long, lLimit As Long, lOffset As Long, lZ As Long
    lMaxRow = UBound(asSortArray)
    lMinRow = LBound(asSortArray)
    ReDim alSortedIndex(lMinRow To lMaxRow)
    For lX = lMinRow To lMaxRow
        alSortedIndex(lX) = lX
    Next
    lOffset = lMaxRow \ 2
    Do While lOffset > 0
        lLimit = lMaxRow - lOffset
        Do
            lSwitch = False
            For lRow = lMinRow To lLimit
                lZ = lZ + 1
                sVal1 = asSortArray(alSortedIndex(lRow))
                sVal2 = asSortArray(alSortedIndex(lRow + lOffset))
                If IgnoreCase Then
                    sVal1 = LCase(sVal1)
                    sVal2 = LCase(sVal2)
                End If
                If sortNatural(sVal1, sVal2) = 1 Then
                'If sVal1 > sVal2 Then 'This is the original comnpare method used, changed to the one above
                    lX = alSortedIndex(lRow)
                    alSortedIndex(lRow) = alSortedIndex(lRow + lOffset)
                    alSortedIndex(lRow + lOffset) = lX
                    lSwitch = lRow
                End If
            Next lRow
            lLimit = lSwitch - lOffset
        Loop While lSwitch
        lOffset = lOffset \ 2
    Loop
End Sub

Private Sub pvArraySortStringOri(ByRef asSortArray() As String, ByRef alSortedIndex() As Long, Optional ByVal IgnoreCase As Boolean = True)
    Dim sVal1 As String, sVal2 As String
    Dim lX As Long, lRow As Long, lMaxRow As Long, lMinRow As Long
    Dim lSwitch As Long, lLimit As Long, lOffset As Long, lZ As Long
    lMaxRow = UBound(asSortArray)
    lMinRow = LBound(asSortArray)
    ReDim alSortedIndex(lMinRow To lMaxRow)
    For lX = lMinRow To lMaxRow
        alSortedIndex(lX) = lX
    Next
    lOffset = lMaxRow \ 2
    Do While lOffset > 0
        lLimit = lMaxRow - lOffset
        Do
            lSwitch = False
            For lRow = lMinRow To lLimit
                lZ = lZ + 1
                sVal1 = asSortArray(alSortedIndex(lRow))
                sVal2 = asSortArray(alSortedIndex(lRow + lOffset))
                If IgnoreCase Then
                    sVal1 = LCase(sVal1)
                    sVal2 = LCase(sVal2)
                End If
                If sVal1 > sVal2 Then
                    lX = alSortedIndex(lRow)
                    alSortedIndex(lRow) = alSortedIndex(lRow + lOffset)
                    alSortedIndex(lRow + lOffset) = lX
                    lSwitch = lRow
                End If
            Next lRow
            lLimit = lSwitch - lOffset
        Loop While lSwitch
        lOffset = lOffset \ 2
    Loop
End Sub

