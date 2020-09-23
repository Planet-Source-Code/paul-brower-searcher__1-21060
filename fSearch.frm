VERSION 5.00
Begin VB.Form fSearch 
   Caption         =   "Search"
   ClientHeight    =   1980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "fSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   8160
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   555
      Left            =   2400
      Picture         =   "fSearch.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOpenFilter 
      Caption         =   "&Open"
      Height          =   555
      Left            =   1620
      Picture         =   "fSearch.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdSort 
      Height          =   315
      Index           =   0
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "None"
      Top             =   1620
      Width           =   375
   End
   Begin VB.PictureBox picDesc 
      Height          =   375
      Left            =   2880
      Picture         =   "fSearch.frx":0E1E
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   18
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox picAsc 
      Height          =   375
      Left            =   3240
      Picture         =   "fSearch.frx":0F40
      ScaleHeight     =   315
      ScaleWidth      =   255
      TabIndex        =   17
      Top             =   60
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "&Execute"
      Height          =   555
      Left            =   840
      Picture         =   "fSearch.frx":1062
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   555
      Left            =   60
      Picture         =   "fSearch.frx":15EC
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Rese&t"
      Height          =   315
      Left            =   5880
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   780
      Width           =   1875
   End
   Begin VB.OptionButton OptOR 
      Caption         =   "Search 'OR'"
      Height          =   315
      Left            =   4200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   1635
   End
   Begin VB.OptionButton optAND 
      Caption         =   "Search 'AND'"
      Height          =   315
      Left            =   4200
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   60
      Width           =   1635
   End
   Begin VB.CommandButton cmdList 
      Height          =   315
      Index           =   0
      Left            =   4680
      Picture         =   "fSearch.frx":1B76
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Click to build list of values"
      Top             =   1620
      Width           =   315
   End
   Begin VB.CommandButton cmdRemoveLine 
      Caption         =   "&Remove Last Line"
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   1875
   End
   Begin VB.CommandButton cmdAddLine 
      Caption         =   "+"
      Height          =   315
      Index           =   0
      Left            =   7860
      TabIndex        =   5
      Top             =   1620
      Width           =   255
   End
   Begin VB.ComboBox cboSearch 
      Height          =   315
      Index           =   0
      Left            =   5100
      TabIndex        =   4
      Top             =   1620
      Width           =   2715
   End
   Begin VB.ComboBox cboType 
      Height          =   315
      Index           =   0
      Left            =   2880
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1620
      Width           =   1755
   End
   Begin VB.ComboBox cboField 
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1620
      Width           =   2235
   End
   Begin VB.Label lblInfo 
      Caption         =   "Field:"
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   5235
   End
   Begin VB.Label Label4 
      Caption         =   "Criteria:"
      Height          =   315
      Left            =   5100
      TabIndex        =   10
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Filter Type:"
      Height          =   315
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Sort:"
      Height          =   315
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Field:"
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   2235
   End
End
Attribute VB_Name = "fSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SEARCH_TYPE$ = "=|<>|>=|<=|Like|Is Null|Is Not Null"

Private i               As Long
Private m_sReturnSQL    As String
Private m_aFilterFields As Variant
Private m_aFieldFilter  As Variant
Private m_sView         As String
Private m_oFrm          As Form

Private Property Get MyParent() As String
    MyParent = m_sView
End Property

Public Function Init(sTableViewName As String, sFieldList As String, Optional ByRef f As Form)
    
    m_sView = sTableViewName
    
    If Not f Is Nothing Then
        Set m_oFrm = f
    End If
    
    m_aFilterFields = Split(sFieldList, ",")
    
    With cboField(0)
        For i = LBound(m_aFilterFields, 1) To UBound(m_aFilterFields, 1)
            .AddItem m_aFilterFields(i)
        Next i
    End With
    Top = 0
    Left = 0
    Height = 2565
    Width = 8280
    optAND.Value = True
    Caption = "SEARCH: " '& frm.MyParent()
    lblInfo.Caption = "Enter Search Criteria.  Press F12 or Execute to run filter." & vbCrLf & _
    "You may leave filter type and criterial blank, to add sorting."
    
    Show


End Function



Private Sub cboType_DropDown(Index As Integer)
    cboType(Index).Clear
    If Len(cboField(Index).Text) > 0 Then
        m_aFieldFilter = Split(SEARCH_TYPE, "|")
        With cboType(Index)
            For i = LBound(m_aFieldFilter, 1) To UBound(m_aFieldFilter, 1)
                .AddItem m_aFieldFilter(i)
            Next i
        End With
    End If
End Sub


Private Sub cmdAddLine_Click(Index As Integer)

    If Len(cboField(Index).Text) = 0 Then
        MsgBox "Please select a field."
        cboField(Index).SetFocus
        Exit Sub
    End If


    'if criteria is not null, make sure type is selected
        
    If Len(cboSearch(Index).Text) > 0 And Len(cboType(Index).Text) = 0 Then
        MsgBox "You entered search criteria without selecting a filter type."
        cboType(Index).SetFocus
        Exit Sub
    End If
    If Len(cboSearch(Index).Text) = 0 And Len(cboType(Index).Text) > 0 Then
        MsgBox "You entered a filter type without specifying search criteria."
        cboSearch(Index).SetFocus
        Exit Sub
    End If
        
    AddLine

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Public Function RunFilter()

    Dim tSQL As String
    Dim tWHERE As String
    Dim tORDERBY As String
    Dim AND_OR As String
    Dim sWrapper As String
        
    If optAND.Value = True Then
        AND_OR = " and "
    Else
        AND_OR = " or "
    End If
    
    'Build SQL Statement
    
    tSQL = "Select * from " & MyParent & " "
    
    'Build Where clause
    For i = cboField.LBound To cboField.UBound
        If Len(cboField(i).Text) > 0 Then
            If IsNumeric(cboSearch(i).Text) = True Then
                    sWrapper = ""
            ElseIf IsDate(cboSearch(i).Text) = True Then
                    sWrapper = "#"
            Else
                    sWrapper = "'"
            End If
            If Len(cboType(i).Text) > 0 And Len(cboSearch(i).Text) > 0 Then
                tWHERE = tWHERE & " " & cboField(i).Text & " " & cboType(i).Text & " " & sWrapper & cboSearch(i).Text & sWrapper & " " & AND_OR
            End If
            If cmdSort(i).Tag <> "None" Then
                'IF FIELD IS NOT ALREADY IN SORT ORDER
                If InStr(tORDERBY, cboField(i).Text & " ") > 0 Then
                    'do nothing
                Else
                    tORDERBY = tORDERBY & cboField(i).Text & " " & cmdSort(i).Tag & ", "
                End If
            End If
        End If
    Next i

    'strip off the last and/or
    tWHERE = Left(tWHERE, Len(tWHERE) - 4)
    tWHERE = " where " & tWHERE & " "
    If Len(tORDERBY) > 0 Then
        tORDERBY = "Order By " & Left(tORDERBY, Len(tORDERBY) - 2)
        tSQL = tSQL & tWHERE & tORDERBY
    Else
        tSQL = tSQL & tWHERE
    End If
    tSQL = Trim(tSQL) & ";"
    DoEvents
    
    MsgBox tSQL
    
    Unload Me

End Function

Sub cmdExecute_Click()

    RunFilter

End Sub

Private Sub cmdList_Click(Index As Integer)

'    If Len(cboField(Index).Text) > 0 Then
'        Select Case fBrow.FieldDataType(cboField(Index).Text)
'        Case "Text"
'            cboSearch(Index).Clear
'            ParentANI True, "Gears.avi", "getting field values ..."
'
'                Dim rs As DAO.Recordset
'                Dim SQL As String
'                SQL = "SELECT " & fBrow.MyParent & "." & cboField(Index).Text & " " & _
'                    "From " & fBrow.MyParent & " " & _
'                    "GROUP BY " & fBrow.MyParent & "." & cboField(Index).Text & " " & _
'                    "HAVING (((" & fBrow.MyParent & "." & cboField(Index).Text & ") Is Not Null));"
'                Set rs = DaoRS(SQL, dbOpenSnapshot)
'
'                If Not rs.EOF Then
'                    rs.MoveFirst
'                    Do Until rs.EOF
'                        With cboSearch(Index)
'                            .AddItem rs.Fields(0)
'                        End With
'                        rs.MoveNext
'                    Loop
'                End If
'
'                Set rs = Nothing
'
'
'        End Select
'    End If
    
    MsgBox "put your own code here"
End Sub


Private Sub cmdOpenFilter_Click()

    'OpenFilter

End Sub

Private Sub cmdRemoveLine_Click()
    If cboField.UBound = 0 Then
    Else
        Unload cboField(cboField.UBound)
        Unload cboType(cboType.UBound)
        Unload cboSearch(cboSearch.UBound)
        Unload cmdSort(cmdSort.UBound)
        Unload cmdList(cmdList.UBound)
        Unload cmdAddLine(cmdAddLine.UBound)
        
        Me.Height = Me.Height - 415
    End If
    
    cboField(cboField.UBound).SetFocus
    
End Sub


Private Sub cmdReset_Click()

    ResetLines

End Sub

Private Function SaveSearch()
    
    'save value to Search.INI
    
'    Dim lMaxRow As Long
'    For i = cboField.UBound To 0 Step -1
'        If Len(cboField(i).Text) > 0 Then
'            lMaxRow = i
'            Exit For
'        End If
'    Next i
'
'    'Save the number of rows used
'    SetInitEntry "SEARCH-" & fBrow.MyParent, "Rows", CStr(lMaxRow), "Search"
'    If optAND.Value = True Then
'        SetInitEntry "SEARCH-" & fBrow.MyParent, "ANDOR", "AND", "Search"
'    Else
'        SetInitEntry "SEARCH-" & fBrow.MyParent, "ANDOR", "OR", "Search"
'    End If
'
'    For i = 0 To lMaxRow
'        SetInitEntry "SEARCH-" & fBrow.MyParent, "Row" & i, cboField(i).ListIndex & "|" & _
'            cmdSort(i).Tag & "|" & cboType(i).ListIndex & "|" & cboSearch(i).Text, "Search"
'    Next i
'
'    Set fBrow = Nothing
'    Set m_oib = Nothing

    MsgBox "put your code here"

End Function

Private Sub cmdSave_Click()

    SaveView

End Sub

Private Sub optAND_Click()

    OptOR.Value = False

End Sub

Private Sub OptOR_Click()
    optAND.Value = False
End Sub

Private Sub cmdSort_Click(Index As Integer)

    Select Case cmdSort(Index).Tag
        Case "Asc"
            cmdSort(Index).Tag = "Desc"
            cmdSort(Index).Picture = picDesc.Picture
        Case "Desc"
            cmdSort(Index).Tag = "None"
            cmdSort(Index).Picture = Nothing
        Case "None"
            cmdSort(Index).Tag = "Asc"
            cmdSort(Index).Picture = picAsc.Picture
    End Select

End Sub

Private Function AddLine()
    If cboField.UBound >= 20 Then
        MsgBox "You have exceeded the maximum fields you can search on."
        Exit Function
    End If
    
    
    Dim lLoop As Long
    i = cboField.UBound
    
    Load cboField(i + 1)
    Load cmdSort(i + 1)
    Load cboType(i + 1)
    Load cboSearch(i + 1)
    Load cmdList(i + 1)
    Load cmdAddLine(i + 1)
        
    cboField(i + 1).Top = (cboField(0).Top + 100 + cboField(0).Height) + (415 * i)
    cboField(i + 1).TabIndex = cboField(i).TabIndex + 6
    cboField(i + 1).Visible = True
        
    With cboField(i + 1)
        For lLoop = LBound(m_aFilterFields, 1) To UBound(m_aFilterFields, 1)
            .AddItem m_aFilterFields(lLoop)
        Next lLoop
    End With
    
    cmdSort(i + 1).Top = (cmdSort(0).Top + 100 + cmdSort(0).Height) + (415 * i)
    cmdSort(i + 1).Tag = "None"
    cmdSort(i + 1).TabIndex = cmdSort(i).TabIndex + 6
    cmdSort(i + 1).Picture = Nothing
    cmdSort(i + 1).Visible = True
    
    cboType(i + 1).Top = (cboType(0).Top + 100 + cboType(0).Height) + (415 * i)
    cboType(i + 1).TabIndex = cboType(i).TabIndex + 6
    cboType(i + 1).Visible = True
    
    cmdList(i + 1).Top = (cmdList(0).Top + 100 + cmdList(0).Height) + (415 * i)
    cmdList(i + 1).TabIndex = cmdList(i).TabIndex + 6
    cmdList(i + 1).Visible = True
    
    cboSearch(i + 1).Top = (cboSearch(0).Top + 100 + cboSearch(0).Height) + (415 * i)
    cboSearch(i + 1).Text = ""
    cboSearch(i + 1).TabIndex = cboSearch(i).TabIndex + 6
    cboSearch(i + 1).Visible = True
    
    cmdAddLine(i + 1).Top = (cmdAddLine(0).Top + 100 + cmdAddLine(0).Height) + (415 * i)
    cmdAddLine(i + 1).TabIndex = cmdAddLine(i).TabIndex + 6
    cmdAddLine(i + 1).Visible = True
    
    
    Me.Height = Me.Height + 415

    cboField(cboField.UBound).SetFocus

End Function
Private Function GetSavedSettings()
'    Dim vinfo As Variant
'    Dim pCount As Long
'    If GetInitEntry("SEARCH-" & fBrow.MyParent, "ANDOR", "AND", "Search") = "AND" Then
'        optAND.Value = True
'        OptOR.Value = False
'    Else
'        optAND.Value = False
'        OptOR.Value = True
'    End If
'    If CLng(GetInitEntry("SEARCH-" & fBrow.MyParent, "Rows", CStr(99), "Search")) <> 99 Then
'        Dim lMaxRows As Long
'        lMaxRows = CLng(GetInitEntry("SEARCH-" & fBrow.MyParent, "Rows", CStr(99), "Search"))
'        'Add lmaxrows to control arrays
'        vinfo = Split(GetInitEntry("SEARCH-" & fBrow.MyParent, "Row0", , "Search"), "|")
'        PopSavedInfo 0, vinfo
'        For pCount = 1 To lMaxRows
'            AddLine
'            vinfo = Split(GetInitEntry("SEARCH-" & fBrow.MyParent, "Row" & pCount, , "Search"), "|")
'            PopSavedInfo pCount, vinfo
'        Next pCount
'    End If
    MsgBox "your own code here"
End Function


Private Function PopSavedInfo(lIdx As Long, vinfo As Variant)
'    Dim NWI As Long
'    If CLng(vinfo(0)) <= cboField(lIdx).ListCount - 1 Then
'        cboField(lIdx).ListIndex = CLng(vinfo(0))
'    End If
'    cmdSort(lIdx).Tag = vinfo(1)
'        If vinfo(1) = "Asc" Then
'            cmdSort(lIdx).Picture = picAsc.Picture
'        ElseIf vinfo(1) = "Desc" Then
'            cmdSort(lIdx).Picture = picDesc.Picture
'        ElseIf vinfo(1) = "None" Then
'            cmdSort(lIdx).Picture = Nothing
'        End If
'
'        If Len(cboField(lIdx).Text) > 0 Then
'
'            m_aFieldFilter = Split(fBrow.FieldFilter(cboField(lIdx).Text), "|")
'            With cboType(lIdx)
'                For NWI = LBound(m_aFieldFilter, 1) To UBound(m_aFieldFilter, 1)
'                    .AddItem m_aFieldFilter(NWI)
'                Next NWI
'            End With
'        End If
'
'    If CLng(vinfo(2)) <= cboType(lIdx).ListCount - 1 Then
'        cboType(lIdx).ListIndex = CLng(vinfo(2))
'    End If
'
'    cboSearch(lIdx).Text = vinfo(3)
End Function

Public Function SaveView()
    
'    Dim sViewName As String
'    Dim bAvailToAll As Boolean
'    Dim sGUID As String
'
'    Dim rsSQL As DAO.Recordset
'    Dim rsSQLDet As DAO.Recordset
'
'    sViewName = InputBox("Enter Name Of View", "Save View")
'    If Len(sViewName) > 0 Then
'        Msg = "Make this view public?"
'        X = MsgBox(Msg, vbYesNoCancel + vbQuestion)
'        If X = vbYes Then
'            bAvailToAll = True
'        ElseIf X = vbNo Then
'            bAvailToAll = False
'        Else
'            Exit Function
'        End If
'    End If
'
'    sGUID = CreateGUID
'    If sGUID = "Error" Then
'        MsgBox "Error creating guid, please see developer"
'        Exit Function
'    End If
'
'    Set rsSQL = DaoRS("Select * from CustomSQL;", , dbAppendOnly)
'    With rsSQL
'        .AddNew
'        !Guid = sGUID
'        !UserLogon = PWK.UserName
'        !SQLName = sViewName
'        !AvailToAll = bAvailToAll
'        !Browser = fBrow.MyParent
'        If optAND.Value = True Then
'            !AND_OR = "And"
'        Else
'            !AND_OR = "Or"
'        End If
'        .Update
'    End With
'    rsSQL.Close: Set rsSQL = Nothing
'    Set rsSQLDet = DaoRS("select * from customsql_det;", , dbAppendOnly)
'
'    Dim lMaxRow As Long
'    For i = cboField.UBound To 0 Step -1
'        If Len(cboField(i).Text) > 0 Then
'            lMaxRow = i
'            Exit For
'        End If
'    Next i
'
'    For i = 0 To lMaxRow
''        SetInitEntry "SEARCH-" & fBrow.MyParent, "Row" & i, cboField(i).ListIndex & "|" & _
'            cmdSort(i).Tag & "|" & cboType(i).ListIndex & "|" & cboSearch(i).text, "Search"
'        With rsSQLDet
'            .AddNew
'            !Guid = sGUID
'            !FieldName = cboField(i).ListIndex
'            !Sort = cmdSort(i).Tag
'            !FilterType = cboType(i).ListIndex
'            !Criteria = cboSearch(i).Text
'            .Update
'        End With
'    Next i
'
'    rsSQLDet.Close
'    Set rsSQLDet = Nothing
'
'    MsgBox "Filter saved successfully."
    
    
End Function

Private Function ResetLines()
    
    Do Until cboField.UBound = 0
        
        Unload cboField(cboField.UBound)
        Unload cboType(cboType.UBound)
        Unload cboSearch(cboSearch.UBound)
        Unload cmdSort(cmdSort.UBound)
        Unload cmdList(cmdList.UBound)
        Unload cmdAddLine(cmdAddLine.UBound)
        
        Me.Height = Me.Height - 415
    Loop

End Function

Public Function GetSavedFilter(sGUID As String)

'    Dim rsSQL As DAO.Recordset
'    Dim rsSQLDet As DAO.Recordset
'
'    ResetLines
'
'    Set rsSQL = DaoRS("select * from customsql where GUID='" & sGUID & "';", dbOpenSnapshot)
'    If rsSQL.EOF Then
'        Set rsSQL = Nothing
'        Exit Function
'    End If
'
'    rsSQL.MoveFirst
'    If rsSQL!AND_OR = "And" Then
'        optAND.Value = True
'        OptOR.Value = False
'    Else
'        optAND.Value = False
'        OptOR.Value = True
'    End If
'
'    Set rsSQLDet = DaoRS("select * from customsql_det where guid = '" & sGUID & "' order by id;", dbOpenSnapshot)
'
'    rsSQLDet.MoveLast: rsSQLDet.MoveFirst
'
'    Dim vinfo As Variant
'    Dim pCount As Long
'
'    Dim lMaxRows As Long
'    lMaxRows = rsSQLDet.RecordCount
'    'Add lmaxrows to control arrays
'
'    With rsSQLDet
'        .MoveFirst
'        vinfo = Split(.Fields("FieldName") & "|" & _
'            .Fields("Sort") & "|" & _
'            .Fields("FilterType") & "|" & _
'            .Fields("Criteria"), "|")
'        PopSavedInfo 0, vinfo
'        For pCount = 1 To lMaxRows - 1
'            .MoveNext
'            AddLine
'            vinfo = Split(.Fields("FieldName") & "|" & _
'                .Fields("Sort") & "|" & _
'                .Fields("FilterType") & "|" & _
'                .Fields("Criteria"), "|")
'            PopSavedInfo pCount, vinfo
'            If .EOF Then Exit For
'        Next pCount
'
'    End With
'
'    rsSQL.Close: Set rsSQL = Nothing
'    rsSQLDet.Close: Set rsSQLDet = Nothing
    
    
    
    
End Function




Public Function OpenFilter()

'    Dim fSF As New fSavedFilter
'    fSF.Init Me
'    Set fSF = Nothing

End Function
