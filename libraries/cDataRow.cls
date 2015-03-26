'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:43 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414216/raw/cDataRow.cls
' a collection of data Cells representing one row of data
Option Explicit
'v 2.02
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Private pCollect As Collection                  ' a collection of data Cells - one for every column in this row
Private pWhere As Range
Private pParent As cDataSet
Private pRow As Long
Private pHidden As Boolean
Public Property Get hidden()
    hidden = pHidden
End Property
Public Property Get parent() As cDataSet
    Set parent = pParent
End Property
Public Property Get row() As Long
    row = pRow
End Property
Public Property Get columns() As Collection
    Set columns = pCollect
End Property
Public Property Get where() As Range
    Set where = pWhere
End Property

Public Property Get cell(sid As Variant, Optional complain As Boolean = False) As cCell
    Dim c As cCell
    Set c = exists(sid)
    If c Is Nothing And complain Then
        MsgBox (CStr(sid) & " is not a known column heading")
    End If
    Set cell = c

End Property
Public Property Get value(sid As Variant) As Variant
    Dim cc As cCell
    Set cc = cell(sid)
    If Not cc Is Nothing Then
        value = cc.value
    End If
End Property
Public Property Get values(Optional bIncludeKey = False) As Variant
    Dim cc As cCell
    ReDim a(1 To columns.count) As Variant
    For Each cc In columns
        If cc.column <> pParent.keyColumn Or bIncludeKey Then
            a(cc.column) = cc.value
        Else
            a(cc.column) = Empty
        End If
    Next cc
    values = a
End Property

Public Function find(v As Variant, Optional bIncludeKey = False) As cCell
    Dim cc As cCell
    For Each cc In columns
        If cc.column <> pParent.keyColumn Or bIncludeKey Then
            If makeKey(cc.value) = makeKey(v) Then
                Set find = cc
                Exit Function
            End If
        End If
    Next cc
End Function
Public Function max(Optional bIncludeKey = False) As Variant
    max = Application.WorksheetFunction.max(values(bIncludeKey))
End Function
Public Function min(Optional bIncludeKey = False) As Variant
    max = Application.WorksheetFunction.min(values(bIncludeKey))
End Function
Public Function refresh(Optional sid As Variant) As Variant
    Dim dt As cCell, v As Variant
    If IsMissing(sid) Then
        For Each dt In columns
            v = dt.refresh
        Next dt

    Else
        refresh = cell(sid).refresh
    End If
End Function

Public Sub Commit(Optional p As Variant, Optional sid As Variant)
    Dim dt As cCell
    If IsMissing(sid) Then
        For Each dt In columns
            dt.Commit p
        Next dt
    Else
      cell(sid).Commit p
    End If
   
End Sub
Public Property Get toString(sid As Variant, Optional sFormat As String = vbNullString) As String
    toString = cell(sid).toString(sFormat)
End Property
Public Function create(dset As cDataSet, rDataRow As Range, nRow As Long, _
                            rv As Variant) As cDataRow
                            
    Dim rCell As Range, dcell As cCell, hcell As cCell, hr As cHeadingRow, n As Long
    Dim r As Range, dc As cDataColumn

    Set pWhere = rDataRow
    Set pParent = dset
    pRow = nRow
    n = 0
    ' recordfilter
    pHidden = False
    If (pParent.recordFilter) Then
        pHidden = rDataRow.EntireRow.hidden
    End If
    
    If pRow = 0 Then        ' we are doing a headingrow
        For Each r In pWhere.Cells
            n = n + 1
            If IsEmpty(r) Then
                MsgBox ("unexpected blank heading cell at " & SAd(r))
                Exit Function
            End If
            Debug.Assert Not IsEmpty(r)
            Set dcell = New cCell
            With dcell
                pCollect.add .create(Me, n, r), makeKey(CStr(r.value))
            End With
        Next r
    Else
        Set hr = pParent.headingRow
        For Each hcell In hr.headings
            ' create a cell to hold it in
            Set rCell = rDataRow.Cells(1, hcell.column)
            Set dcell = New cCell
            dcell.create Me, hcell.column, rCell, rv(nRow - 1 + LBound(rv, 1), hcell.column - 1 + LBound(rv, 2))
            pCollect.add dcell

            ' set the type of column
            Set dc = pParent.columns(hcell.column)
            With dc
                If Not IsEmpty(rCell) Then
                    If .typeofColumn <> eTCmixed Then
                        If IsDate(rCell.value) Then
                            If .typeofColumn <> eTCdate Then
                                If .typeofColumn = eTCunknown Then
                                    .typeofColumn = eTCdate
                                Else
                                    .typeofColumn = eTCmixed
                                End If
                            End If
                            
                        ElseIf IsNumeric(rCell.value) Then
                            If .typeofColumn <> eTCnumeric Then
                                If .typeofColumn = eTCunknown Then
                                    .typeofColumn = eTCnumeric
                                Else
                                    .typeofColumn = eTCmixed
                                End If
                            End If
                        
                        Else
                            If .typeofColumn <> eTCtext Then
                                If .typeofColumn = eTCunknown Then
                                    .typeofColumn = eTCtext
                                Else
                                    .typeofColumn = eTCmixed
                                End If
                            End If
                        End If
                    End If
                End If
            End With
            
        Next hcell
    End If
    Set create = Me
End Function

Private Function exists(sid As Variant) As cCell
    On Error GoTo handle
    If VarType(sid) = vbLong Or VarType(sid) = vbInteger Then
        Set exists = pCollect(sid)
    Else
        Set exists = pCollect(pParent.headings(makeKey(CStr(sid))).column)
    End If
    Exit Function
handle:
    Set exists = Nothing
End Function
Public Sub tearDown()
    ' clean up
    Dim cc As cCell
    If Not pCollect Is Nothing Then
        For Each cc In columns
            cc.tearDown
        Next cc
        Set pCollect = Nothing
    End If
    
    Set pParent = Nothing
End Sub

Private Sub Class_Initialize()
    Set pCollect = New Collection
End Sub




