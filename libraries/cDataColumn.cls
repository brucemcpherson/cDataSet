'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:42 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414216/raw/cDataColumn.cls
' a collection of data Cells representing one column of data
' v2.04 -
Option Explicit
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Private pCollect As Collection                      ' a collection of data Cells - one for every row in this column
Private pWhere As Range
Private pParent As cDataSet
Private pColumn As Long
Private pTypeofColumn As eTypeofColumn
Private pHeadingCell As cCell    ' we can use this to find the heading for this column
Private pDirty As Boolean


Public Property Get googleType() As String
    Select Case pTypeofColumn
        Case eTCnumeric
            googleType = "number"
        Case eTCdate
            googleType = "date"
        Case Else
            googleType = "string"
            
    End Select
End Property
Public Property Get dirty() As Boolean
    dirty = pDirty
End Property
Public Property Let dirty(p As Boolean)
        pDirty = p
End Property
Public Property Get typeofColumn() As eTypeofColumn
    typeofColumn = pTypeofColumn
End Property
Public Property Let typeofColumn(p As eTypeofColumn)
    pTypeofColumn = p
End Property
Public Property Get column() As Long
    column = pColumn
End Property
Public Property Get rows() As Collection
    Set rows = pCollect
End Property
Public Property Get parent() As cDataSet
    Set parent = pParent
End Property
Public Property Get where() As Range
    If Not pWhere Is Nothing Then
        Set where = pWhere.Resize(pParent.rows.count)
    End If
End Property
Public Property Get cell(rowID As Variant) As cCell
    Set cell = pParent.cell(rowID, pHeadingCell.column)
End Property
Public Property Get value(rowID As Variant) As Variant
    value = cell(rowID).value
End Property
Public Function refresh(Optional rowID As Variant) As Variant
    Dim dt As cCell
    If IsMissing(rowID) Then
        For Each dt In rows
            refresh = dt.refresh
        Next dt
        refresh = Empty
    Else
        refresh = cell(rowID).refresh
    End If

End Function
Public Function filtered(v As Variant) As Collection
    ' this creates a filtered collection of cells for this column based on matching some value
    Dim c As Collection, cc As cCell
    Set c = New Collection
    For Each cc In rows
        ' this filter is in addition to any excel ones in operations
        If Not cc.parent.hidden And v = cc.value Then c.add cc
    Next cc
    Set filtered = c
End Function


Public Property Get uniqueValues(Optional es As eSort = eSortNone) As Collection
    ' return a collection of unique values for this column
    Dim cc As cCell
    Dim vUnique As Collection
    Set vUnique = New Collection
    
    For Each cc In rows
        If (Not cc.parent.hidden) Then
            If exists(vUnique, cc.toString) Is Nothing Then vUnique.add cc, CStr(cc.value)
        End If
    Next cc
    If es <> eSortNone Then SortColl vUnique, es

    Set uniqueValues = vUnique
End Property
Public Sub Commit(Optional p As Variant, Optional rowID As Variant)
    Dim dt As cCell, v As Variant
    
    If IsMissing(rowID) Then
        For Each dt In pCollect
            dt.Commit p
        Next dt
    Else
       cell(rowID).Commit p
    End If
   
End Sub
Public Property Get values() As Variant
    Dim cc As cCell
    ReDim a(1 To parent.visibleRowsCount) As Variant
    For Each cc In rows
        If Not cc.parent.hidden Then a(cc.row) = cc.value
    Next cc
    values = a
End Property
Public Function find(v As Variant) As cCell
    Dim cc As cCell
    For Each cc In rows
        If makeKey(cc.value) = makeKey(v) Then
            Set find = cc
            Exit Function
        End If
    Next cc
End Function
Public Function max() As Variant
    max = Application.WorksheetFunction.max(values)
End Function
Public Function min() As Variant
    min = Application.WorksheetFunction.min(values)
End Function
Public Property Get toString(rowNum As Long, Optional sFormat As String = vbNullString) As String
    toString = cell(rowNum).toString(sFormat)
End Property
Public Function create(dset As cDataSet, hcell As cCell, ncol As Long) As cDataColumn
    Dim rCell As Range, dcell As cCell
    pTypeofColumn = eTCunknown
    Set pParent = dset

    pColumn = ncol
    If Not pParent.where Is Nothing Then
        Set pWhere = hcell.where.Offset(1).Resize(dset.where.rows.count)
    End If
    Set pHeadingCell = hcell
    Set create = Me
End Function
Private Function exists(vCollect As Collection, sid As Variant) As cCell
    If Not vCollect Is Nothing Then
        On Error GoTo handle
        Set exists = vCollect(sid)
        Exit Function
    End If
handle:
    Set exists = Nothing
End Function

Public Sub tearDown()
    ' clean up
    Set pCollect = Nothing
    Set pParent = Nothing
End Sub

Private Sub Class_Initialize()
    Set pCollect = New Collection
End Sub




