'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:43 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414216/raw/cHeadingRow.cls
' a collection of Cells that contain the headings associated with a dataset
' v2.03 - 3414216
Option Explicit
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Private pDataRow As cDataRow
Public Property Get parent() As cDataSet
    Set parent = pDataRow.parent
End Property
Public Property Get dataRow() As cDataRow
    Set dataRow = pDataRow
End Property
Public Property Get headings() As Collection
    Set headings = pDataRow.columns
End Property
Public Property Get where() As Range
    Set where = pDataRow.where
End Property
Public Function create(dset As cDataSet, rHeading As Range, Optional keepFresh As Boolean = False) As cHeadingRow
    Dim rCell As Range, hcell As cCell, n As Long, dr As cDataRow

    With pDataRow
        .create dset, rHeading, 0, keepFresh
    End With
    Set create = Me
End Function
Public Function exists(s As String) As cCell
    If headings.count > 0 Then
        On Error GoTo handle
        Set exists = headings(makeKey(s))
        Exit Function
    End If
handle:
    Set exists = Nothing
End Function
Public Property Get headingList() As String
    ' return a comma separated list of the headings
    Dim t As cStringChunker, cc As cCell
    Set t = New cStringChunker
    For Each cc In headings
        t.add cc.toString & ","
    Next cc
    ' remove final comma if there is one
    headingList = t.chop.content
    Set t = Nothing
End Property

Public Function validate(complain As Boolean, ParamArray args() As Variant) As Boolean
    Dim i As Long, s As String
    s = ""
    For i = LBound(args) To UBound(args)
        If exists(CStr(args(i))) Is Nothing Then
            s = s & args(i) & ","
        End If
    Next i
    If Len(s) = 0 Then
        validate = True
    Else
        s = left(s, Len(s) - 1)
        If complain Then
            MsgBox "The following required columns are missing from dataset " & parent.name & ":" & s
        End If
    End If
End Function
Public Sub tearDown()
    ' clean up
    pDataRow.tearDown
    Set pDataRow = Nothing
End Sub

Private Sub Class_Initialize()
    Set pDataRow = New cDataRow
End Sub

