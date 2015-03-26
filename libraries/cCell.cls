'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:40 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414216/raw/cCell.cls
' a data Cell - holds value at time of loading, or can be kept fresh if there might be formula updates
Option Explicit
' Version 2.04 -
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
Private pValue As Variant                   ' value of cell when first loaded
Private pColumn As Long                     ' column number
Private pParent As cDataRow                 ' cDataRow to which this belongs
Public Property Get row() As Long
    row = pParent.row
End Property
Public Property Get column() As Long
    column = pColumn
End Property
Public Property Get parent() As cDataRow
    Set parent = pParent
End Property
Public Property Get myKey() As String
    myKey = makeKey(pParent.parent.headings(pColumn).toString)
End Property

Public Property Get where() As Range    ' return the range from whence it came
    If row = 0 Then
    ' its a heading
        Set where = pParent.where.Resize(1, 1).Offset(row, pColumn - 1)
    Else
        Set where = pParent.where.Resize(1, 1).Offset(, pColumn - 1)
    End If
End Property
Public Property Get refresh() As Variant ' refresh the current value and return it
    pValue = where.value
    refresh = pValue
End Property
Public Property Get toString(Optional sFormat As String = vbNullString, _
            Optional followFormat As Boolean = False, _
            Optional deLocalize As Boolean = False) As String ' Convert to a string, applying a format if supplied
    Dim s As String, os As String, ts As String
    If Len(sFormat) > 0 Then
        os = Format(value, sFormat)
    Else
        If followFormat Then
            s = where.NumberFormat
            If Len(s) > 0 And s <> "General" Then
                os = Format(value, s)
            Else
                os = CStr(value)
            End If
        Else
            os = CStr(value)
        End If
    End If
    
    If deLocalize Then
        If VarType(value) = vbDouble Or VarType(value) = vbCurrency Or VarType(value) = vbSingle Then
            ' commas to dots
            ts = Mid(CStr(1.1), 2, 1)
            os = Replace(os, ts, ".")
            
        ElseIf VarType(value) = vbBoolean Then
        
            If value Then
                os = "true"
            Else
                os = "false"
            
            End If
        
        End If
    
    End If
    toString = os
    
End Property
Public Property Get value() As Variant   ' return the value, refreshing it if necessary
    If pParent.parent.keepFresh Then
        value = refresh
    Else
        value = pValue
    End If
End Property
Public Property Let value(p As Variant)
    parent.parent.columns(pColumn).dirty = True
    If pParent.parent.keepFresh Then
        Commit p
    Else
        pValue = p
    End If
End Property
Public Function needSwap(cc As cCell, e As eSort) As Boolean
    ' this can be used from a sorting alogirthm
    Select Case e
        Case eSortAscending
            needSwap = LCase(toString) > LCase(cc.toString)
            
        Case eSortDescending
            needSwap = LCase(toString) < LCase(cc.toString)
    
        Case Else
            needSwap = False
    End Select
End Function
Public Function Commit(Optional p As Variant) As Variant
    Dim v As Variant
    If Not IsMissing(p) Then
        pValue = p
    End If
    
    where.value = pValue
    Commit = refresh

End Function
Public Function create(par As cDataRow, colNum As Long, rCell As Range, _
            Optional v As Variant) As cCell        ' Fill the Cell up
    
    ' if v is specifed we knw the value without needing to access the sheet
    If IsMissing(v) Then
        pValue = rCell.value
    Else
        pValue = v
    End If
    pColumn = colNum
    Set pParent = par
    Set create = Me                     ' return for convenience
End Function


Public Sub tearDown()
    ' clean up
    Set pParent = Nothing
End Sub






