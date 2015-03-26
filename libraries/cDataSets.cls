'gistThat@mcpher.com :do not modify this line - see ramblings.mcpher.com for details: updated on 8/18/2014 4:47:41 PM : from manifest:3414394 gist https://gist.github.com/brucemcpherson/3414216/raw/cDataSets.cls
Option Explicit
' v2.01
'for more about this
' http://ramblings.mcpher.com/Home/excelquirks/classeslink/data-manipulation-classes
'to contact me
' http://groups.google.com/group/excel-ramblings
'reuse of code
' http://ramblings.mcpher.com/Home/excelquirks/codeuse
' CdataSets
Private pCollect As Collection
Private pName As String
Public Property Get dataSets() As Collection
    Set dataSets = pCollect
End Property
Public Property Get dataSet(sn As String, Optional complain As Boolean = False) As cDataSet
    Dim ds As cDataSet
    Set ds = exists(sn)
    If ds Is Nothing Then
        If complain Then MsgBox ("data set " & sn & " doesnt exist")
    End If
    Set dataSet = ds
End Property

Public Property Get name() As String
    name = pName
End Property
Public Function create(Optional sName As String = "DataSets") As cDataSets
    pName = sName

    Set create = Me
End Function
Public Function init(Optional rInput As Range = Nothing, Optional keepFresh As Boolean = False, _
                Optional sn As String = vbNullString, _
                Optional blab As Boolean = False, Optional blockstarts As String, _
                Optional bLikely As Boolean = False, _
                Optional sKey As String = vbNullString, _
                Optional respectFilter As Boolean = False) As cDataSet
    Dim ds As cDataSet
    Set ds = New cDataSet
    With ds
        .populateData rInput, keepFresh, sn, blab, blockstarts, Me, bLikely, sKey, , , respectFilter
    End With
    
    ''pCollect.add ds, ds.name
    Set init = ds
End Function
Private Function exists(sid As Variant) As cDataSet
    On Error GoTo handle
    Set exists = pCollect(sid)
    Exit Function
handle:
    Set exists = Nothing
End Function

Public Sub tearDown()
    ' clean up
    Dim ds As cDataSet
    If Not pCollect Is Nothing Then
        For Each ds In dataSets
            ds.tearDown
        Next ds
        Set pCollect = Nothing
    End If
End Sub

Private Sub Class_Initialize()
    Set pCollect = New Collection
End Sub






