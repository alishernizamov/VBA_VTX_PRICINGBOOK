Private Sub TextBox1_Change()

Dim xStr As String, xName As String
Dim v As Variant, vv As Variant
Dim xWS As Worksheet
Dim xRg As Range
    On Error GoTo Err01
    Application.ScreenUpdating = False
    xName = "Materials"
    xStr = TextBox1.Text
    Set xWS = ActiveSheet
    Set xRg = xWS.ListObjects(xName).Range
    If xStr <> "" Then
        v = Split(Application.Trim("*" & Replace(xStr, ",", "*,*") & "*"), ",")
        If IsArray(v) Then
            Select Case UBound(v)
            Case 0
                xRg.AutoFilter Field:=2, Criteria1:="*" & v(0) & "*", Operator:=xlFilterValues
            Case 1
                xRg.AutoFilter Field:=2, Criteria1:="*" & v(0) & "*", Criteria2:="*" & v(1) & "*", Operator:=xlFilterValues
            Case Else
                For Each vv In v
                    xRg.AutoFilter Field:=2, Criteria1:="*" & vv & "*", Operator:=xlFilterValues
                Next
            End Select
        Else
            xRg.AutoFilter Field:=2, Criteria1:="*" & v & "*", Operator:=xlFilterValues
        End If
    Else
        xRg.AutoFilter Field:=1, Operator:=xlFilterValues
    End If
Err01:

End Sub
