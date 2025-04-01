Attribute VB_Name = "Module1"
Sub NhapDiemTheoID_3SoCuoi()
    Dim ws As Worksheet
    Dim idCot As String, diemCot As String
    Dim idNhap3SoCuoi As String
    Dim diemNhap As Variant
    Dim i As Long, lastRow As Long
    Dim foundRow As Long

    
    idCot = "A"
    diemCot = "E"
    Set ws = ThisWorkbook.Sheets("Sheet1")

    lastRow = ws.Cells(ws.Rows.Count, idCot).End(xlUp).Row

    Do While True
        idNhap3SoCuoi = InputBox("nhap 3 so cuoi id hsinh :3 (hoac cancel de ket thuc :>):", "nhap 3 so cuoi id >.<")

        If idNhap3SoCuoi = "" Then
            Exit Do
        End If

        
        If Not IsNumeric(idNhap3SoCuoi) Or Len(idNhap3SoCuoi) <> 3 Then
            MsgBox "nhap dung 3 so cuoi id.", vbCritical
            GoTo TiepTucNhapID
        End If

        
        foundRow = 0
        For i = 2 To lastRow
            Dim idHocSinhTrongExcel As String
            idHocSinhTrongExcel = ws.Cells(i, idCot).Value

            
            If VarType(idHocSinhTrongExcel) = vbString And Len(idHocSinhTrongExcel) >= 3 Then
                Dim baSoCuoiID_Excel As String
                baSoCuoiID_Excel = Right(idHocSinhTrongExcel, 3)

                If baSoCuoiID_Excel = idNhap3SoCuoi Then
                    foundRow = i
                    Exit For
                End If
            End If
        Next i

        If foundRow > 0 Then
            diemNhap = InputBox("nhap diem cho hsinh co id: " & idNhap3SoCuoi, "Nh?p Ði?m")

            If diemNhap <> "" Then
                If IsNumeric(diemNhap) Then
                    ws.Cells(foundRow, diemCot).Value = CDbl(diemNhap)
                Else
                    MsgBox "gia tri nhap ko hop le", vbCritical
                End If
            End If
        Else
            MsgBox "khong tim the hsinh id " & idNhap3SoCuoi, vbCritical
        End If

TiepTucNhapID:
    Loop

    MsgBox "nhap xonggg!", vbInformation
End Sub
