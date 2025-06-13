Attribute VB_Name = "Module1"

Sub ToiUuOng_TuDong_NhieuOng()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim colTenOng As String
    Dim colKichThuoc As String
    Dim colChieuDai As String

    colTenOng = InputBox("Cot chua ten bo cap (ex: A):")
    colKichThuoc = InputBox("Cot chua duong kinh ong (ex: C):")
    colChieuDai = InputBox("Cot chua chieu dai thiet ke (ex: D):")

    If colTenOng = "" Or colKichThuoc = "" Or colChieuDai = "" Then
        MsgBox "Thieu thong tin", vbExclamation
        Exit Sub
    End If

    Dim fullPipe As Double: fullPipe = 11.8
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, colChieuDai).End(xlUp).Row
    Dim dictDu As Object: Set dictDu = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow
        Dim chieuDaiYeuCau As Double
        chieuDaiYeuCau = ws.Cells(i, colChieuDai).Value

        Dim kichThuoc As String
        kichThuoc = Trim(ws.Cells(i, colKichThuoc).Value)

        Dim tenOng As String
        tenOng = ws.Cells(i, colTenOng).Value

        Dim soOngCan As Long
        soOngCan = WorksheetFunction.Ceiling(chieuDaiYeuCau / fullPipe, 1)
        ws.Cells(i, "E").Value = soOngCan

        Dim chieuDaiTruocCuoi As Double
        chieuDaiTruocCuoi = (soOngCan - 1) * fullPipe
        ws.Cells(i, "F").Value = chieuDaiTruocCuoi

        Dim phanThieu As Double
        phanThieu = chieuDaiYeuCau - chieuDaiTruocCuoi
        ws.Cells(i, "G").Value = Round(phanThieu, 2)

        Dim usedDu As Boolean: usedDu = False
        Dim arrDu, j As Long

        If dictDu.exists(kichThuoc) Then
            arrDu = dictDu(kichThuoc)
            For j = LBound(arrDu) To UBound(arrDu)
                If arrDu(j)(0) >= phanThieu Then
                    ' D?ng ph?n du
                    ws.Cells(i, "H").Value = arrDu(j)(0)
                    ws.Cells(i, "I").Value = Round(arrDu(j)(0) - phanThieu, 2)
                    ws.Cells(i, "J").Value = arrDu(j)(1)

                    ' C?p nh?t ph?n du c?n l?i
                    arrDu(j)(0) = arrDu(j)(0) - phanThieu
                    dictDu(kichThuoc) = arrDu
                    usedDu = True

                    ' ? Gi?m s? lu?ng ?ng n?u d?ng ph?n du
                    ws.Cells(i, "E").Value = soOngCan - 1
                    Exit For
                End If
            Next j
        End If

        If Not usedDu Then
            ' D?ng th?m 1 c?y ?ng m?i
            ws.Cells(i, "H").Value = fullPipe
            Dim duMoi As Double: duMoi = fullPipe - phanThieu
            ws.Cells(i, "I").Value = Round(duMoi, 2)
            ws.Cells(i, "J").Value = "new"

            ' Th?m ph?n du v?o danh s?ch n?u c?
            If duMoi > 0.01 Then
                Dim tempArr
                If dictDu.exists(kichThuoc) Then
                    tempArr = dictDu(kichThuoc)
                    ReDim Preserve tempArr(UBound(tempArr) + 1)
                    tempArr(UBound(tempArr)) = Array(duMoi, tenOng)
                    dictDu(kichThuoc) = tempArr
                Else
                    dictDu.Add kichThuoc, Array(Array(duMoi, tenOng))
                End If
            End If
        End If
    Next i

    MsgBox "done", vbInformation

End Sub


