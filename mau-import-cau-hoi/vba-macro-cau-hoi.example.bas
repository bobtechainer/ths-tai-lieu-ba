' Bản mẫu của vba-macro-cau-hoi.bas: đã thay mật khẩu khóa sheet bằng chữ giữ chỗ.
' Sửa DOI_MAT_KHAU_KHOA_SHEET thành đúng mật khẩu đang đặt cho sheet trong file mẫu.
' ============================================================
' VBA CODE CHO SHEET "CÂU HỎI"
' ============================================================
' CÁCH SỬ DỤNG:
' 1. Mở file .xlsx trong Excel
' 2. File > Save As > Chọn "Excel Macro-Enabled Workbook (*.xlsm)"
' 3. Nhấn Alt+F11 để mở VBA Editor
' 4. Ở cột trái, tìm và double-click vào Sheet "Câu hỏi" (Sheet2)
' 5. Paste toàn bộ code bên dưới vào cửa sổ code
' 6. Nhấn Ctrl+S để lưu, đóng VBA Editor
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    ' Chỉ xử lý khi thay đổi cột B (Loại câu hỏi), từ dòng 3 trở đi
    If Target.Column <> 2 Then Exit Sub
    If Target.Row < 3 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub
    
    Application.EnableEvents = False
    On Error GoTo CleanUp
    
    Dim r As Long
    r = Target.Row
    Dim qt As String
    qt = Trim(CStr(Target.Value))
    
    ' Range đáp án A-D (cột F-I) và Đáp án đúng (cột J)
    Dim rngA As Range: Set rngA = Me.Range(Me.Cells(r, 6), Me.Cells(r, 9))
    Dim rngJ As Range: Set rngJ = Me.Cells(r, 10)
    
    ' Tạm bỏ sheet protection
    On Error Resume Next
    Me.Unprotect Password:="DOI_MAT_KHAU_KHOA_SHEET"
    On Error GoTo CleanUp
    
    ' Reset trạng thái cũ
    rngA.Value = ""
    rngA.Locked = False
    rngA.Interior.ColorIndex = xlNone
    rngA.Font.ColorIndex = xlAutomatic
    On Error Resume Next
    rngJ.Validation.Delete
    On Error GoTo CleanUp
    
    ' Kiểm tra loại nhận định bằng cách tìm pattern trong chuỗi
    Dim isBinary As Boolean
    Dim isTernary As Boolean
    isBinary = (InStr(1, qt, "/S)") > 0) And (InStr(1, qt, "KX") = 0)
    isTernary = (InStr(1, qt, "KX") > 0)
    
    If isBinary Then
        ' Nhận định (Đ/S) - 2 lựa chọn
        Me.Cells(r, 6).Value = ChrW(272) & ChrW(250) & "ng"  ' Đúng
        Me.Cells(r, 7).Value = "Sai"
        rngA.Locked = True
        rngA.Interior.Color = RGB(235, 240, 248)  ' Light blue
        rngA.Font.Color = RGB(100, 100, 100)       ' Gray text
        ' Dropdown cho đáp án đúng: chỉ A hoặc B
        rngJ.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="A,B"
        rngJ.Validation.ShowInput = True
        
    ElseIf isTernary Then
        ' Nhận định (Đ/S/KXĐ) - 3 lựa chọn
        Me.Cells(r, 6).Value = ChrW(272) & ChrW(250) & "ng"  ' Đúng
        Me.Cells(r, 7).Value = "Sai"
        Me.Cells(r, 8).Value = "Kh" & ChrW(244) & "ng x" & ChrW(225) & "c " & ChrW(273) & ChrW(7883) & "nh"  ' Không xác định
        rngA.Locked = True
        rngA.Interior.Color = RGB(235, 240, 248)
        rngA.Font.Color = RGB(100, 100, 100)
        ' Dropdown cho đáp án đúng: A, B hoặc C
        rngJ.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="A,B,C"
        rngJ.Validation.ShowInput = True
    End If
    
    ' Bật lại sheet protection
    Me.Protect Password:="DOI_MAT_KHAU_KHOA_SHEET", UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, AllowInsertingRows:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True

CleanUp:
    Application.EnableEvents = True
End Sub

' Tự động bật sheet protection khi mở sheet
Private Sub Worksheet_Activate()
    On Error Resume Next
    Me.Protect Password:="DOI_MAT_KHAU_KHOA_SHEET", UserInterfaceOnly:=True, _
        AllowFormattingCells:=True, AllowInsertingRows:=True, _
        AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
End Sub
