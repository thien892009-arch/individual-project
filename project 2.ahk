#Requires AutoHotkey v1.1
#NoEnv
#SingleInstance Force
SetWorkingDir %A_ScriptDir%


; ==== Biến toàn cục ====
global pageCount := 1
global currentPage := 1
global lineInPage := 0
global maxLinesPerPage := 10
global data := [] ; Mảng lưu tất cả dữ liệu
global discordWebhook := "https://discord.com/api/webhooks/1414626399683219456/Tc0D7JZfqwx2bOU-IkR95i993GNOc6nu_IooAQp2kS8CahBdMmJBIjtZeLYqT7YA-vrV"
global csvFile := "study.csv"

; ==== Kích thước GUI ====
h := 700
w := 1000
h1 := h - 100
w1 := w - 30

; ==== Đọc dữ liệu từ CSV nếu có ====
LoadCSV()

; ==== Tự động cập nhật Next Review nếu cần ====
AutoUpdateReview()

; ==== Khởi tạo Page1 và hiển thị dữ liệu ====
CreatePage(1)
ShowData()
SendWebhook("ACTIVATED SUCCESSFULLY")
return

; ==== Hàm gửi webhook ====
SendWebhook(msg)
{
    global discordWebhook
    FormatTime, messageTime,, yyyy-MM-dd hh:mm:ss
    fullMessage := "[" . messageTime . "] " . msg

    ; Escape các dấu " và xuống dòng
    StringReplace, fullMessage, fullMessage, ", `\", All
    StringReplace, fullMessage, fullMessage, `n, \n, All

    ; Tạo JSON chuẩn
    json := "{""content"":""" fullMessage """}"

    whr := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    try {
        whr.Open("POST", discordWebhook, false)
        whr.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
        whr.Send(json)
        whr.WaitForResponse()
        status := whr.Status
        ; MsgBox % "Webhook status: " status "`nResponse: " whr.ResponseText
    } catch {
        MsgBox % "⚠ Webhook gửi thất bại!"
    }
}
; ==== Hàm tạo trang mới ====
CreatePage(pageNum)
{
    global w1, h1, w, h
    Gui, %pageNum%:New
    Gui, %pageNum%:Font, s10 Bold cWhite, Times New Roman
    Gui, %pageNum%:Color, 000000
    Gui, %pageNum%:Add, Tab2, w%w1% h%h1%, Page%pageNum%
    ; Tiêu đề các cột
    Gui, %pageNum%:Add, Text, x30 y40 cYellow, STT
    Gui, %pageNum%:Add, Text, x80 y40 cYellow, Môn học
    Gui, %pageNum%:Add, Text, x170 y40 cYellow, Tên bài học
    Gui, %pageNum%:Add, Text, x400 y40 cYellow, Ngày bắt đầu
    Gui, %pageNum%:Add, Text, x550 y40 cYellow, Last Review
    Gui, %pageNum%:Add, Text, x700 y40 cYellow, Next Review
    Gui, %pageNum%:Add, Text, x850 y40 cYellow, Ghi chú
    ; Nút điều hướng
    Gui, %pageNum%:Add, Button, gPrevPage x20 y560 w80 h30, Prev
    Gui, %pageNum%:Add, Button, gNextPage x890 y560 w80 h30, Next
    Gui, %pageNum%:Add, Button, gUpdate1 x850 y650 w50 h30 , Add
    Gui, %pageNum%:Add, Button, gDelete1 x910 y650 w50 h30 , Delete
    ; Ô nhập dữ liệu
    Gui, %pageNum%:Add, Text, x20 y620 cWhite, Nhập Dữ Liệu:
    Gui, %pageNum%:Add, Edit, x50 y650 w100 h30 CBlack vMyEdit +0x1, Môn học
    Gui, %pageNum%:Add, Edit, x160 y650 w150 h30 CBlack vMyLesson +0x1, Tên bài
    Gui, %pageNum%:Add, Edit, x320 y650 w150 h30 CBlack vMyDate +0x1, %A_YYYY%-%A_MM%-%A_DD%
    Gui, %pageNum%:Add, Edit, x480 y650 w150 h30 CBlack vMyLast +0x1, %A_YYYY%-%A_MM%-%A_DD%
    Gui, %pageNum%:Add, Edit, x640 y650 w200 h30 CBlack vMyEdit3 +0x1, Ghi chú
    Gui, %pageNum%:Show, w%w% h%h%, Page%pageNum%
}

; ==== Hiển thị dữ liệu từ mảng lên GUI ====
ShowData()
{
    global data, maxLinesPerPage, pageCount, currentPage, lineInPage
    lineInPage := 0
    pageCount := Ceil(data.MaxIndex() / maxLinesPerPage)
    if (pageCount < 1)
        pageCount := 1
    idx := 1
    Loop, %pageCount%
    {
        p := A_Index
        CreatePage(p)
        startIdx := (p - 1) * maxLinesPerPage + 1
        endIdx := Min(startIdx + maxLinesPerPage - 1, data.MaxIndex())
        yPos := 70
        Loop, % (endIdx - startIdx + 1)
        {
            item := data[startIdx + A_Index - 1]
            Gui, %p%:Add, Text, x42 y%yPos% cWhite, % (startIdx + A_Index - 1)
            Gui, %p%:Add, Text, x80 y%yPos% cWhite, % item[1]
            Gui, %p%:Add, Text, x170 y%yPos% cWhite, % item[2]
            Gui, %p%:Add, Text, x400 y%yPos% cWhite, % item[3]
            Gui, %p%:Add, Text, x550 y%yPos% cWhite, % item[4]
            Gui, %p%:Add, Text, x700 y%yPos% cWhite, % item[5]
            Gui, %p%:Add, Text, x850 y%yPos% cWhite, % item[6]
            yPos += 30
        }
    }
}

; ==== Nút Add ====
Update1:
Gui, Submit, NoHide
FormatTime, nextReview, % MyDateAdd(MyLast, 2, "days"), yyyy-MM-dd
data.Push([MyEdit, MyLesson, MyDate, MyLast, nextReview, MyEdit3, 0])
SaveCSV()
msg := "➕ Thêm bài học mới:`n"
msg .= "📚 Môn: " MyEdit "`n"
msg .= "📖 Bài: " MyLesson "`n"
msg .= "📅 Ngày bắt đầu: " MyDate "`n"
msg .= "⏳ Next Review: " nextReview "`n"
if (MyEdit3 != "")
    msg .= "📝 Ghi chú: " MyEdit3 "`n"
msg .= "________END________"
SendWebhook(msg)
Gui, Destroy
ShowData()
currentPage := pageCount
Gui, %currentPage%:Show
return

; ==== Nút Delete ====
Delete1:
if (data.MaxIndex() < 1)
{
    MsgBox, 48, Thông báo, Không có dữ liệu để xóa!
    return
}
MsgBox, 36, Xác nhận, Bạn có chắc muốn xóa dòng cuối cùng?
IfMsgBox, No
    return
data.RemoveAt(data.MaxIndex())
SaveCSV()
Gui, Destroy
ShowData()
if (currentPage > pageCount)
    currentPage := pageCount
Gui, %currentPage%:Show
return

; ==== Hàm cộng ngày ====
MyDateAdd(dateStr, addValue, unit)
{
    StringSplit, part, dateStr, -
    year := part1, month := part2, day := part3
    dateObj := year . month . day
    EnvAdd, dateObj, %addValue%, %unit%
    return dateObj
}

; ==== Nút Prev ====
PrevPage:
if (currentPage > 1)
{
    Gui, %currentPage%:Hide
    currentPage--
    Gui, %currentPage%:Show
}
return

; ==== Nút Next ====
NextPage:
if (currentPage < pageCount)
{
    Gui, %currentPage%:Hide
    currentPage++
    Gui, %currentPage%:Show
}
return

; ==== Load CSV ====
LoadCSV()
{
    global csvFile, data
    if !FileExist(csvFile)
        return
    FileRead, content, %csvFile%
    content := StrReplace(content, "`r") ; loại bỏ carriage return
    lines := StrSplit(content, "`n")
    For k, line in lines
    {
        if (k = 1)
            continue
        if (line = "")
            continue
        cols := StrSplit(line, ",")
        if (cols.MaxIndex() >= 6)
        {
            reviewCount := (cols.MaxIndex() >= 8) ? cols[8] : 0
            data.Push([cols[2], cols[3], cols[4], cols[5], cols[6], cols[7], reviewCount])
        }
    }
}

; ==== Save CSV (UTF-8) ====
SaveCSV()
{
    global csvFile, data
    header := "STT,Môn học,Tên bài học,Ngày bắt đầu,Last Review,Next Review,Ghi chú,ReviewCount`n"
    out := header
    For i, item in data
    {
        out .= i "," item[1] "," item[2] "," item[3] "," item[4] "," item[5] "," item[6] "," item[7] "`n"
    }
    FileDelete, %csvFile%
    FileAppend, %out%, %csvFile%, UTF-8
}

; ==== Auto Update Review ====
AutoUpdateReview()
{
    global data
    today := A_YYYY . "-" . A_MM . "-" . A_DD
    changed := false
    msg := ""
    For i, item in data
    {
        if (item[5] <= today)
        {
            item[4] := today
            count := item[7] + 1
            interval := GetInterval(count)
            FormatTime, nextReview, % MyDateAdd(today, interval, "days"), yyyy-MM-dd
            item[5] := nextReview
            item[7] := count
            msg .= "✔ " item[1] " - " item[2] " | Lần " count " | Next: " nextReview "`n"
            changed := true
        }
    }
    if (changed)
    {
        SaveCSV()
        SendWebhook("Cập nhật Review hôm nay:`n" msg)
    }
}

; ==== Khoảng cách ====
GetInterval(count)
{
    static intervals := [3, 7, 14, 21, 30]
    if (count > intervals.MaxIndex())
        return intervals[intervals.MaxIndex()]
    return intervals[count]
}

; ==== Thoát ====
f10::ExitApp