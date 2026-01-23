Option Explicit

' =============================================================================
' 太陽光発電データ自動ダウンロードモジュール
' サイト: ららぽーとTOKYO-BAY第10駐車場屋上太陽光発電所
' =============================================================================

' エンドポイントURL（定数）
Private Const ENDPOINT_URL As String = "https://services.energymntr.com/megasolar/TQL0228280/watch_dl_data/download.php"
Private Const LOGIN_URL As String = "https://laplaceid.energymntr.com/"
Private Const USER_ID As String = "KPG0228280"

' クッキー保存先
Private Const COOKIE_FILE_PATH As String = "\data\巡回data.xlsx"
Private Const COOKIE_SHEET_NAME As String = "Sheet1"
Private Const COOKIE_CELL As String = "Z1"
Private Const COOKIE_DATE_CELL As String = "Z2"

' =============================================================================
' 【要件1】日次月報（PCS）ダウンロード
' 用途: 1日単位の月報データをPCSから取得
' セル参照: C6（C6の年月を使用）
' ファイル名: 月報_{年}_{月}.csv
' =============================================================================
Sub Download_DailyMonthly()
    On Error GoTo ErrorHandler

    ' アクティブシートのC6から年月を取得
    Dim targetDate As Date
    Dim targetYear As Integer
    Dim targetMonth As Integer

    If Not IsDate(ActiveSheet.Range("C6").Value) Then
        MsgBox "エラー: アクティブシートのC6セルに正しい日付を入力してください。", vbCritical
        Exit Sub
    End If

    targetDate = ActiveSheet.Range("C6").Value
    targetYear = Year(targetDate)
    targetMonth = Month(targetDate)

    ' パラメータ組立
    Dim params As String
    params = "viewType=download" & _
             "&gid=1" & _
             "&nid=-1" & _
             "&unit=day" & _
             "&range=monthly" & _
             "&type=pcs" & _
             "&startYear=" & targetYear & _
             "&startMonth=" & targetMonth

    ' ファイル名生成（ゼロ埋め2桁）
    Dim fileName As String
    fileName = "月報_" & targetYear & "_" & Format(targetMonth, "00") & ".csv"

    ' ダウンロード実行
    Dim filePath As String
    filePath = DownloadCSV(params, fileName)

    If filePath <> "" Then
        MsgBox "✅ ダウンロード完了：" & vbCrLf & filePath, vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました（Download_DailyMonthly）：" & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' 【要件2】前日の1分間隔日報（PCS）ダウンロード
' 用途: C6セルの前日を、1分間隔で日報取得
' セル参照: C6=Excelシリアル値（日付）
' ファイル名: 日報_1分_{年}_{月}_{日}.csv
' =============================================================================
Sub Download_MinuteDaily()
    On Error GoTo ErrorHandler

    ' アクティブシートのC6から前日の日付を計算
    Dim targetDate As Date
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer

    If Not IsDate(ActiveSheet.Range("C6").Value) Then
        MsgBox "エラー: アクティブシートのC6セルに正しい日付を入力してください。", vbCritical
        Exit Sub
    End If

    targetDate = ActiveSheet.Range("C6").Value - 1  ' 前日
    targetYear = Year(targetDate)
    targetMonth = Month(targetDate)
    targetDay = Day(targetDate)

    ' パラメータ組立
    Dim params As String
    params = "viewType=download" & _
             "&gid=1" & _
             "&nid=-1" & _
             "&unit=minute" & _
             "&range=daily" & _
             "&type=pcs" & _
             "&startYear=" & targetYear & _
             "&startMonth=" & targetMonth & _
             "&startDay=" & targetDay

    ' ファイル名生成（ゼロ埋め2桁）
    Dim fileName As String
    fileName = "日報_1分_" & targetYear & "_" & Format(targetMonth, "00") & "_" & Format(targetDay, "00") & ".csv"

    ' ダウンロード実行
    Dim filePath As String
    filePath = DownloadCSV(params, fileName)

    If filePath <> "" Then
        MsgBox "✅ ダウンロード完了：" & vbCrLf & filePath, vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました（Download_MinuteDaily）：" & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' 【要件3】前日の1時間間隔日報（PCS）ダウンロード
' 用途: C6セルの前日を、1時間間隔で日報取得
' セル参照: C6=Excelシリアル値（日付）
' ファイル名: 日報_1時間_{年}_{月}_{日}.csv
' =============================================================================
Sub Download_HourlyDaily()
    On Error GoTo ErrorHandler

    ' アクティブシートのC6から前日の日付を計算
    Dim targetDate As Date
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer

    If Not IsDate(ActiveSheet.Range("C6").Value) Then
        MsgBox "エラー: アクティブシートのC6セルに正しい日付を入力してください。", vbCritical
        Exit Sub
    End If

    targetDate = ActiveSheet.Range("C6").Value - 1  ' 前日
    targetYear = Year(targetDate)
    targetMonth = Month(targetDate)
    targetDay = Day(targetDate)

    ' パラメータ組立
    Dim params As String
    params = "viewType=download" & _
             "&gid=1" & _
             "&nid=-1" & _
             "&unit=hour" & _
             "&range=daily" & _
             "&type=pcs" & _
             "&startYear=" & targetYear & _
             "&startMonth=" & targetMonth & _
             "&startDay=" & targetDay

    ' ファイル名生成（ゼロ埋め2桁）
    Dim fileName As String
    fileName = "日報_1時間_" & targetYear & "_" & Format(targetMonth, "00") & "_" & Format(targetDay, "00") & ".csv"

    ' ダウンロード実行
    Dim filePath As String
    filePath = DownloadCSV(params, fileName)

    If filePath <> "" Then
        MsgBox "✅ ダウンロード完了：" & vbCrLf & filePath, vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました（Download_HourlyDaily）：" & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' 【統合版】3ファイル自動ダウンロード + 既存Import処理の呼び出し
' 用途: 手動ダウンロード→手動ファイル選択の作業を完全自動化
' セル参照: C6（すべてC6の日付を基準）
' 処理順序:
'   1. C6の前日の1分間隔日報ダウンロード
'   2. C6の前日の1時間間隔日報ダウンロード
'   3. C6の年月の月報ダウンロード
'   4. ImportDailyCsv_BB_BC_BD_ToDB を呼び出し（※既存プロシージャが必要）
' =============================================================================
Sub AutoDownloadAndImport()
    On Error GoTo ErrorHandler

    Dim startTime As Double
    startTime = Timer

    ' 進行状況メッセージ
    Application.StatusBar = "データダウンロード開始..."

    ' ================================================
    ' ステップ1: 前日の1分間隔日報ダウンロード
    ' ================================================
    Application.StatusBar = "1/3: 1分間隔日報ダウンロード中..."

    Dim targetDate As Date
    Dim targetYear As Integer
    Dim targetMonth As Integer
    Dim targetDay As Integer

    If Not IsDate(ActiveSheet.Range("C6").Value) Then
        MsgBox "エラー: アクティブシートのC6セルに正しい日付を入力してください。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    targetDate = ActiveSheet.Range("C6").Value - 1
    targetYear = Year(targetDate)
    targetMonth = Month(targetDate)
    targetDay = Day(targetDate)

    ' 1分間隔
    Dim params1 As String
    params1 = "viewType=download" & _
              "&gid=1" & _
              "&nid=-1" & _
              "&unit=minute" & _
              "&range=daily" & _
              "&type=pcs" & _
              "&startYear=" & targetYear & _
              "&startMonth=" & targetMonth & _
              "&startDay=" & targetDay

    Dim fileName1 As String
    fileName1 = "日報_1分_" & targetYear & "_" & Format(targetMonth, "00") & "_" & Format(targetDay, "00") & ".csv"

    Dim filePath1 As String
    filePath1 = DownloadCSV(params1, fileName1)

    If filePath1 = "" Then
        MsgBox "1分間隔日報のダウンロードに失敗しました。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' ================================================
    ' ステップ2: 前日の1時間間隔日報ダウンロード
    ' ================================================
    Application.StatusBar = "2/3: 1時間間隔日報ダウンロード中..."

    ' 1時間間隔
    Dim params2 As String
    params2 = "viewType=download" & _
              "&gid=1" & _
              "&nid=-1" & _
              "&unit=hour" & _
              "&range=daily" & _
              "&type=pcs" & _
              "&startYear=" & targetYear & _
              "&startMonth=" & targetMonth & _
              "&startDay=" & targetDay

    Dim fileName2 As String
    fileName2 = "日報_1時間_" & targetYear & "_" & Format(targetMonth, "00") & "_" & Format(targetDay, "00") & ".csv"

    Dim filePath2 As String
    filePath2 = DownloadCSV(params2, fileName2)

    If filePath2 = "" Then
        MsgBox "1時間間隔日報のダウンロードに失敗しました。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' ================================================
    ' ステップ3: 月報ダウンロード（C6の年月を使用）
    ' ================================================
    Application.StatusBar = "3/3: 月報ダウンロード中..."

    Dim reportYear As Integer
    Dim reportMonth As Integer

    ' C6の日付から年月を取得
    reportYear = Year(ActiveSheet.Range("C6").Value)
    reportMonth = Month(ActiveSheet.Range("C6").Value)

    Dim params3 As String
    params3 = "viewType=download" & _
              "&gid=1" & _
              "&nid=-1" & _
              "&unit=day" & _
              "&range=monthly" & _
              "&type=pcs" & _
              "&startYear=" & reportYear & _
              "&startMonth=" & reportMonth

    Dim fileName3 As String
    fileName3 = "月報_" & reportYear & "_" & Format(reportMonth, "00") & ".csv"

    Dim filePath3 As String
    filePath3 = DownloadCSV(params3, fileName3)

    If filePath3 = "" Then
        MsgBox "月報のダウンロードに失敗しました。", vbCritical
        Application.StatusBar = False
        Exit Sub
    End If

    ' ================================================
    ' ステップ4: 既存のImport処理を呼び出し
    ' ================================================
    Application.StatusBar = "データインポート中..."

    ' ※既存のImportDailyCsv_BB_BC_BD_ToDBプロシージャを呼び出す
    ' ※このプロシージャが存在しない場合は、以下の行をコメントアウトしてください
    On Error Resume Next
    Application.Run "ImportDailyCsv_BB_BC_BD_ToDB", filePath1, filePath2, filePath3
    If Err.Number <> 0 Then
        MsgBox "注意: ImportDailyCsv_BB_BC_BD_ToDB プロシージャが見つかりません。" & vbCrLf & _
               "ダウンロードのみ完了しました。", vbExclamation
        On Error GoTo ErrorHandler
    Else
        On Error GoTo ErrorHandler
    End If

    ' 完了メッセージ
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime

    Application.StatusBar = False
    MsgBox "✅ すべての処理が完了しました！" & vbCrLf & vbCrLf & _
           "ダウンロードファイル:" & vbCrLf & _
           "  - " & fileName1 & vbCrLf & _
           "  - " & fileName2 & vbCrLf & _
           "  - " & fileName3 & vbCrLf & vbCrLf & _
           "処理時間: " & Format(elapsedTime, "0.0") & "秒", vbInformation

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "エラーが発生しました（AutoDownloadAndImport）：" & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' 【クッキー管理】Laplace ID ログイン
' 機能: ログインしてセッションクッキーを取得
' 戻り値: クッキー文字列（例: "PHPSESSID=abc123..."）
' =============================================================================
Private Function LoginAndGetCookie() As String
    On Error GoTo ErrorHandler

    ' パスワードを環境変数から取得
    Dim password As String
    password = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%SOLAR_PASSWORD%")

    If password = "%SOLAR_PASSWORD%" Or password = "" Then
        MsgBox "エラー: 環境変数 SOLAR_PASSWORD が設定されていません。", vbCritical
        LoginAndGetCookie = ""
        Exit Function
    End If

    ' ログインリクエスト
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    ' コールバックURL
    Dim callback As String
    callback = "https://services.energymntr.com/megasolar/TQL0228280/login/index.php"

    ' ログインURL組み立て（GETパラメータ）
    Dim loginUrl As String
    loginUrl = LOGIN_URL & "?callback=" & callback & _
               "&idtext=" & USER_ID & _
               "&pwtext=" & password

    ' GETリクエスト送信
    http.Open "GET", loginUrl, False
    http.send

    ' ステータス確認
    If http.Status <> 200 Then
        MsgBox "ログインエラー: ステータスコード " & http.Status, vbCritical
        LoginAndGetCookie = ""
        Exit Function
    End If

    ' Set-Cookie ヘッダーからクッキー抽出
    Dim allHeaders As String
    allHeaders = http.getAllResponseHeaders()

    Dim cookieValue As String
    cookieValue = ExtractCookieFromHeaders(allHeaders)

    If cookieValue = "" Then
        MsgBox "ログインエラー: クッキーを取得できませんでした。", vbCritical
        LoginAndGetCookie = ""
        Exit Function
    End If

    LoginAndGetCookie = cookieValue
    Exit Function

ErrorHandler:
    MsgBox "ログインエラー：" & vbCrLf & Err.Description, vbCritical
    LoginAndGetCookie = ""
End Function

' =============================================================================
' 【クッキー管理】レスポンスヘッダーからクッキー抽出
' =============================================================================
Private Function ExtractCookieFromHeaders(ByVal headers As String) As String
    On Error Resume Next

    Dim lines() As String
    lines = Split(headers, vbCrLf)

    Dim line As Variant
    For Each line In lines
        If InStr(1, line, "Set-Cookie:", vbTextCompare) > 0 Then
            ' "Set-Cookie: PHPSESSID=abc123; path=/; ..." から抽出
            Dim cookiePart As String
            cookiePart = Mid(line, InStr(line, ":") + 2)

            ' セミコロンまでを取得
            If InStr(cookiePart, ";") > 0 Then
                cookiePart = Left(cookiePart, InStr(cookiePart, ";") - 1)
            End If

            ExtractCookieFromHeaders = Trim(cookiePart)
            Exit Function
        End If
    Next line

    ExtractCookieFromHeaders = ""
End Function

' =============================================================================
' 【クッキー管理】巡回data.xlsxからクッキー読み込み
' =============================================================================
Private Function LoadCookie() As String
    On Error GoTo ErrorHandler

    Dim cookieFilePath As String
    cookieFilePath = ThisWorkbook.Path & COOKIE_FILE_PATH

    ' ファイル存在チェック
    If Dir(cookieFilePath) = "" Then
        LoadCookie = ""
        Exit Function
    End If

    ' Excelファイルを開く
    Dim wb As Workbook
    Set wb = Workbooks.Open(cookieFilePath, ReadOnly:=True)

    ' クッキー読み込み
    Dim cookieValue As String
    cookieValue = wb.Sheets(COOKIE_SHEET_NAME).Range(COOKIE_CELL).Value

    wb.Close SaveChanges:=False

    LoadCookie = cookieValue
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    LoadCookie = ""
End Function

' =============================================================================
' 【クッキー管理】巡回data.xlsxにクッキー保存
' =============================================================================
Private Sub SaveCookie(ByVal cookieValue As String)
    On Error GoTo ErrorHandler

    Dim cookieFilePath As String
    cookieFilePath = ThisWorkbook.Path & COOKIE_FILE_PATH

    ' dataフォルダ作成（存在しない場合）
    Dim dataFolder As String
    dataFolder = ThisWorkbook.Path & "\data"
    If Dir(dataFolder, vbDirectory) = "" Then
        MkDir dataFolder
    End If

    Dim wb As Workbook
    Dim fileExists As Boolean
    fileExists = (Dir(cookieFilePath) <> "")

    If fileExists Then
        ' 既存ファイルを開く
        Set wb = Workbooks.Open(cookieFilePath)
    Else
        ' 新規ファイル作成
        Set wb = Workbooks.Add
    End If

    ' クッキーと日時を保存
    wb.Sheets(COOKIE_SHEET_NAME).Range(COOKIE_CELL).Value = cookieValue
    wb.Sheets(COOKIE_SHEET_NAME).Range(COOKIE_DATE_CELL).Value = Now

    ' 保存
    If fileExists Then
        wb.Close SaveChanges:=True
    Else
        wb.SaveAs cookieFilePath, FileFormat:=51  ' xlOpenXMLWorkbook
        wb.Close SaveChanges:=False
    End If

    Exit Sub

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "クッキー保存エラー：" & vbCrLf & Err.Description, vbCritical
End Sub

' =============================================================================
' 【共通処理】CSV ダウンロード関数（クッキー管理付き）
' 機能: HTTPリクエストを送信してCSVファイルを保存
' 引数:
'   postParams - POSTパラメータ（URLエンコード済み文字列）
'   fileName - 保存するファイル名（例: "月報_2026_01.csv"）
' 戻り値:
'   成功時: 保存したファイルのフルパス
'   失敗時: 空文字列
' =============================================================================
Private Function DownloadCSV(ByVal postParams As String, ByVal fileName As String) As String
    On Error GoTo ErrorHandler

    ' 1. 保存されたクッキーを読み込み
    Dim cookie As String
    cookie = LoadCookie()

    ' クッキーがない場合はログイン
    If cookie = "" Then
        cookie = LoginAndGetCookie()
        If cookie = "" Then
            MsgBox "ログインに失敗しました。", vbCritical
            DownloadCSV = ""
            Exit Function
        End If
        SaveCookie cookie
    End If

    ' 2. クッキー付きでダウンロード試行
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    http.Open "POST", ENDPOINT_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Cookie", cookie
    http.send postParams

    ' 3. ステータス確認（403の場合は再ログイン）
    If http.Status = 403 Then
        ' クッキー無効 → 再ログイン
        cookie = LoginAndGetCookie()
        If cookie = "" Then
            MsgBox "再ログインに失敗しました。", vbCritical
            DownloadCSV = ""
            Exit Function
        End If
        SaveCookie cookie

        ' 再度ダウンロード試行
        Set http = CreateObject("MSXML2.XMLHTTP.6.0")
        http.Open "POST", ENDPOINT_URL, False
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        http.setRequestHeader "Cookie", cookie
        http.send postParams
    End If

    ' ステータスコード確認
    If http.Status <> 200 Then
        MsgBox "HTTPエラー: ステータスコード " & http.Status & vbCrLf & _
               "URL: " & ENDPOINT_URL, vbCritical
        DownloadCSV = ""
        Exit Function
    End If

    ' レスポンスボディをバイナリで取得
    Dim responseBody As Variant
    responseBody = http.responseBody

    ' ファイル保存先パス生成（Excelファイルと同じフォルダ）
    Dim savePath As String
    savePath = ThisWorkbook.Path & "\" & fileName

    ' ADODB.Stream でバイナリデータをCSVファイルに書き込み
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")

    stream.Type = 1  ' adTypeBinary
    stream.Open
    stream.Write responseBody
    stream.SaveToFile savePath, 2  ' adSaveCreateOverWrite
    stream.Close

    ' 成功時はフルパスを返す
    DownloadCSV = savePath

    Exit Function

ErrorHandler:
    MsgBox "ファイル保存エラー：" & vbCrLf & Err.Description & vbCrLf & _
           "保存先: " & savePath, vbCritical
    DownloadCSV = ""
End Function
