Option Explicit

' =============================================================================
' 太陽光発電データ自動ダウンロードモジュール
' サイト: ららぽーとTOKYO-BAY第10駐車場屋上太陽光発電所
' =============================================================================

' エンドポイントURL（定数）
Private Const ENDPOINT_URL As String = "https://services.energymntr.com/megasolar/TQL0228280/watch_dl_data/download.php"

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
' 【共通処理】CSV ダウンロード関数
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

    ' HTTP リクエストオブジェクト生成
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    ' POST リクエスト送信
    http.Open "POST", ENDPOINT_URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.send postParams

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
