# セットリスト集計ツール（Google Cloud Vision API版）
# Google Cloud Visionを使用した高精度OCR

<#
.SYNOPSIS
Google Cloud Vision APIを使用して画像からセットリストを抽出・集計

.DESCRIPTION
このスクリプトは以下の方法で動作します:
1. 手動でダウンロードした画像ファイルを読み込み
2. Google Cloud Vision APIでテキスト抽出（高精度）
3. セットリストをパース
4. 曲名を集計して出力

.NOTES
前提条件:
- Google Cloud アカウント
- Cloud Vision API が有効化されていること
- サービスアカウントキー（JSON）が準備されていること
#>

# ====================
# 設定セクション
# ====================

# Google Cloud Vision API 設定
$ServiceAccountKeyPath = ".\google-credentials.json"  # サービスアカウントキーのパス

# 入力ディレクトリ（画像を手動で配置）
$InputImagesDir = ".\InputImages"

# 出力ディレクトリ
$OutputDir = ".\SetlistAnalysis"
$OcrOutputDir = "$OutputDir\OCR"
$LogDir = "$OutputDir\Logs"

# Google Cloud Vision API エンドポイント
$VisionApiEndpoint = "https://vision.googleapis.com/v1/images:annotate"

# ディレクトリの作成
function Initialize-Directories {
    @($InputImagesDir, $OutputDir, $OcrOutputDir, $LogDir) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Force -Path $_ | Out-Null
            Write-Host "作成しました: $_" -ForegroundColor Green
        }
    }
}

# ====================
# ログ機能
# ====================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    $logFile = "$LogDir\setlist_$(Get-Date -Format 'yyyyMMdd').log"
    Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
    
    switch ($Level) {
        "ERROR" { Write-Host $logMessage -ForegroundColor Red }
        "WARN"  { Write-Host $logMessage -ForegroundColor Yellow }
        "INFO"  { Write-Host $logMessage -ForegroundColor Cyan }
        default { Write-Host $logMessage }
    }
}

# ====================
# Google Cloud認証
# ====================

function Get-GoogleAccessToken {
    if (-not (Test-Path $ServiceAccountKeyPath)) {
        Write-Log "サービスアカウントキーが見つかりません: $ServiceAccountKeyPath" -Level "ERROR"
        Write-Host @"

⚠ Google Cloud サービスアカウントキーが必要です！

セットアップ手順:
1. https://console.cloud.google.com/ にアクセス
2. プロジェクトを作成（または既存のプロジェクトを選択）
3. 「APIとサービス」→「ライブラリ」
4. 「Cloud Vision API」を検索して有効化
5. 「APIとサービス」→「認証情報」
6. 「認証情報を作成」→「サービスアカウント」
7. サービスアカウントを作成し、キー（JSON）をダウンロード
8. ダウンロードしたJSONファイルを「google-credentials.json」として保存

無料枠: 月1000リクエストまで無料

"@ -ForegroundColor Yellow
        return $null
    }
    
    try {
        $serviceAccount = Get-Content $ServiceAccountKeyPath -Raw | ConvertFrom-Json
        
        # JWT作成
        $now = [Math]::Floor([decimal](Get-Date(Get-Date).ToUniversalTime()-uformat "%s"))
        $exp = $now + 3600
        
        $header = @{
            alg = "RS256"
            typ = "JWT"
        } | ConvertTo-Json -Compress
        
        $payload = @{
            iss = $serviceAccount.client_email
            scope = "https://www.googleapis.com/auth/cloud-vision"
            aud = "https://oauth2.googleapis.com/token"
            exp = $exp
            iat = $now
        } | ConvertTo-Json -Compress
        
        $headerBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $payloadBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        
        $signatureInput = "$headerBase64.$payloadBase64"
        
        # RSA署名
        $privateKeyPem = $serviceAccount.private_key
        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportFromPem($privateKeyPem)
        
        $signatureBytes = $rsa.SignData(
            [System.Text.Encoding]::UTF8.GetBytes($signatureInput),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
        )
        
        $signatureBase64 = [Convert]::ToBase64String($signatureBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $jwt = "$signatureInput.$signatureBase64"
        
        # アクセストークン取得
        $tokenResponse = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body @{
            grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
            assertion = $jwt
        } -ContentType "application/x-www-form-urlencoded"
        
        Write-Log "Google Cloud認証成功"
        return $tokenResponse.access_token
    }
    catch {
        Write-Log "Google Cloud認証エラー: $_" -Level "ERROR"
        return $null
    }
}

# ====================
# Google Cloud Vision OCR
# ====================

function Extract-TextFromImage-GoogleVision {
    param(
        [string]$ImagePath,
        [string]$AccessToken
    )
    
    $fileName = Split-Path $ImagePath -Leaf
    
    try {
        Write-Log "Google Vision OCR開始: $fileName"
        
        # 画像をBase64エンコード
        $imageBytes = [System.IO.File]::ReadAllBytes($ImagePath)
        $imageBase64 = [Convert]::ToBase64String($imageBytes)
        
        # APIリクエストボディ
        $requestBody = @{
            requests = @(
                @{
                    image = @{
                        content = $imageBase64
                    }
                    features = @(
                        @{
                            type = "TEXT_DETECTION"
                            maxResults = 1
                        }
                    )
                    imageContext = @{
                        languageHints = @("ja", "en")
                    }
                }
            )
        } | ConvertTo-Json -Depth 10
        
        # API呼び出し
        $headers = @{
            "Authorization" = "Bearer $AccessToken"
            "Content-Type" = "application/json; charset=utf-8"
        }
        
        $response = Invoke-RestMethod -Uri $VisionApiEndpoint -Method Post -Headers $headers -Body $requestBody
        
        if ($response.responses -and $response.responses[0].textAnnotations) {
            $text = $response.responses[0].textAnnotations[0].description
            
            Write-Log "OCR完了: $fileName ($(($text -split "`n").Count) 行)"
            
            # デバッグ用に抽出テキストを保存
            $debugFile = "$OcrOutputDir\debug_$fileName.txt"
            $text | Out-File -FilePath $debugFile -Encoding UTF8
            
            # 完全なレスポンスも保存（デバッグ用）
            $jsonFile = "$OcrOutputDir\response_$fileName.json"
            $response | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFile -Encoding UTF8
            
            return $text
        }
        else {
            Write-Log "テキストが検出されませんでした: $fileName" -Level "WARN"
            return ""
        }
    }
    catch {
        Write-Log "Google Vision OCRエラー: $fileName - $_" -Level "ERROR"
        if ($_.ErrorDetails.Message) {
            Write-Log "詳細: $($_.ErrorDetails.Message)" -Level "ERROR"
        }
        return ""
    }
}

# ====================
# セットリスト解析
# ====================

function Parse-Setlist {
    param(
        [string]$Text,
        [string]$FileName
    )
    
    $songs = @()
    $lines = $Text -split "`r?`n"
    
    Write-Log "セットリスト解析開始: $FileName (全 $($lines.Count) 行)"
    
    foreach ($line in $lines) {
        $line = $line.Trim()
        
        # 空行や短すぎる行をスキップ
        if ($line.Length -lt 2) { continue }
        
        # ノイズ除去（日付、曜日、時刻など）
        if ($line -match '^\d{4}[/-年]\d{1,2}[/-月]\d{1,2}' -or 
            $line -match '^(月|火|水|木|金|土|日)曜日' -or
            $line -match '^\d{1,2}:\d{2}' -or
            $line -match '^(OPEN|START|開場|開演)' -or
            $line -match '@') {
            continue
        }
        
        $songName = $null
        
        # パターン1: 番号付きリスト（1. 曲名、1) 曲名、01. 曲名）
        if ($line -match '^\d+[\.\)]\s*(.+)$') {
            $songName = $Matches[1].Trim()
        }
        # パターン2: 番号とスペースのみ（1 曲名、01 曲名）
        elseif ($line -match '^0?\d+\s+(.{2,})$') {
            $songName = $Matches[1].Trim()
        }
        # パターン3: アンコール表記
        elseif ($line -match '^(EN|encore|アンコール|EC|Encore)[:：\s]+(.+)$') {
            $songName = $Matches[2].Trim()
        }
        # パターン4: 記号で始まる行
        elseif ($line -match '^[・●○◆■□▲△\-\*]\s*(.+)$') {
            $songName = $Matches[1].Trim()
        }
        # パターン5: M1, M2 などの表記
        elseif ($line -match '^M0?\d+[\.\s]+(.+)$') {
            $songName = $Matches[1].Trim()
        }
        # パターン6: 【】や［］で囲まれた後の曲名
        elseif ($line -match '^[【\[].+?[】\]]\s*(.+)$') {
            $songName = $Matches[1].Trim()
        }
        
        # 曲名の後処理
        if ($songName) {
            # 末尾の不要な文字を削除
            $songName = $songName -replace '[　\s]+$', ''
            $songName = $songName -replace '[、。，．]+$', ''
            
            # カッコ内の注釈を削除（オプション）
            # $songName = $songName -replace '\s*[\(（][^\)）]*[\)）]\s*', ''
            
            # あまりに短い or 長い曲名は除外
            if ($songName.Length -ge 2 -and $songName.Length -le 100) {
                # よくあるノイズを除外
                if ($songName -notmatch '^(setlist|SET LIST|セットリスト|会場|チケット)') {
                    $songs += $songName
                    Write-Log "  検出: $songName" -Level "INFO"
                }
            }
        }
    }
    
    Write-Log "解析完了: $FileName - $($songs.Count) 曲検出"
    return $songs
}

# ====================
# 曲名の正規化
# ====================

function Normalize-SongName {
    param([string]$SongName)
    
    # 全角・半角の統一
    $normalized = $SongName
    
    # 全角スペースを半角に
    $normalized = $normalized -replace '　', ' '
    
    # 連続するスペースを1つに
    $normalized = $normalized -replace '\s+', ' '
    
    # 前後のスペース削除
    $normalized = $normalized.Trim()
    
    # 大文字小文字の統一（オプション - 英語曲名用）
    # $normalized = $normalized.ToLower()
    
    return $normalized
}

# ====================
# 集計処理
# ====================

function Get-SongStatistics {
    param([array]$AllSongs)
    
    Write-Log "曲名の集計開始: 全 $($AllSongs.Count) 曲"
    
    $songCounts = @{}
    
    foreach ($song in $AllSongs) {
        $normalizedSong = Normalize-SongName -SongName $song
        
        if ($normalizedSong -and $normalizedSong.Length -gt 0) {
            if ($songCounts.ContainsKey($normalizedSong)) {
                $songCounts[$normalizedSong]++
            }
            else {
                $songCounts[$normalizedSong] = 1
            }
        }
    }
    
    Write-Log "集計完了: ユニークな曲数 $($songCounts.Count)"
    
    return $songCounts
}

# ====================
# 結果出力
# ====================

function Export-Results {
    param(
        [hashtable]$SongCounts,
        [int]$TotalImages
    )
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    
    $script:rank = 1
    
    # ソート済み結果の作成
    $results = $songCounts.GetEnumerator() | 
        Sort-Object Value -Descending |
        Select-Object @{Name="順位";Expression={$script:rank++}},
                      @{Name="曲名";Expression={$_.Key}}, 
                      @{Name="演奏回数";Expression={$_.Value}},
                      @{Name="出現率(%)";Expression={[math]::Round(($_.Value / $TotalImages) * 100, 1)}}
    
    # コンソール出力
    Write-Host "`n" + ("=" * 80) -ForegroundColor Cyan
    Write-Host "セットリスト集計結果" -ForegroundColor Cyan
    Write-Host ("=" * 80) -ForegroundColor Cyan
    Write-Host "分析画像数: $TotalImages 枚"
    Write-Host "検出曲数: $($results.Count) 曲"
    Write-Host "総演奏回数: $($songCounts.Values | Measure-Object -Sum | Select-Object -ExpandProperty Sum) 回"
    Write-Host ("=" * 80) -ForegroundColor Cyan
    
    $results | Format-Table -AutoSize
    
    # CSV出力
    $csvPath = "$OutputDir\setlist_results_$timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Log "CSV保存完了: $csvPath"
    Write-Host "`nCSV: $csvPath" -ForegroundColor Green
    
    # JSON出力
    $jsonPath = "$OutputDir\setlist_results_$timestamp.json"
    $jsonData = @{
        AnalysisDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalImages = $TotalImages
        TotalUniqueSongs = $results.Count
        TotalPerformances = $songCounts.Values | Measure-Object -Sum | Select-Object -ExpandProperty Sum
        Results = $results
    }
    $jsonData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
    Write-Log "JSON保存完了: $jsonPath"
    Write-Host "JSON: $jsonPath" -ForegroundColor Green
    
    # サマリーレポート
    $reportPath = "$OutputDir\summary_report_$timestamp.txt"
    $top10 = $results | Select-Object -First 10 | ForEach-Object { 
        "$($_.'順位'). $($_.'曲名') - $($_.'演奏回数')回 ($($_.'出現率(%)')%)" 
    }
    $frequent = $results | Where-Object { $_.'出現率(%)' -ge 50 } | ForEach-Object { 
        "- $($_.'曲名') ($($_.'演奏回数')回, $($_.'出現率(%)')%)" 
    }
    
    $report = @"
========================================
セットリスト集計サマリー
========================================
分析日時: $(Get-Date -Format "yyyy年MM月dd日 HH:mm:ss")
分析画像数: $TotalImages 枚
検出曲数: $($results.Count) 曲
総演奏回数: $($jsonData.TotalPerformances) 回
----------------------------------------

【トップ10】
$($top10 -join "`n")

【出現率50%以上の曲】
$($frequent -join "`n")

【出現率30%以上の曲数】
$($results | Where-Object { $_.'出現率(%)' -ge 30 } | Measure-Object | Select-Object -ExpandProperty Count) 曲

========================================
"@
    $report | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Log "レポート保存完了: $reportPath"
    Write-Host "レポート: $reportPath" -ForegroundColor Green
}

# ====================
# メイン処理
# ====================

function Main {
    Write-Host @"

╔═══════════════════════════════════════════════════════════╗
║   セットリスト集計ツール (Google Cloud Vision版)         ║
║   高精度OCRで自動的にセットリストを抽出・集計            ║
╚═══════════════════════════════════════════════════════════╝

"@ -ForegroundColor Cyan
    
    Write-Log "処理開始"
    
    # 1. 初期化
    Write-Host "[1] ディレクトリ初期化..." -ForegroundColor Yellow
    Initialize-Directories
    
    # 2. Google Cloud認証
    Write-Host "`n[2] Google Cloud認証..." -ForegroundColor Yellow
    $accessToken = Get-GoogleAccessToken
    
    if (-not $accessToken) {
        return
    }
    Write-Host "✓ Google Cloud認証成功" -ForegroundColor Green
    
    # 3. 画像ファイル確認
    Write-Host "`n[3] 画像ファイル確認..." -ForegroundColor Yellow
    $imageFiles = Get-ChildItem -Path $InputImagesDir -Include @("*.jpg", "*.jpeg", "*.png", "*.bmp", "*.gif") -Recurse
    
    if ($imageFiles.Count -eq 0) {
        Write-Host @"

⚠ 画像ファイルが見つかりません！

次の手順で画像を準備してください:
1. Twitterで @onstage_kanda のプロフィールにアクセス
2. 検索ボックスで「from:onstage_kanda since:2021-12-01 until:2021-12-31」と入力
   （年を変えて複数年分検索）
3. セットリストが含まれる画像を手動でダウンロード
4. ダウンロードした画像を「$InputImagesDir」フォルダに保存
5. このスクリプトを再実行

"@ -ForegroundColor Yellow
        return
    }
    
    Write-Host "✓ $($imageFiles.Count) 枚の画像を検出" -ForegroundColor Green
    Write-Log "検出画像数: $($imageFiles.Count)"
    
    # 無料枠チェック
    if ($imageFiles.Count -gt 1000) {
        Write-Host "`n⚠ 警告: 画像数が1000枚を超えています" -ForegroundColor Yellow
        Write-Host "Google Cloud Vision APIの無料枠（月1000リクエスト）を超える可能性があります" -ForegroundColor Yellow
        $continue = Read-Host "続行しますか？ (Y/N)"
        if ($continue -ne "Y" -and $continue -ne "y") {
            Write-Host "処理を中断しました" -ForegroundColor Yellow
            return
        }
    }
    
    # 4. OCR処理
    Write-Host "`n[4] Google Cloud Vision OCR処理実行中..." -ForegroundColor Yellow
    Write-Host "（高精度処理のため、少し時間がかかります）" -ForegroundColor Gray
    $allSongs = @()
    $processedCount = 0
    
    foreach ($imageFile in $imageFiles) {
        $processedCount++
        Write-Host "  [$processedCount/$($imageFiles.Count)] $($imageFile.Name)" -ForegroundColor Gray
        
        $text = Extract-TextFromImage-GoogleVision -ImagePath $imageFile.FullName -AccessToken $accessToken
        
        if ($text) {
            $songs = Parse-Setlist -Text $text -FileName $imageFile.Name
            $allSongs += $songs
        }
        
        # APIレート制限対策（念のため）
        Start-Sleep -Milliseconds 200
    }
    
    Write-Host "✓ OCR処理完了: 全 $($allSongs.Count) 曲検出" -ForegroundColor Green
    
    if ($allSongs.Count -eq 0) {
        Write-Host "`n⚠ 曲が検出されませんでした" -ForegroundColor Yellow
        Write-Host "OCR出力を確認してください: $OcrOutputDir" -ForegroundColor Yellow
        return
    }
    
    # 5. 集計
    Write-Host "`n[5] 曲名の集計..." -ForegroundColor Yellow
    $songCounts = Get-SongStatistics -AllSongs $allSongs
    
    # 6. 結果出力
    Write-Host "`n[6] 結果出力..." -ForegroundColor Yellow
    Export-Results -SongCounts $songCounts -TotalImages $imageFiles.Count
    
    Write-Host "`n" + ("=" * 80) -ForegroundColor Green
    Write-Host "✓ 処理完了！" -ForegroundColor Green
    Write-Host ("=" * 80) -ForegroundColor Green
    Write-Log "処理完了"
}

# ====================
# スクリプト実行
# ====================

try {
    Main
}
catch {
    Write-Log "予期しないエラー: $_" -Level "ERROR"
    Write-Host "`nエラーが発生しました。ログファイルを確認してください。" -ForegroundColor Red
    Write-Host "エラー詳細: $_" -ForegroundColor Red
}