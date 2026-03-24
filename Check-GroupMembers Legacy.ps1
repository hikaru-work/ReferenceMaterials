<#
.SYNOPSIS
    指定ローカルグループの全メンバーを列挙・CSV 出力するスクリプト

.DESCRIPTION
    $TargetGroups に定義した各ローカルグループに所属する全アカウント
    （ローカル / Active Directory / Microsoft アカウント / Azure AD）を
    コンソール表示および CSV ファイルに出力

    SID 未解決アカウント（退職者等の残骸）を検出した場合は
    Status 列に "Unresolved" を記録し、戻り値 3 で運用者に通知する

.OUTPUTS
    ローカル（一次出力）:
        <スクリプトフォルダ>\Logs\
            <HostName>_Operations_<yyyyMMdd_HHmmss>.log  … セッション単位ログ
            <HostName>_GroupMembers_<yyyyMMdd_HHmmss>.csv … 全メンバー一覧

    共有フォルダ（二次コピー）:
        <$OutputFolder>\<yyyyMMdd>\
            上記と同名のファイルを .tmp で書き込み → 完了後にリネーム（不完全ファイルの残留を防止）

    戻り値:
        0 … 検出対象アカウントなし・問題なし
        1 … 検出対象アカウントあり（$DetectAccountTypes で定義した種別）
        2 … 実行前エラー（環境不備）
        3 … SID 未解決アカウントあり（残骸の可能性・要確認）
        4 … 共有フォルダへのコピー失敗（ローカルにデータは残存）

.NOTES
    実行要件 : 管理者権限の PowerShell（FullLanguage モード）
    対象 OS  : Windows 7 SP1 以降（PowerShell 2.0 対応レガシー版）
    備考     : PS 3.0 以上の環境では Check-GroupMembers.ps1（モダン版）を使用してください
#>

# 注意: #Requires -RunAsAdministrator は PS 4.0 以降の構文のため使用しない
# 管理者権限のチェックは Run_Check-GroupMembers.bat で行う

# ============================================================
# ■ パラメータ定義（探索列挙する対象ローカルグループ）
# ============================================================
$TargetGroups = @(
    "Administrators"
    # "Remote Desktop Users"
    # "Power Users"
    # "Backup Operators"
    # "S-1-5-32-580"        # Remote Management Users（SID 指定例）
)

# ============================================================
# ■ パラメータ定義（戻り値 1 で検出通知する対象アカウント種別）
# ============================================================
# 以下に定義した種別のアカウントが1件でも存在すれば戻り値 1 を返す
# 有効な値: ActiveDirectory, MicrosoftAccount, AzureAD, Local, Unknown
$DetectAccountTypes = @(
    "ActiveDirectory"
    # "MicrosoftAccount"
    # "AzureAD"
    # "Local"
    # "Unknown"
)

# ============================================================
# ■ パラメータ定義（検出カウントから除外するアカウント）
# ============================================================
# グループごとに、検出カウント・通知（戻り値 1）から除外するアカウント名を指定
# CSV・ログ出力には影響しない（全メンバーが出力される）
# アカウント名は完全一致（大文字小文字は区別しない）
$ExcludeAccounts = @{
    # グループ名 = @(除外アカウント名)
     "Administrators" = @(
         "PARK24\Domain Admins"
    #     "PC-001\Administrator"
     )
    # "Remote Desktop Users" = @(
    #     "TEST\helpdesk01"
    # )
}

# ============================================================
# ■ パラメータ定義（出力先 共有フォルダ接続）
# ============================================================
# net use で認証する共有パス（\\サーバ名\共有名）
$NetworkShare  = "\\fileserver\share"
# ログ・CSV のコピー先ベースフォルダ（配下に日付フォルダが自動作成される）
$OutputFolder  = "\\fileserver\share\Logs"
# 接続用アカウント（DOMAIN\Username 形式）
$ShareUser     = "DOMAIN\svc_logwriter"
# 接続用パスワード（※ スクリプトに平文で保持されます）
$SharePassword = "P@ssw0rd"

# ============================================================
# ■ PS 2.0 互換ヘルパー
# ============================================================
# [string]::IsNullOrWhiteSpace は .NET 4.0 以降のため、互換関数を定義
function Test-StringNullOrWhiteSpace {
    param([string]$Value)
    if ($null -eq $Value) { return $true }
    if ($Value.Trim() -eq "") { return $true }
    return $false
}

# ============================================================
# ■ 事前チェック
# ============================================================

# 言語モードチェック
$langMode = $ExecutionContext.SessionState.LanguageMode
if ($langMode -ne "FullLanguage") {
    Write-Host "[ERROR] 言語モードが '$langMode' のため実行できません。FullLanguage モードで実行してください。" -ForegroundColor Red
    exit 2
}

# スクリプトパスの取得（PS 2.0 互換）
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if (Test-StringNullOrWhiteSpace $ScriptRoot) {
    Write-Host "[ERROR] スクリプトパスが取得できません。.ps1 ファイルとして保存してから実行してください。" -ForegroundColor Red
    exit 2
}

# ============================================================
# ■ 出力先フォルダ・ファイル定義（ローカル一次出力）
# ============================================================
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$DateStamp = $Timestamp.Substring(0, 8)
$script:HostName = $env:COMPUTERNAME
$LogDir    = Join-Path $ScriptRoot "Logs"

try {
    if (-not (Test-Path $LogDir)) {
        New-Item -Path $LogDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Host "[ERROR] ローカルログフォルダの作成に失敗しました: $_" -ForegroundColor Red
    exit 2
}

# ログファイル（セッション単位・ホスト名プレフィックス）
$LogFile = Join-Path $LogDir ("{0}_Operations_{1}.log" -f $script:HostName, $Timestamp)

# CSV ファイル（ホスト名プレフィックス・実行ごとにタイムスタンプ付き）
$CsvFile = Join-Path $LogDir ("{0}_GroupMembers_{1}.csv" -f $script:HostName, $Timestamp)

# ログファイルへの書き込みテスト
try {
    "" | Out-File -FilePath $LogFile -Encoding UTF8 -ErrorAction Stop
}
catch {
    Write-Host "[ERROR] ローカルログファイルに書き込みできません: $LogFile" -ForegroundColor Red
    exit 2
}

# ============================================================
# ■ ログ出力ヘルパー
# ============================================================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")]
        [string]$Level = "INFO"
    )
    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message

    $color = switch ($Level) {
        "ERROR"   { "Red"    }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green"  }
        default   { "White"  }
    }
    Write-Host $entry -ForegroundColor $color

    try {
        $entry | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host "[ERROR] ログ書き込み失敗: $_" -ForegroundColor Red
    }
}

# ============================================================
# ■ 実行ユーザー情報取得ヘルパー
# ============================================================
function Get-ExecutionUserInfo {
    $info = @{
        Domain   = ""
        UserName = ""
        Display  = ""
    }

    # 第1候補: 環境変数
    $domain   = $env:USERDOMAIN
    $username = $env:USERNAME

    if (-not (Test-StringNullOrWhiteSpace $domain) -and -not (Test-StringNullOrWhiteSpace $username)) {
        $info.Domain   = $domain
        $info.UserName = $username
        $info.Display  = "{0}\{1}" -f $domain, $username
        return $info
    }

    # 第2候補: whoami コマンド
    try {
        $whoami = whoami 2>$null
        if (-not (Test-StringNullOrWhiteSpace $whoami)) {
            $parts = $whoami -split "\\"
            if ($parts.Count -ge 2) {
                $info.Domain   = $parts[0]
                $info.UserName = $parts[1]
            }
            else {
                $info.UserName = $whoami
            }
            $info.Display = $whoami
            return $info
        }
    }
    catch {
        # whoami 失敗時は次へ
    }

    $info.Display = "(取得不可)"
    return $info
}

# ============================================================
# ■ SID 未解決判定ヘルパー
# ============================================================
function Test-UnresolvedSID {
    param([string]$Name)

    # 空文字・null は「名前取得不可」であり SID 未解決とは断定できない
    if (Test-StringNullOrWhiteSpace $Name) {
        return $false
    }

    $parts    = $Name -split "\\"
    $namePart = $parts[-1]

    if ($namePart -match '^S-1-\d+-') {
        return $true
    }

    return $false
}

# ============================================================
# ■ グループメンバー取得ヘルパー
# ============================================================

$script:KnownGroupSids = @{
    "Administrators"                  = "S-1-5-32-544"
    "Remote Desktop Users"            = "S-1-5-32-555"
    "Power Users"                     = "S-1-5-32-547"
    "Backup Operators"                = "S-1-5-32-551"
    "Users"                           = "S-1-5-32-545"
    "Guests"                          = "S-1-5-32-546"
    "Network Configuration Operators" = "S-1-5-32-556"
    "Event Log Readers"               = "S-1-5-32-573"
    "Remote Management Users"         = "S-1-5-32-580"
}

function Check-GroupMembersByCommandlet {
    param([string]$GroupIdentifier)

    if ($GroupIdentifier -match '^S-1-\d+-') {
        return @(Get-LocalGroupMember -SID $GroupIdentifier -ErrorAction Stop)
    }

    try {
        return @(Get-LocalGroupMember -Group $GroupIdentifier -ErrorAction Stop)
    }
    catch {
        if ($script:KnownGroupSids.ContainsKey($GroupIdentifier)) {
            $sid = $script:KnownGroupSids[$GroupIdentifier]
            Write-Log ("  グループ名 '{0}' で取得失敗 → SID {1} で再試行" -f $GroupIdentifier, $sid) "WARN"
            return @(Get-LocalGroupMember -SID $sid -ErrorAction Stop)
        }
        throw
    }
}

function Check-GroupMembersByADSI {
    param([string]$GroupIdentifier)

    $groupName = $GroupIdentifier

    if ($GroupIdentifier -match '^S-1-\d+-') {
        try {
            $sidObj  = New-Object System.Security.Principal.SecurityIdentifier($GroupIdentifier)
            $account = $sidObj.Translate([System.Security.Principal.NTAccount])
            $groupName = ($account.Value -split "\\")[-1]
        }
        catch {
            throw "SID '$GroupIdentifier' からグループ名を解決できません: $_"
        }
    }

    $adsiGroup = $null
    try {
        $adsiGroup = [ADSI]"WinNT://$($script:HostName)/$groupName,group"
    }
    catch {
        throw "ADSI でグループ '$groupName' を取得できません: $_"
    }

    try {
        $memberPaths = @($adsiGroup.Invoke("Members"))
    }
    catch {
        throw "ADSI でメンバー一覧を取得できません: $_"
    }

    $results = @()

    foreach ($memberObj in $memberPaths) {
        try {
            $adsPath   = $memberObj.GetType().InvokeMember("ADsPath", "GetProperty", $null, $memberObj, $null)
            $name      = $memberObj.GetType().InvokeMember("Name", "GetProperty", $null, $memberObj, $null)
            $className = $memberObj.GetType().InvokeMember("Class", "GetProperty", $null, $memberObj, $null)

            $sid      = ""
            $fullName = ""

            $pathParts = $adsPath -replace "^WinNT://", "" -split "/"
            if ($pathParts.Count -ge 2) {
                # 2パート: WinNT://DOMAIN/user       → DOMAIN\user
                # 3パート: WinNT://DOMAIN/PC/user    → PC\user（ローカルアカウント）
                $fullName = "{0}\{1}" -f $pathParts[-2], $pathParts[-1]
            }
            else {
                $fullName = $name
            }

            try {
                $sidBytes = $memberObj.GetType().InvokeMember("ObjectSid", "GetProperty", $null, $memberObj, $null)
                if ($null -ne $sidBytes) {
                    $sidObj = New-Object System.Security.Principal.SecurityIdentifier($sidBytes, 0)
                    $sid = $sidObj.Value
                }
            }
            catch {
                $sid = "(取得不可)"
            }

            $isUnresolved = Test-UnresolvedSID -Name $fullName

            $obj = New-Object PSObject
            $obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $fullName
            $obj | Add-Member -MemberType NoteProperty -Name "ObjectClass" -Value $className
            $obj | Add-Member -MemberType NoteProperty -Name "SID" -Value $sid
            $obj | Add-Member -MemberType NoteProperty -Name "Source" -Value "ADSI"
            $obj | Add-Member -MemberType NoteProperty -Name "IsUnresolved" -Value $isUnresolved
            $results += $obj
        }
        catch {
            Write-Log ("    ADSI メンバー情報の取得に失敗（スキップ）: {0}" -f $_.Exception.Message) "WARN"
        }
    }

    return $results
}

function Check-GroupMembersSafe {
    param([string]$GroupIdentifier)

    try {
        $members = Check-GroupMembersByCommandlet -GroupIdentifier $GroupIdentifier
        if ($null -ne $members) {
            Write-Log "  取得方法: Get-LocalGroupMember"
            return @{
                Members = $members
                Method  = "Cmdlet"
            }
        }
    }
    catch {
        Write-Log ("  Get-LocalGroupMember 失敗: {0}" -f $_.Exception.Message) "WARN"
        Write-Log "  ADSI フォールバックを試行します..." "WARN"
    }

    try {
        $members = Check-GroupMembersByADSI -GroupIdentifier $GroupIdentifier
        if ($null -ne $members) {
            Write-Log "  取得方法: ADSI (WinNT プロバイダ)" "WARN"
            Write-Log "  ※ Get-LocalGroupMember が失敗したため ADSI で取得しました。SID 未解決アカウント（残骸）が含まれている可能性があります。" "WARN"
            return @{
                Members = $members
                Method  = "ADSI"
            }
        }
    }
    catch {
        Write-Log ("  ADSI フォールバックも失敗: {0}" -f $_.Exception.Message) "ERROR"
    }

    return $null
}

# ============================================================
# ■ アカウント種別判定ヘルパー
# ============================================================
function Get-AccountType {
    param(
        $Member,
        [string]$Method
    )

    if ($Method -eq "Cmdlet") {
        if ($null -ne $Member.PSObject -and $Member.PSObject.Properties.Name -contains "PrincipalSource") {
            $source = $Member.PrincipalSource
            if ($null -ne $source) {
                switch ($source.ToString()) {
                    "Local"            { return "Local" }
                    "ActiveDirectory"  { return "ActiveDirectory" }
                    "MicrosoftAccount" { return "MicrosoftAccount" }
                    default            { return $source.ToString() }
                }
            }
        }
    }

    $memberName = $Member.Name

    if (Test-StringNullOrWhiteSpace $memberName) {
        return "Unknown"
    }

    $parts = $memberName -split "\\"
    if ($parts.Count -lt 2) {
        return "Unknown"
    }

    $prefix = $parts[0].ToUpper()

    if ($prefix -eq $script:HostName.ToUpper()) {
        return "Local"
    }
    elseif ($prefix -eq "MICROSOFTACCOUNT") {
        return "MicrosoftAccount"
    }
    elseif ($prefix -eq "AZUREAD") {
        return "AzureAD"
    }
    else {
        return "ActiveDirectory"
    }
}

function Get-MemberName {
    param($Member)

    $name = $Member.Name
    if (Test-StringNullOrWhiteSpace $name) {
        return "(名前取得不可)"
    }
    return $name
}

function Get-MemberObjectClass {
    param($Member)

    $class = $Member.ObjectClass
    if (Test-StringNullOrWhiteSpace $class) {
        return "(不明)"
    }
    return $class
}

function Get-MemberSID {
    param($Member, [string]$Method)

    if ($Method -eq "Cmdlet") {
        if ($null -ne $Member.SID) {
            return $Member.SID.Value
        }
    }
    else {
        if (-not (Test-StringNullOrWhiteSpace $Member.SID)) {
            return $Member.SID
        }
    }
    return "(取得不可)"
}

function Get-MemberStatus {
    param(
        $Member,
        [string]$Method
    )

    if ($Method -eq "Cmdlet") {
        $rawName = $Member.Name
        if (Test-UnresolvedSID -Name $rawName) {
            return "Unresolved"
        }
        # Name が空でも SID 文字列から未解決を判定（フォールバック）
        if ((Test-StringNullOrWhiteSpace $rawName) -and $null -ne $Member.SID) {
            try {
                $Member.SID.Translate([System.Security.Principal.NTAccount]) | Out-Null
            }
            catch {
                return "Unresolved"
            }
        }
        return "Normal"
    }

    if ($Method -eq "ADSI") {
        if ($null -ne $Member.PSObject -and $Member.PSObject.Properties.Name -contains "IsUnresolved") {
            if ($Member.IsUnresolved -eq $true) {
                return "Unresolved"
            }
        }
        $rawName = $Member.Name
        if (Test-UnresolvedSID -Name $rawName) {
            return "Unresolved"
        }
        return "Normal"
    }

    return "Unknown"
}

# ============================================================
# ■ 除外判定ヘルパー
# ============================================================
function Test-ExcludedAccount {
    param(
        [string]$GroupName,
        [string]$AccountName,
        [hashtable]$ExcludeList
    )

    if ($null -eq $ExcludeList -or $ExcludeList.Count -eq 0) {
        return $false
    }

    if (-not $ExcludeList.ContainsKey($GroupName)) {
        return $false
    }

    $excludeNames = $ExcludeList[$GroupName]
    if ($null -eq $excludeNames -or @($excludeNames).Count -eq 0) {
        return $false
    }

    # 完全一致（大文字小文字は区別しない — PowerShell の -contains は既定で case-insensitive）
    # ".\" をコンピュータ名に展開して比較
    $normalizedExclude = @($excludeNames) | ForEach-Object {
        if ($_ -like '.\*') {
            "{0}\{1}" -f $script:HostName, $_.Substring(2)
        }
        else {
            $_
        }
    }
    $normalizedAccount = if ($AccountName -like '.\*') {
        "{0}\{1}" -f $script:HostName, $AccountName.Substring(2)
    }
    else {
        $AccountName
    }

    return $normalizedExclude -contains $normalizedAccount
}

# ============================================================
# 処理開始
# ============================================================
$userInfo = Get-ExecutionUserInfo

Write-Log ""
Write-Log ("=" * 60)
Write-Log ("[SESSION START] Check-GroupMembers {0}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
Write-Log ("=" * 60)
Write-Log "スクリプト開始  コンピュータ名: $($script:HostName)"
Write-Log ("実行ユーザー  : {0}" -f $userInfo.Display)
Write-Log ("言語モード    : {0}" -f $langMode)
Write-Log ("ローカル出力  : {0}" -f $LogDir)
Write-Log ("共有コピー先  : {0}\{1}\" -f $OutputFolder, $DateStamp)

# ============================================================
# パラメータ検証
# ============================================================
if ($null -eq $TargetGroups -or $TargetGroups.Count -eq 0) {
    Write-Log "対象グループが未定義です。スクリプトを終了します。" "ERROR"
    Write-Log ("[SESSION END] 環境不備 (exit 2)")
    Write-Log ("=" * 60)
    exit 2
}

$validGroups = @()
foreach ($g in $TargetGroups) {
    if (Test-StringNullOrWhiteSpace $g) {
        Write-Log "対象グループに空文字が含まれています。スキップします。" "WARN"
    }
    else {
        $validGroups += $g
    }
}

if ($validGroups.Count -eq 0) {
    Write-Log "有効な対象グループがありません。スクリプトを終了します。" "ERROR"
    Write-Log ("[SESSION END] 環境不備 (exit 2)")
    Write-Log ("=" * 60)
    exit 2
}

Write-Log ("対象グループ: {0}" -f ($validGroups -join ", "))

# 検出対象種別の検証
$validAccountTypes = @("ActiveDirectory", "MicrosoftAccount", "AzureAD", "Local", "Unknown")

if ($null -eq $DetectAccountTypes -or $DetectAccountTypes.Count -eq 0) {
    Write-Log "検出対象種別が未定義です。列挙のみ実施し、戻り値 1 は返しません。" "WARN"
    $DetectAccountTypes = @()
}
else {
    $invalidTypes = @($DetectAccountTypes | Where-Object { $validAccountTypes -notcontains $_ })
    if ($invalidTypes.Count -gt 0) {
        Write-Log ("検出対象種別に不正な値があります: {0}" -f ($invalidTypes -join ", ")) "WARN"
        Write-Log ("有効な値: {0}" -f ($validAccountTypes -join ", ")) "WARN"
        $DetectAccountTypes = @($DetectAccountTypes | Where-Object { $validAccountTypes -contains $_ })
        if ($DetectAccountTypes.Count -eq 0) {
            Write-Log "有効な検出対象種別が残りません。列挙のみ実施します。" "WARN"
        }
    }
}

Write-Log ("検出対象種別: {0}" -f $(if ($DetectAccountTypes.Count -gt 0) { $DetectAccountTypes -join ", " } else { "(なし — 列挙のみ)" }))

# 除外アカウントの検証
if ($null -eq $ExcludeAccounts) {
    $ExcludeAccounts = @{}
}

if ($ExcludeAccounts.Count -gt 0) {
    $excludeTotal = 0
    foreach ($exGroup in @($ExcludeAccounts.Keys)) {
        # 対象グループに含まれないキーを警告
        if ($validGroups -notcontains $exGroup) {
            Write-Log ("除外リストのグループ '{0}' は対象グループに含まれていません。無視されます。" -f $exGroup) "WARN"
            continue
        }

        # 空文字エントリを除去
        $entries = @($ExcludeAccounts[$exGroup])
        $validEntries = @($entries | Where-Object { -not (Test-StringNullOrWhiteSpace $_) })
        $emptyCount = $entries.Count - $validEntries.Count
        if ($emptyCount -gt 0) {
            Write-Log ("除外リスト '{0}' に空文字が {1} 件含まれています。スキップします。" -f $exGroup, $emptyCount) "WARN"
            $ExcludeAccounts[$exGroup] = $validEntries
        }

        $excludeTotal += $validEntries.Count
    }

    Write-Log ("除外アカウント: {0} 件（{1} グループ）" -f $excludeTotal, @($ExcludeAccounts.Keys | Where-Object { $validGroups -contains $_ }).Count)
    foreach ($exGroup in @($ExcludeAccounts.Keys) | Where-Object { $validGroups -contains $_ }) {
        foreach ($exName in @($ExcludeAccounts[$exGroup])) {
            Write-Log ("  除外: [{0}] {1}" -f $exGroup, $exName)
        }
    }
}
else {
    Write-Log "除外アカウント: (なし)"
}

# ============================================================
# グループメンバー列挙
# ============================================================
$allCsvRecords    = @()
$detectedCount    = 0
$excludedCount    = 0
$unresolvedCount  = 0

foreach ($group in $validGroups) {
    Write-Log ""
    Write-Log (">>> グループ: {0}" -f $group)

    $result = Check-GroupMembersSafe -GroupIdentifier $group
    if ($null -eq $result) {
        Write-Log ("  グループ '{0}' のメンバーを取得できませんでした。" -f $group) "ERROR"
        continue
    }

    $members = $result.Members
    $method  = $result.Method

    if ($null -eq $members -or @($members).Count -eq 0) {
        Write-Log "  メンバーなし"
        continue
    }

    Write-Log ("  メンバー取得成功（{0} 件）" -f @($members).Count)

    foreach ($member in $members) {
        if ($null -eq $member) {
            Write-Log "    null メンバーをスキップ" "WARN"
            continue
        }

        $accountType = Get-AccountType -Member $member -Method $method
        $memberName  = Get-MemberName -Member $member
        $objectClass = Get-MemberObjectClass -Member $member
        $memberSID   = Get-MemberSID -Member $member -Method $method
        $status      = Get-MemberStatus -Member $member -Method $method

        # SID 未解決の場合、種別判定は信頼できないため Unknown に強制
        if ($status -eq "Unresolved") {
            $accountType = "Unknown"
        }

        # 除外判定
        $isExcluded = Test-ExcludedAccount -GroupName $group -AccountName $memberName -ExcludeList $ExcludeAccounts

        # 検出対象アカウント種別カウント（SID 未解決は除外）
        if ($status -ne "Unresolved" -and $DetectAccountTypes -contains $accountType) {
            if ($isExcluded) {
                $excludedCount++
            }
            else {
                $detectedCount++
            }
        }
        if ($status -eq "Unresolved") {
            $unresolvedCount++
        }

        # ログレベル判定
        $logLevel = if ($status -eq "Unresolved") {
            "ERROR"
        }
        elseif ($isExcluded) {
            "INFO"
        }
        elseif ($DetectAccountTypes -contains $accountType) {
            "SUCCESS"
        }
        else {
            switch ($accountType) {
                "MicrosoftAccount" { "WARN" }
                "AzureAD"          { "WARN" }
                "Unknown"          { "WARN" }
                default            { "INFO" }
            }
        }

        $statusTag  = if ($status -eq "Unresolved") { " ★残骸の可能性" } else { "" }
        $excludeTag = if ($isExcluded) { " [除外]" } else { "" }
        Write-Log ("    [{0}] {1}  (種類: {2} / SID: {3} / 状態: {4}{5}{6})" -f $accountType, $memberName, $objectClass, $memberSID, $status, $statusTag, $excludeTag) $logLevel

        # 検出フラグ（検出対象種別に該当 かつ 除外されていない かつ SID 未解決でない）
        $detected = ($status -ne "Unresolved" -and -not $isExcluded -and $DetectAccountTypes -contains $accountType)

        $rec = New-Object PSObject
        $rec | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $script:HostName
        $rec | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $group
        $rec | Add-Member -MemberType NoteProperty -Name "AccountName" -Value $memberName
        $rec | Add-Member -MemberType NoteProperty -Name "AccountType" -Value $accountType
        $rec | Add-Member -MemberType NoteProperty -Name "ObjectClass" -Value $objectClass
        $rec | Add-Member -MemberType NoteProperty -Name "SID" -Value $memberSID
        $rec | Add-Member -MemberType NoteProperty -Name "Status" -Value $status
        $rec | Add-Member -MemberType NoteProperty -Name "Detected" -Value $detected
        $rec | Add-Member -MemberType NoteProperty -Name "RetrievalMethod" -Value $method
        $rec | Add-Member -MemberType NoteProperty -Name "CollectedAt" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        $allCsvRecords += $rec
    }

    $groupRecords = @($allCsvRecords | Where-Object { $_.GroupName -eq $group })
    if ($groupRecords.Count -gt 0) {
        $typeSummary = $groupRecords | Group-Object AccountType | ForEach-Object {
            "{0}={1}" -f $_.Name, $_.Count
        }
        Write-Log ("  種別サマリ  : {0}" -f ($typeSummary -join "  "))

        $unresolvedInGroup = @($groupRecords | Where-Object { $_.Status -eq "Unresolved" }).Count
        if ($unresolvedInGroup -gt 0) {
            Write-Log ("  ★ SID 未解決 : {0} 件（残骸の可能性があります。確認してください）" -f $unresolvedInGroup) "ERROR"
        }
    }
}

# ============================================================
# CSV 出力
# ============================================================
$csvColumns = @("ComputerName","GroupName","AccountName","AccountType","ObjectClass","SID","Status","Detected","RetrievalMethod","CollectedAt")

try {
    if ($allCsvRecords.Count -gt 0) {
        $allCsvRecords | Select-Object $csvColumns | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log ""
        Write-Log ("CSV 出力完了 ({0} 件): {1}" -f $allCsvRecords.Count, $CsvFile) "SUCCESS"
    }
    else {
        # ヘッダのみの空 CSV を出力（ダミー行なし）
        $header = '"ComputerName","GroupName","AccountName","AccountType","ObjectClass","SID","Status","Detected","RetrievalMethod","CollectedAt"'
        $header | Out-File -FilePath $CsvFile -Encoding UTF8 -ErrorAction Stop
        Write-Log ""
        Write-Log ("CSV 出力完了（全グループメンバーなし）: {0}" -f $CsvFile) "WARN"
    }
}
catch {
    Write-Log ("CSV 出力に失敗しました: {0}" -f $_.Exception.Message) "ERROR"
    Write-Log ("[SESSION END] CSV 出力失敗 (exit 2)") "ERROR"
    Write-Log ("=" * 60)
    exit 2
}

# ============================================================
# コンソール テーブル表示
# ============================================================
Write-Host ""
Write-Host ("─" * 75) -ForegroundColor Cyan
Write-Host " 指定グループ メンバー一覧" -ForegroundColor Cyan
Write-Host ("─" * 75) -ForegroundColor Cyan

if ($allCsvRecords.Count -gt 0) {
    $allCsvRecords | Format-Table GroupName, AccountName, AccountType, Status, Detected, RetrievalMethod, SID -AutoSize
}
else {
    Write-Host "  (該当なし)" -ForegroundColor DarkGray
}

# ============================================================
# 完了・戻り値判定
# ============================================================
Write-Log ""
Write-Log "検出サマリ:"
Write-Log ("  検出対象種別     : {0}" -f $(if ($DetectAccountTypes.Count -gt 0) { $DetectAccountTypes -join ", " } else { "(なし — 列挙のみ)" }))
Write-Log ("  検出対象アカウント: {0} 件" -f $detectedCount)
Write-Log ("  除外アカウント   : {0} 件" -f $excludedCount)
Write-Log ("  SID 未解決       : {0} 件" -f $unresolvedCount)
Write-Log ("  CSV ファイル     : {0}" -f $CsvFile)
Write-Log ""

$exitCode = 0

if ($unresolvedCount -gt 0) {
    $exitCode = 3
    Write-Log "★★ SID 未解決アカウントが検出されました。" "ERROR"
    Write-Log "   退職者等の AD アカウント残骸の可能性があります。" "ERROR"
    Write-Log "   CSV の Status 列が 'Unresolved' のレコードを確認し、" "ERROR"
    Write-Log "   不要であればローカルグループから削除してください。" "ERROR"
}
elseif ($detectedCount -gt 0) {
    $exitCode = 1
    Write-Log ("★ 検出対象アカウントが {0} 件見つかりました。（除外: {1} 件）" -f $detectedCount, $excludedCount) "WARN"
}
else {
    if ($excludedCount -gt 0) {
        Write-Log ("検出対象アカウントは除外分のみでした。（除外: {0} 件）" -f $excludedCount)
    }
}

# ============================================================
# 共有フォルダへの安全コピー（.tmp 経由）
# ============================================================
Write-Log ""
Write-Log "共有フォルダへのコピーを開始します..."

$copySuccess = $true

# net use 接続
net use $NetworkShare /delete /y 2>$null | Out-Null
$null = net use $NetworkShare /user:$ShareUser $SharePassword /persistent:no 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Log ("共有フォルダへの接続に失敗しました: {0}" -f $NetworkShare) "ERROR"
    Write-Log ("  net use 終了コード: {0}" -f $LASTEXITCODE) "ERROR"
    $copySuccess = $false
}

if ($copySuccess) {
    # 日付フォルダ作成
    $remoteDateDir = Join-Path $OutputFolder $DateStamp
    try {
        if (-not (Test-Path $remoteDateDir)) {
            New-Item -Path $remoteDateDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }
    catch {
        Write-Log ("日付フォルダの作成に失敗しました: {0}" -f $_.Exception.Message) "ERROR"
        $copySuccess = $false
    }
}

if ($copySuccess) {
    # ファイルコピー（.tmp で書き込み → 完了後にリネーム）
    $filesToCopy = @($LogFile, $CsvFile)

    foreach ($localFile in $filesToCopy) {
        if (-not (Test-Path $localFile)) {
            Write-Log ("ローカルファイルが存在しません（スキップ）: {0}" -f $localFile) "WARN"
            continue
        }

        $fileName  = [System.IO.Path]::GetFileName($localFile)
        $remoteTmp  = Join-Path $remoteDateDir ("$fileName.tmp")
        $remoteFinal = Join-Path $remoteDateDir $fileName

        try {
            Copy-Item -Path $localFile -Destination $remoteTmp -Force -ErrorAction Stop
            # 同名ファイルが既にある場合は上書き（リトライ等）
            if (Test-Path $remoteFinal) {
                Remove-Item -Path $remoteFinal -Force -ErrorAction Stop
            }
            Rename-Item -Path $remoteTmp -NewName $fileName -ErrorAction Stop
            Write-Log ("  コピー完了: {0}" -f $remoteFinal) "SUCCESS"
        }
        catch {
            Write-Log ("  コピー失敗: {0} — {1}" -f $fileName, $_.Exception.Message) "ERROR"
            # .tmp が残っていれば削除を試みる
            if (Test-Path $remoteTmp) {
                Remove-Item -Path $remoteTmp -Force -ErrorAction SilentlyContinue
            }
            $copySuccess = $false
        }
    }
}

# net use 切断
net use $NetworkShare /delete /y 2>$null | Out-Null

if ($copySuccess) {
    Write-Log "共有フォルダへのコピーが完了しました。" "SUCCESS"
}
else {
    Write-Log "★ 共有フォルダへのコピーに失敗しました。ローカルにデータは残存しています。" "ERROR"
    Write-Log ("  ローカルログ : {0}" -f $LogFile) "ERROR"
    Write-Log ("  ローカル CSV : {0}" -f $CsvFile) "ERROR"
    $exitCode = 4
}

# ============================================================
# セッション終了
# ============================================================
Write-Log ""
switch ($exitCode) {
    0 { Write-Log ("[SESSION END] 正常終了 (exit {0})" -f $exitCode) }
    1 { Write-Log ("[SESSION END] 検出対象アカウントあり (exit {0})" -f $exitCode) "WARN" }
    3 { Write-Log ("[SESSION END] SID 未解決検出 (exit {0})" -f $exitCode) "ERROR" }
    4 { Write-Log ("[SESSION END] 共有フォルダコピー失敗 (exit {0})" -f $exitCode) "ERROR" }
}
Write-Log ("=" * 60)

exit $exitCode