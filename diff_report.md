# 差分レポート

Check-GroupMembers.ps1（モダン版）→ Check-GroupMembers_Legacy.ps1（レガシー版）

-----

## 差分①：ヘッダ .NOTES

```diff
 .NOTES
     実行要件 : 管理者権限の PowerShell（FullLanguage モード）
-    対象 OS  : Windows 10 / 11, Windows Server 2016 以降
+    対象 OS  : Windows 7 SP1 以降（PowerShell 2.0 対応レガシー版）
+    備考     : PS 3.0 以上の環境では Check-GroupMembers.ps1（モダン版）を使用してください
 #>
```

-----

## 差分②：#Requires 削除

```diff
-#Requires -RunAsAdministrator
+# 注意: #Requires -RunAsAdministrator は PS 4.0 以降の構文のため使用しない
+# 管理者権限のチェックは Run_Check-GroupMembers.bat で行う

 # ============================================================
 # ■ パラメータ定義（探索列挙する対象ローカルグループ）
```

-----

## 差分③：PS 2.0 互換ヘルパー関数追加（事前チェックの直前に挿入）

```diff
 $SharePassword = "P@ssw0rd"

+# ============================================================
+# ■ PS 2.0 互換ヘルパー
+# ============================================================
+# [string]::IsNullOrWhiteSpace は .NET 4.0 以降のため、互換関数を定義
+function Test-StringNullOrWhiteSpace {
+    param([string]$Value)
+    if ($null -eq $Value) { return $true }
+    if ($Value.Trim() -eq "") { return $true }
+    return $false
+}
+
 # ============================================================
 # ■ 事前チェック
 # ============================================================
```

-----

## 差分④：言語モードチェック

```diff
 # 言語モードチェック
 $langMode = $ExecutionContext.SessionState.LanguageMode
 if ($langMode -ne "FullLanguage") {
     Write-Host "[ERROR] 言語モードが '$langMode' のため実行できません。FullLanguage モードで実行してください。" -ForegroundColor Red
     exit 2
 }

-# $PSScriptRoot の有効性チェック
-if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) {
-    Write-Host "[ERROR] `$PSScriptRoot が取得できません。.ps1 ファイルとして保存してから実行してください。" -ForegroundColor Red
+# スクリプトパスの取得（PS 2.0 互換）
+$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
+if ([string]::IsNullOrEmpty($ScriptRoot)) {
+    Write-Host "[ERROR] スクリプトパスが取得できません。.ps1 ファイルとして保存してから実行してください。" -ForegroundColor Red
     exit 2
 }
```

-----

## 差分⑤：出力先フォルダ定義

```diff
 $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
 $DateStamp = $Timestamp.Substring(0, 8)
 $script:HostName = $env:COMPUTERNAME
-$LogDir    = Join-Path $PSScriptRoot "Logs"
+$LogDir    = Join-Path $ScriptRoot "Logs"

 try {
     if (-not (Test-Path $LogDir)) {
```

-----

## 差分⑥：実行ユーザー情報取得ヘルパー

```diff
     $domain   = $env:USERDOMAIN
     $username = $env:USERNAME

-    if (-not [string]::IsNullOrWhiteSpace($domain) -and -not [string]::IsNullOrWhiteSpace($username)) {
+    if (-not (Test-StringNullOrWhiteSpace $domain) -and -not (Test-StringNullOrWhiteSpace $username)) {
         $info.Domain   = $domain
         $info.UserName = $username
         $info.Display  = "{0}\{1}" -f $domain, $username
```

```diff
     try {
         $whoami = whoami 2>$null
-        if (-not [string]::IsNullOrWhiteSpace($whoami)) {
+        if (-not (Test-StringNullOrWhiteSpace $whoami)) {
             $parts = $whoami -split "\\"
```

-----

## 差分⑦：SID 未解決判定ヘルパー

```diff
 function Test-UnresolvedSID {
     param([string]$Name)

     # 空文字・null は「名前取得不可」であり SID 未解決とは断定できない
-    if ([string]::IsNullOrWhiteSpace($Name)) {
+    if (Test-StringNullOrWhiteSpace $Name) {
         return $false
     }
```

-----

## 差分⑧：ADSI フォールバック — PSCustomObject → Add-Member

```diff
             $isUnresolved = Test-UnresolvedSID -Name $fullName

-            $results += [PSCustomObject]@{
-                Name         = $fullName
-                ObjectClass  = $className
-                SID          = $sid
-                Source       = "ADSI"
-                IsUnresolved = $isUnresolved
-            }
+            $obj = New-Object PSObject
+            $obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $fullName
+            $obj | Add-Member -MemberType NoteProperty -Name "ObjectClass" -Value $className
+            $obj | Add-Member -MemberType NoteProperty -Name "SID" -Value $sid
+            $obj | Add-Member -MemberType NoteProperty -Name "Source" -Value "ADSI"
+            $obj | Add-Member -MemberType NoteProperty -Name "IsUnresolved" -Value $isUnresolved
+            $results += $obj
         }
         catch {
             Write-Log ("    ADSI メンバー情報の取得に失敗（スキップ）: {0}" -f $_.Exception.Message) "WARN"
```

-----

## 差分⑨：アカウント種別判定ヘルパー

```diff
 function Get-AccountType {
     param(
         $Member,
         [string]$Method
     )

     if ($Method -eq "Cmdlet") {
         ...
     }

     $memberName = $Member.Name

-    if ([string]::IsNullOrWhiteSpace($memberName)) {
+    if (Test-StringNullOrWhiteSpace $memberName) {
         return "Unknown"
     }
```

-----

## 差分⑩：Get-MemberName / Get-MemberObjectClass

```diff
 function Get-MemberName {
     param($Member)

     $name = $Member.Name
-    if ([string]::IsNullOrWhiteSpace($name)) {
+    if (Test-StringNullOrWhiteSpace $name) {
         return "(名前取得不可)"
     }
     return $name
 }

 function Get-MemberObjectClass {
     param($Member)

     $class = $Member.ObjectClass
-    if ([string]::IsNullOrWhiteSpace($class)) {
+    if (Test-StringNullOrWhiteSpace $class) {
         return "(不明)"
     }
     return $class
 }
```

-----

## 差分⑪：Get-MemberSID

```diff
     else {
-        if (-not [string]::IsNullOrWhiteSpace($Member.SID)) {
+        if (-not (Test-StringNullOrWhiteSpace $Member.SID)) {
             return $Member.SID
         }
     }
```

-----

## 差分⑫：Get-MemberStatus

```diff
     if ($Method -eq "Cmdlet") {
         $rawName = $Member.Name
         if (Test-UnresolvedSID -Name $rawName) {
             return "Unresolved"
         }
         # Name が空でも SID 文字列から未解決を判定（フォールバック）
-        if ([string]::IsNullOrWhiteSpace($rawName) -and $null -ne $Member.SID) {
+        if ((Test-StringNullOrWhiteSpace $rawName) -and $null -ne $Member.SID) {
             try {
                 $Member.SID.Translate([System.Security.Principal.NTAccount]) | Out-Null
```

-----

## 差分⑬：パラメータ検証 — グループ空文字チェック

```diff
 foreach ($g in $TargetGroups) {
-    if ([string]::IsNullOrWhiteSpace($g)) {
+    if (Test-StringNullOrWhiteSpace $g) {
         Write-Log "対象グループに空文字が含まれています。スキップします。" "WARN"
     }
```

-----

## 差分⑭：除外アカウント検証 — 空文字エントリ除去

```diff
         $entries = @($ExcludeAccounts[$exGroup])
-        $validEntries = @($entries | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
+        $validEntries = @($entries | Where-Object { -not (Test-StringNullOrWhiteSpace $_) })
         $emptyCount = $entries.Count - $validEntries.Count
```

-----

## 差分⑮：CSV レコード — PSCustomObject → Add-Member

```diff
         $detected = ($status -ne "Unresolved" -and -not $isExcluded -and $DetectAccountTypes -contains $accountType)

-        $allCsvRecords += [PSCustomObject]@{
-            ComputerName    = $script:HostName
-            GroupName       = $group
-            AccountName     = $memberName
-            AccountType     = $accountType
-            ObjectClass     = $objectClass
-            SID             = $memberSID
-            Status          = $status
-            Detected        = $detected
-            RetrievalMethod = $method
-            CollectedAt     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
-        }
+        $rec = New-Object PSObject
+        $rec | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $script:HostName
+        $rec | Add-Member -MemberType NoteProperty -Name "GroupName" -Value $group
+        $rec | Add-Member -MemberType NoteProperty -Name "AccountName" -Value $memberName
+        $rec | Add-Member -MemberType NoteProperty -Name "AccountType" -Value $accountType
+        $rec | Add-Member -MemberType NoteProperty -Name "ObjectClass" -Value $objectClass
+        $rec | Add-Member -MemberType NoteProperty -Name "SID" -Value $memberSID
+        $rec | Add-Member -MemberType NoteProperty -Name "Status" -Value $status
+        $rec | Add-Member -MemberType NoteProperty -Name "Detected" -Value $detected
+        $rec | Add-Member -MemberType NoteProperty -Name "RetrievalMethod" -Value $method
+        $rec | Add-Member -MemberType NoteProperty -Name "CollectedAt" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
+        $allCsvRecords += $rec
     }
```

-----

## 差分⑯：CSV 出力 — Select-Object で列順保証

```diff
 # ============================================================
 # CSV 出力
 # ============================================================
+$csvColumns = @("ComputerName","GroupName","AccountName","AccountType","ObjectClass","SID","Status","Detected","RetrievalMethod","CollectedAt")
+
 try {
     if ($allCsvRecords.Count -gt 0) {
-        $allCsvRecords | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
+        $allCsvRecords | Select-Object $csvColumns | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
         Write-Log ""
         Write-Log ("CSV 出力完了 ({0} 件): {1}" -f $allCsvRecords.Count, $CsvFile) "SUCCESS"
     }
```

-----

## 差分まとめ

|#  |変更内容                                                             |PS 2.0 非対応の理由        |
|---|-----------------------------------------------------------------|---------------------|
|①  |.NOTES 対象 OS 変更                                                  |—                    |
|②  |`#Requires` 削除                                                   |PS 4.0+ 構文           |
|③  |`Test-StringNullOrWhiteSpace` 関数追加                               |.NET 4.0+ メソッド       |
|④  |`$PSScriptRoot` → `Split-Path`                                   |PS 3.0+ 自動変数         |
|⑤  |`$PSScriptRoot` → `$ScriptRoot` 参照変更                             |同上                   |
|⑥〜⑭|`IsNullOrWhiteSpace` → `Test-StringNullOrWhiteSpace`（10箇所）       |.NET 4.0+ メソッド       |
|⑮⑧ |`[PSCustomObject]@{}` → `New-Object PSObject` + `Add-Member`（2箇所）|PS 3.0+ 構文           |
|⑯  |`Select-Object $csvColumns` 追加                                   |`Add-Member` 方式での列順保証|
