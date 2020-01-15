# ==============================================================================
# 処理内容
# ------------------------------------------------------------------------------
<#
   指定フォルダのファイルをFTPでアップロードする。
   - FTPに使用するユーザーIDとパスワードは初回に入力し、記録する
     - ファイルで記録する。
     - パスワードはwindows資格情報で暗号化する（同一アカウントでしか復元できない仕様）
     - 2回目以降はファイルからIDとパスワードを取得し実行する
   - 指定フォルダ内のファイルをFTPアップロード
     - 対象のファイルは、指定フォルダ内の指定ファイル名ルールに合致するもの
     - サーバーにファイルがある場合は、アップロード前にファイルをバックアップ
     - アップロード済みファイルは、処理済みとして指定のフォルダに移動
#>

# 未定義変数参照をエラー化
set-psdebug -strict

# 設定
$LOCAL_FOLDER  = "C:\temp"                     # LOCALアップロードファイル
$LOCAL_FILEFORMAT = "^[0-9]{8}.xml"            # LOCALファイルフォーマット
$BACKUP_FOLDER = "$LOCAL_FOLDER\beforeUpload"  # アップロード前のバックアップフォルダ
$DONE_FOLDER   = "$LOCAL_FOLDER\Uploaded"      # アップロード処理済みフォルダ
$CRED_FOLDER   = "C:\temp"                     # FTP認証情報フォルダ
$FTP_URL       = "ftp://<FTPのURL>"            # FTP_URL
$FTP_UP_FOLDER = "$2f/<FTPアップ先フォルダ>"   # FTPアップロードフォルダ（$2fはRootを示す）

# ==============================================================================
# 関数定義
# ------------------------------------------------------------------------------
#  名前：Write-Log
#  概要：ログに書き込む
# ------------------------------------------------------------------------------
Function Write-Log{
	Param(
        [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
        [string]$text
	)
	Process {
		$time = Get-Date -format yyyy/MM/dd-HH:mm:ss
    	Write-Host "$text"
        $time + " " + $text -Replace("`r`n", "`r`n$time ") | Out-File $logFilePath -Append
    }
}

# ==============================================================================
# メイン処理
# ------------------------------------------------------------------------------
try{
    # 初期処理
    $procName     = ($MyInvocation.MyCommand.Name).Replace(".ps1", "") # 処理名
    $userName     = Get-Content env:username                           # ユーザ名
    $credFileName = "$procName`_$userName.txt"                   # FTP認証情報ファイル
    $logFilePath  = "$LOCAL_FOLDER\$procName`_$userName.log"     # ログファイルパス
    Write-Log ("=" * 70 + "`r`n$procName Proc-Start`r`n" + "=" * 70) 

    # FTP認証情報ファイルが無いなら作成する
    if ((Test-Path $CRED_FOLDER) -eq $false){
        throw "FTP認証情報を保存するフォルダを作成して下さい：$CRED_FOLDER"
    }else{
        $credFilePath = "$CRED_FOLDER\$credFileName"
        if ((Test-Path $credFilePath) -eq $false){
            Write-Log "Create Credential File. $credFilePath"
            Write-Host "--- FTP認証情報の設定 ----------------------------------"
            Write-Host " 3秒後にダイアログが表示されます。"
            Write-Host " 使用するFTPアカウントのIDとパスワードを入力して下さい。"
            Write-Host "--------------------------------------------------------"
            for($i=3; $i-gt0; $i--){ Write-Host " $i..."; Start-Sleep -Seconds 1 }
            $cred     = Get-Credential
            $secPw    = ConvertFrom-SecureString $cred.Password
            $credHash = @{userName = $cred.UserName; password = $secPw;}
            $credHash | ConvertTo-Json | Out-File $credFilePath -Encoding utf8
        }
    }

    # FTP認証情報ファイルを読み込み
    Write-Log "[Load credential file]`r`n$credFilePath"
    $credSec  = Get-Content $credFilePath -Encoding UTF8 -Raw | ConvertFrom-Json
    $pwSec    = $credSec.password | ConvertTo-SecureString
    $secStr   = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pwSec)
    $pw       = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($secStr)
    $ftpCred  = @{userName = $credSec.UserName; password = $pw;}

    # アップロード対象フォルダのファイルを取得
    Write-Log "[Check upload file]`r`n$LOCAL_FOLDER"
    $localFiles = Get-ChildItem $LOCAL_FOLDER | Where-Object { ! $_.PSIsContainer -and 
                                                                ($_.Name) -match $LOCAL_FILEFORMAT}
    $upFiles = @()
    ForEach($lf in $localFiles){ 
        $nowYmdStr = (Get-Date).ToString("yyyyMMddHHmmss")
        $upFiles += @{serverPath = $FTP_UP_FOLDER + $lf.Name;
                      localPath  = $lf.FullName;
                      backupPath = $BACKUP_FOLDER + "\" + $lf.Name + "_$nowYmdStr.bak";
                      donePath   = $DONE_FOLDER   + "\" + $lf.Name + "_$nowYmdStr.bak"; }
    }

    # 対象ファイルがあれば処理続行
    if ($upfiles.Length -eq 0){
        Write-Log "no file"
    }else{
        Write-Log ([String]$upfiles.Length +" files exist")
        # FTP接続
        Write-Log ("[Connect ftp server]`r`n$FTP_URL " + $ftpCred.userName)
        $wc = New-Object System.Net.WebClient
        $wc.Credentials = New-Object System.Net.NetworkCredential($ftpCred.userName, $ftpCred.password)
        $wc.BaseAddress = $FTP_URL

        # ファイルループ
        ForEach($uf in $upfiles){ 
            # 1）サーバー側ファイルをバックアップ
            Write-Log ("[Backup server file]`r`nfrom:" + $uf.serverPath + "`r`n  to:" + $uf.backupPath)
            try {
                $wc.DownloadFile($uf.serverPath, $uf.backupPath)
            }catch [Exception]{
                if (([String]$_.Exception).Contains("ログインされていません")){
                    throw ("FtpConnectError:" + [String]$_.Exception)
                }elseif (([String]$_.Exception).Contains("ファイルが見つからない")){
                    Write-Log " --> This file is not found in ftp server."
                }else{
                    throw ("FtpConnectError:" + [String]$_.Exception)
                }
            }
            # 2）アップ
            Write-Log("[Upload file]`r`nfrom:" + $uf.localPath + "`r`n  to:" + $uf.serverPath)
            try {
                $wc.UploadFile($uf.serverPath, $uf.localPath)
                Start-Sleep -s 1 # 1秒WAIT
            }catch [Exception]{
                throw ("FtpLoadError:" + [String]$_.Exception)
            }
            # 3）アップ済みファイルは移動
            Write-Log("[Move files]`r`nfrom:" + $uf.localPath + "`r`n  to:" + $uf.donePath)
            try {
                Move-Item $uf.localPath $uf.donePath
            }catch [Exception]{
                throw ("BackupError:" + [String]$_.Exception)
            }
        }
    }
}catch [Exception]{
    Write-Log $error[0]
}finally{
    Write-Log ("-" * 70 + "`r`n$procName Proc-End`r`n") 
}
#Start-Sleep -s 3
