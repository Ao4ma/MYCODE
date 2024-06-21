
# SSH接続のためのユーザー名、パスワード、サーバー名を定義します
$sshUser = "ycsvm103\administrator" # replace with your actual SSH username
$sshPassword = ConvertTo-SecureString "#YamadaVM03" -AsPlainText -Force # replace with your actual SSH password
$sshServer = "Ycsvm103" # replace with your actual server hostname or IP address

# SSHのユーザー名とパスワードを使用してPSCredentialオブジェクトを作成します
$sshCredential = New-Object System.Management.Automation.PSCredential ($sshUser, $sshPassword)

try {
    # 与えられた認証情報を使用して指定したサーバーにSSHセッションを作成します
    $sshSession = New-SSHSession -ComputerName $sshServer -Credential $sshCredential

    # リモートサーバー上でコードページをUTF-8（65001）に変更します
    $command = "chcp 65001"
    Invoke-SSHCommand -SSHSession $sshSession -Command $command

    # 'tasklist'コマンドをリモートで実行し、"rdp"と"dwm"を含むタスクのリストを取得します
    $command = "tasklist /v | findstr /i /c:""rdp"" /c:""dwm"""
    $taskList = Invoke-SSHCommand -SSHSession $sshSession -Command $command

    # タスクリストを出力します
    $taskList.Output

    # 'query session'コマンドをリモートで実行し、セッションのリストを取得します
    $command = "query session"
    $sessionList = Invoke-SSHCommand -SSHSession $sshSession -Command $command

    # セッションリストを出力します
    $sessionList.Output
} catch {
    # SSHコマンド実行の失敗を処理します
    Write-Error "SSHコマンドの実行に失敗しました: $_"
} finally {
    # SSHセッションが存在する場合は閉じます
    if ($sshSession) {
        Remove-SSHSession -SessionId $sshSession.SessionId
    }
}