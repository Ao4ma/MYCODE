# 出力エンコーディングをUTF-8に設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    # コンピュータの配列を定義
    $computers = @(
        @{Name="ycsvm103"; User="ycsvm103\administrator"; Password="#YamadaVM03"; Type="Server"},
        @{Name="Server2"; User="ServerUser2"; Password="ServerPassword2"; Type="Server"},
        @{Name="Server3"; User="ServerUser3"; Password="ServerPassword3"; Type="Server"},
        @{Name="delld022"; User="ygijutubu"; Password="YCg-7741315-4"; Type="LicensePC"},
        @{Name="LicensePC2"; User="LicensePCUser2"; Password="LicensePCPassword2"; Type="LicensePC"},
        @{Name="LicensePC3"; User="LicensePCUser3"; Password="LicensePCPassword3"; Type="LicensePC"}
    )

    # ClientPCクラスの定義
    class ClientPC {
        [string] $Type
        [array] $AccessibleComputers

        ClientPC([string] $type, [array] $computers) {
            $this.Type = $type
            if ($type -eq "Admin") {
                $this.AccessibleComputers = $computers
            } else {
                $this.AccessibleComputers = $computers | Where-Object {$_.Type -eq "LicensePC"}
            }
        }
    }

    # ClientPCクラスのインスタンスを作成
    $adminPC = [ClientPC]::new("Admin", $computers)

    # 選択したコンピュータ名の変数
    $selectedComputerName = "ycsvm103"

    # アクセス可能なコンピュータから選択したコンピュータを取得
    $selectedComputer = $adminPC.AccessibleComputers | Where-Object {$_.Name -eq $selectedComputerName}

    # 選択したコンピュータをファイルに保存
    $selectedComputer | Out-File -FilePath 'output.txt' -Encoding 'UTF8'

    # 自己署名証明書を作成
    $cert = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "cert:\CurrentUser\My" -KeyUsage DigitalSignature -TextExtension @("2.5.29.37={text}1.3.6.1.5.5.7.3.3")
    # ファイルパスを自動で取得
    $scriptFilePath = $MyInvocation.MyCommand.Path

    # スクリプトにデジタル署名を付ける
    Set-AuthenticodeSignature -FilePath $scriptFilePath -Certificate $cert

    # Computerクラスの定義
    class Computer {
        [string] CheckRdpSession() {
            return "RDPセッションのチェック中..."
        }

        [string] GetSessionList() {
            return "セッションリストの取得中..."
        }
    }

    # Serverクラスの定義
    class Server : Computer {
        # Serverクラス固有のプロパティやメソッドをここに追加できます
    }

    # LicensePCクラスの定義
    class LicensePC : Computer {
        # LicensePCクラス固有のプロパティやメソッドをここに追加できます
    }

    # 選択したコンピュータのタイプに基づいて適切なコンピュータクラスのインスタンスを作成
    if ($selectedComputer.Type -eq "Server") {
        $computer = [Server]::new()
    } else {
        $computer = [LicensePC]::new()
    }

    # コンピュータオブジェクトのメソッドを呼び出す
    $computer.CheckRdpSession()
    $computer.GetSessionList()
}
catch {
    Write-Host "An error occurred: $_"
}

# SIG # Begin signature block
# MIIFlgYJKoZIhvcNAQcCoIIFhzCCBYMCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDltXQPWLEwa/aj
# y5KsS7SR9FNFcpyGbRAx7vNkFmNK1KCCAxIwggMOMIIB9qADAgECAhB6bP6sxNa7
# skL43Vz0ZaqsMA0GCSqGSIb3DQEBCwUAMBQxEjAQBgNVBAMMCWxvY2FsaG9zdDAe
# Fw0yNDA2MjExNDI3NDJaFw0yNTA2MjExNDQ3NDJaMBQxEjAQBgNVBAMMCWxvY2Fs
# aG9zdDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALsFtsSlCMHyQrWS
# AH2vQ7H9ptAL/SPKbjBWwU3DvXvkcL8VVWAtti9OIKjDtb4ztZ8CG6yxtnAzahtZ
# SD38rGwxhrksigHLNXHc4dI23uBscJtUfwAQ9WI/Cy+9fi5sg6feKB2pB8I4galb
# WAtVGkuqCoPh6kEqtgZeZ5MgxrGPEapga55Meu4EeGEYemxbGsGA2s4dRaDLK7rA
# yPbVMLDRAS2pI5fZnr6qPqYP3swALEa1RdVm8A16fxsnLdMHk+A6j9hXyuUmnDsf
# wb7fzmMcpSZCffjosoxLhyCay5R3Wek1qbbuKV9nCl/ptep81T2J1ZdVJjLg7zTz
# 6oq7HmUCAwEAAaNcMFowDgYDVR0PAQH/BAQDAgeAMBQGA1UdEQQNMAuCCWxvY2Fs
# aG9zdDATBgNVHSUEDDAKBggrBgEFBQcDAzAdBgNVHQ4EFgQUZESab+CMZnPGpi7o
# +fYlPkLVM4IwDQYJKoZIhvcNAQELBQADggEBAITY2aIS9KKCfL3hQlit9/TkxbjG
# IoPEM69hxlAffPtCmAa6gfy/vVAzlFYTHwlUAUEKUNNSffU4un6ansmWCeRsoqGL
# H0OIL7rSTf5owApkMWMgkhX7rajzsuKDzlMVSjutwWdBidTbzvcsiSgz+lCdYKEh
# ph1Be2kzMSxkuvUubR925SXE7hWVYogAbrE97n4FUko5iM6ncbbld0Vccmv2Zr4N
# F1D8DizL/eaHM7LdZOPDrP7XJiwU3oU1+blm0WKoY5blld8hhiDhCa2sCoyeipUx
# hk5zvn7JhomB/g2GK/hr1kukKaO8EdE0xXsNFncWMrbrQCa4H7Ng1MIIOhcxggHa
# MIIB1gIBATAoMBQxEjAQBgNVBAMMCWxvY2FsaG9zdAIQemz+rMTWu7JC+N1c9GWq
# rDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkG
# CSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEE
# AYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAIKteE66NlcNnEQUBcBJSbhaMREiQ/FxeI
# MKrfnsnRQzANBgkqhkiG9w0BAQEFAASCAQB/J7zdg/R5DHzz4Fr98HqYEDnEy/B7
# LlglOaEm+gKff8SDzVTFNk4wV0rw8YXzg4Z0SQ2lVh/vuGXLUyUoT5P/Bx/OJy3N
# AVnj8WWZm0FXqVYASW4RLhkmQAU/7Oy50CdjCoFt7jWTG8Sl8lSZjHNlh1zwEQiK
# zREYbXUXhkGuNuNRrWWeUw93ZHBTvDg2mWjmO+uqDpw4AHR0cfba9S7rdaf2/Fsz
# N6lk3CzUcBtHpot+zdg2Jg4IJV4S+E8TXt/vdweD99OgRnSwdN/ewxi9wk1EC1oO
# 7CrWXBt9attOl7teZhT9dQok6VZj+VUgz3vFzKMoffU5XLo27CQM52lL
# SIG # End signature block
