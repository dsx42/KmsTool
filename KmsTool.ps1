function RequireAdmin {
    $CurrentWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $CurrentWindowsPrincipal = New-Object -TypeName System.Security.Principal.WindowsPrincipal `
        -ArgumentList $CurrentWindowsID
    $Admin = $CurrentWindowsPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    if (!$Admin) {
        Start-Process -FilePath PowerShell.exe -ArgumentList `
            "-NoProfile -ExecutionPolicy RemoteSigned -File `"$PSCommandPath`" $PSBoundParameters" -Verb RunAs `
            -WindowStyle Normal
        [System.Environment]::Exit(0)
    }
}

function GetVertion {
    $ProductJsonPath = "$PSScriptRoot\product.json"

    if (!(Test-Path -Path $ProductJsonPath -PathType Leaf)) {
        Write-Warning -Message ("$ProductJsonPath 不存在")
        [System.Environment]::Exit(0)
    }

    $ProductInfo = $null
    try {
        $ProductInfo = Get-Content -Path $ProductJsonPath | ConvertFrom-Json
    }
    catch {
        Write-Warning -Message ("$ProductJsonPath 解析失败")
        [System.Environment]::Exit(0)
    }
    if (!$ProductInfo -or $ProductInfo -isNot [PSCustomObject]) {
        Write-Warning -Message ("$ProductJsonPath 解析失败")
        [System.Environment]::Exit(0)
    }

    $Version = $ProductInfo.'version'
    if (!$Version) {
        Write-Warning -Message ("$ProductJsonPath 不存在 version 信息")
        [System.Environment]::Exit(0)
    }

    return $Version
}

function GetOfficeActiveInfo {
    param($OsppPath)

    $OfficeActiveInfo = @{
        'Name'          = '';
        'IsVolume'      = $false;
        'KmsHost'       = '';
        'IsActive'      = $false;
        'LeftMinutes'   = 0; # 激活有效时间，单位为分钟
        'ActiveEndTime' = ''
    }

    $Results = CScript //Nologo "$OsppPath" /dstatus
    if (!$Results) {
        return $OfficeActiveInfo
    }

    foreach ($Result in $Results) {
        if (!$Result) {
            continue
        }

        if ($Result.Contains('LICENSE NAME: ')) {
            if ($Result.Contains('Office21Project')) {
                $OldName = $OfficeActiveInfo['Name']
                if (!$OldName) {
                    $OfficeActiveInfo['Name'] = 'Project 2021'
                }
                else {
                    $OfficeActiveInfo['Name'] = $OldName + ', Project 2021'
                }
            }
            elseif ($Result.Contains('Office21Visio')) {
                $OldName = $OfficeActiveInfo['Name']
                if (!$OldName) {
                    $OfficeActiveInfo['Name'] = 'Visio 2021'
                }
                else {
                    $OfficeActiveInfo['Name'] = $OldName + ', Visio 2021'
                }
            }
            elseif ($Result.Contains('Office21ProPlus') -or $Result.Contains('Office21Standard')) {
                $OldName = $OfficeActiveInfo['Office']
                if (!$OldName) {
                    $OfficeActiveInfo['Name'] = 'Office 2021'
                }
                else {
                    $OfficeActiveInfo['Name'] = $OldName + ', Office 2021'
                }
            }
            continue
        }

        if ($Result.Contains('LICENSE DESCRIPTION: ') -and $Result.Contains('KMS')) {
            $OfficeActiveInfo['IsVolume'] = $true
            continue
        }

        if ($Result.Contains('KMS machine registry override defined: ')) {
            $Entries = $Result.Split(':', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 2) {
                $OfficeActiveInfo['KmsHost'] = ''
                continue
            }
            $OfficeActiveInfo['KmsHost'] = $Entries[1].Trim()
            continue
        }

        if ($Result.Contains('LICENSE STATUS: ') -and $Result.Contains('---LICENSED---')) {
            $OfficeActiveInfo['IsActive'] = $true
            continue
        }

        if ($Result.Contains('REMAINING GRACE: ')) {
            $Entries = $Result.Split('(', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 2) {
                $OfficeActiveInfo['LeftMinutes'] = 0
                $OfficeActiveInfo['ActiveEndTime'] = ''
                continue
            }

            $SubEntries = $null
            if ($Result.Contains('minute')) {
                $SubEntries = $Entries[1].Split('minute', [System.StringSplitOptions]::RemoveEmptyEntries)
            }
            if (!$SubEntries -or $SubEntries.Length -lt 1) {
                $OfficeActiveInfo['LeftMinutes'] = 0
                $OfficeActiveInfo['ActiveEndTime'] = ''
                continue
            }

            $OfficeActiveInfo['LeftMinutes'] = [System.Int32]$SubEntries[0].Trim()
            $OfficeActiveInfo['ActiveEndTime'] = (Get-Date).AddMinutes([System.Int32]$SubEntries[0].Trim()).`
                ToString('yyyy-MM-dd HH:mm:ss')
        }
    }

    return $OfficeActiveInfo
}

function GetWindowsActiveInfo {

    $WindowsActiveInfo = @{
        'IsVolume'      = $false;
        'KmsIp'         = '';
        'KmsHost'       = '';
        'IsActive'      = $false;
        'TailKey'       = '';
        'LeftMinutes'   = 0; # 激活有效时间，单位为分钟
        'ActiveEndTime' = ''
    }

    $Results = CScript //Nologo "$env:windir\System32\slmgr.vbs" /dli
    if (!$Results) {
        return $WindowsActiveInfo
    }

    foreach ($Result in $Results) {
        if (!$Result) {
            continue
        }
        if ($Result.Contains('VOLUME_KMSCLIENT') -or $Result.Contains('VOLUME_KMS')) {
            $WindowsActiveInfo['IsVolume'] = $true
            continue
        }
        if ($Result.Contains('KMS 计算机 IP 地址: ') -or $Result.Contains('KMS machine IP address: ')) {
            if ($Result.Contains('不可用') -or $Result.Contains('not available')) {
                $WindowsActiveInfo['KmsIp'] = ''
                continue
            }
            $Entries = $Result.Split(':', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 2) {
                $WindowsActiveInfo['KmsIp'] = ''
                continue
            }
            $WindowsActiveInfo['KmsIp'] = $Entries[1].Trim()
            continue
        }
        if ($Result.Contains('已注册的 KMS 计算机名称: ') -or $Result.Contains('Registered KMS machine name: ') `
                -or $Result.Contains('来自 DNS 的 KMS 计算机名称: ') `
                -or $Result.Contains('KMS machine name from DNS: ')) {
            if ($Result.Contains('KMS 名称不可用') -or $Result.Contains('KMS name not available')) {
                $WindowsActiveInfo['KmsHost'] = ''
                continue
            }
            $Entries = $Result.Split(':', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 3) {
                $WindowsActiveInfo['KmsHost'] = ''
                continue
            }
            $WindowsActiveInfo['KmsHost'] = $Entries[1].Trim()
            continue
        }
        if ($Result.Contains('许可证状态: ') -or $Result.Contains('License Status: ')) {
            if ($Result.Contains('已授权') -or $Result.COntains('Licensed')) {
                $WindowsActiveInfo['IsActive'] = $true
            }
            else {
                $WindowsActiveInfo['IsActive'] = $false
            }
            continue
        }
        if ($Result.Contains('部分产品密钥: ') -or $Result.Contains('Partial Product Key: ')) {
            $Entries = $Result.Split(':', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 2) {
                $WindowsActiveInfo['TailKey'] = ''
                continue
            }
            $WindowsActiveInfo['TailKey'] = $Entries[1].Trim()
            continue
        }
        if ($Result.Contains('激活过期: ') -or $Result.Contains('activation expiration: ')) {

            $Entries = $Result.Split(':', [System.StringSplitOptions]::RemoveEmptyEntries)
            if (!$Entries -or $Entries.Length -lt 2) {
                $WindowsActiveInfo['LeftMinutes'] = 0
                $WindowsActiveInfo['ActiveEndTime'] = ''
                continue
            }

            $SubEntries = $null
            if ($Result.Contains('分钟')) {
                $SubEntries = $Entries[1].Split('分钟', [System.StringSplitOptions]::RemoveEmptyEntries)
            }
            elseif ($Result.Contains('minute')) {
                $SubEntries = $Entries[1].Split('minute', [System.StringSplitOptions]::RemoveEmptyEntries)
            }
            if (!$SubEntries -or $SubEntries.Length -lt 1) {
                $WindowsActiveInfo['LeftMinutes'] = 0
                $WindowsActiveInfo['ActiveEndTime'] = ''
                continue
            }

            $WindowsActiveInfo['LeftMinutes'] = [System.Int32]$SubEntries[0].Trim()
            $WindowsActiveInfo['ActiveEndTime'] = (Get-Date).AddMinutes([System.Int32]$SubEntries[0].Trim()).`
                ToString('yyyy-MM-dd HH:mm:ss')
        }
    }

    return $WindowsActiveInfo
}

function GetSelectWindowsProduction {

    $SystemInfo = Get-CimInstance -ClassName Win32_OperatingSystem

    Write-Host -Object ''
    Write-Host -Object "不支持激活 $($SystemInfo.Caption), 需要转换为如下批量授权版本才能激活"
    Write-Host -Object ''
    Write-Host -Object '1: 企业版 Enterprise'
    Write-Host -Object ''
    Write-Host -Object '2: 教育版 Education'
    Write-Host -Object ''
    Write-Host -Object '3: 专业版 Pro'
    Write-Host -Object ''
    Write-Host -Object '4: 专业教育版 Pro Education'
    Write-Host -Object ''
    Write-Host -Object '5: 专业工作站版 Pro For Workstations'
    Write-Host -Object ''
    Write-Host -Object '0: 不激活'

    while ($true) {
        Write-Host -Object ''
        $InputOption = Read-Host -Prompt '请输入选择的序号，按回车键确认'
        if ($null -eq $InputOption -or '' -eq $InputOption) {
            Write-Host -Object ''
            Write-Warning -Message '选择无效，请重新输入'
            continue
        }
        if ('0' -ieq $InputOption) {
            return ''
        }
        if ('1' -ieq $InputOption) {
            return 'Enterprise'
        }
        if ('2' -ieq $InputOption) {
            return 'Education'
        }
        if ('3' -ieq $InputOption) {
            return 'Pro'
        }
        if ('4' -ieq $InputOption) {
            return 'Pro Education'
        }
        if ('5' -ieq $InputOption) {
            return 'Pro For Workstations'
        }
        Write-Host -Object ''
        Write-Warning -Message '选择无效，请重新输入'
    }
}

function GetWindowsGvlk {
    param($WindowsActiveInfo)

    $WindowsProducts = [ordered]@{
        'Enterprise'           = @{
            'CN'   = '企业版';
            'US'   = 'Enterprise';
            'gvlk' = 'NPPR9-FWDCX-D2C8J-H872K-2YT43'
        };
        'Education'            = @{
            'CN'   = '教育版';
            'US'   = 'Education';
            'gvlk' = 'NW6C2-QMPVW-D7KKK-3GKT6-VCFB2'
        };
        'Pro'                  = @{
            'CN'   = '专业版';
            'US'   = 'Pro';
            'gvlk' = 'W269N-WFGWX-YVC9B-4J6C9-T83GX'
        };
        'Pro Education'        = @{
            'CN'   = '专业教育版';
            'US'   = 'Pro Education';
            'gvlk' = '6TP4R-GNPTD-KYYHQ-7B7DP-J447Y'
        };
        'Pro For Workstations' = @{
            'CN'   = '专业工作站版';
            'US'   = 'Pro For Workstations';
            'gvlk' = 'NRG8B-VKK3Q-CXVCJ-9G2XF-6Q84J'
        }
    }

    $Edition = (Get-WindowsEdition -Online).Edition

    if (!$WindowsProducts.Contains($Edition)) {
        $Edition = GetSelectWindowsProduction
    }
    if (!$WindowsProducts.Contains($Edition)) {
        return $null
    }

    $CurrentProduct = $WindowsProducts[$Edition]
    $Gvlk = $CurrentProduct['gvlk']
    if ($null -eq $WindowsActiveInfo['TailKey'] -or '' -eq $WindowsActiveInfo['TailKey'] `
            -or !$Gvlk.Contains($WindowsActiveInfo['TailKey']) -or !$WindowsActiveInfo['IsVolume']) {
        return $Gvlk
    }

    return ''
}

function TestKms {
    param ($KmsHost, $OsppPath)

    $IsValid = Test-NetConnection -ComputerName $KmsHost -Port 1688 -InformationLevel Quiet `
        -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    if (!$IsValid) {
        Write-Host -Object ''
        Write-Warning -Message "KMS 激活服务器 $KmsHost 不可用"
        return $false
    }

    if (!$OsppPath) {
        CScript //Nologo "$env:windir\System32\slmgr.vbs" /skms $KmsHost | Out-Null
        $Results = CScript //Nologo "$env:windir\System32\slmgr.vbs" /ato
        if (!$Results) {
            Write-Host -Object ''
            Write-Warning -Message "系统所在网络无法访问此 KMS 激活服务器 $KmsHost"
            return $false
        }

        foreach ($Result in $Results) {
            if (!$Result) {
                continue
            }
            if ($Result.Contains('成功地激活了产品') -or $Result.Contains('Product activated successfully')) {
                return $true
            }
        }

        Write-Host -Object ''
        Write-Warning -Message "系统所在网络无法访问此 KMS 激活服务器 $KmsHost"
        return $false
    }

    CScript //Nologo "$OsppPath" /sethst:$KmsHost | Out-Null
    $Results = CScript //Nologo "$OsppPath" /act
    if (!$Results) {
        Write-Host -Object ''
        Write-Warning -Message "系统所在网络无法访问此 KMS 激活服务器 $KmsHost"
        return $false
    }

    foreach ($Result in $Results) {
        if (!$Result) {
            continue
        }
        if ($Result.Contains('Product activation successful') `
                -or $Result.Contains('Offline product activation successful')) {
            return $true
        }
    }

    Write-Host -Object ''
    Write-Warning -Message "系统所在网络无法访问此 KMS 激活服务器 $KmsHost"
    return $false
}

function GetValidKmsServer {
    param($KmsHost, $KmsIp, $OsppPath)

    $NeedTestKmsServer = @()
    if ($KmsHost) {
        $NeedTestKmsServer += $KmsHost
    }
    if (!$KmsHost -and $KmsIp) {
        $NeedTestKmsServer += $KmsIp
    }
    foreach ($Kms in $Script:kmsServers) {
        $NeedTestKmsServer += $Kms
    }

    foreach ($Kms in $NeedTestKmsServer) {

        $Valid = TestKms -KmsHost $Kms -OsppPath $OsppPath
        if ($Valid) {
            return $Kms
        }
    }

    while ($true) {
        Write-Host -Object ''
        $InputOption = Read-Host -Prompt ('无可用的 KMS 激活服务器, 请输入可用的 KMS 激活服务器域名或 IP ' `
                + '(0 表示退出激活), 按回车键确认')
        if ($null -eq $InputOption -or '' -eq $InputOption) {
            Write-Host -Object ''
            Write-Warning -Message '输入无效，请重新输入'
            continue
        }
        if ('0' -ieq $InputOption) {
            return $null
        }
        $Valid = TestKms -KmsHost $InputOption -OsppPath $OsppPath
        if ($Valid) {
            return $InputOption
        }
    }
}

function ConfirmOfficeProducts {
    param ($NeedOfficeProducts, $OfficeProducts)

    Write-Host -Object ''
    if ($NeedOfficeProducts.Count -le 0) {
        Write-Host -Object '未选择安装任何 Office 2021 组件'
        while ($true) {
            Write-Host -Object ''
            $InputOption = Read-Host -Prompt '请选择 (0: 退出安装; 2: 重置所有选择), 按回车键确认'
            if ($null -eq $InputOption -or '' -eq $InputOption) {
                Write-Host -Object ''
                Write-Warning -Message '选择无效，请重新输入'
                continue
            }
            if ('0' -ieq $InputOption) {
                return 0
            }
            if ('2' -ieq $InputOption) {
                return 2
            }
            Write-Host -Object ''
            Write-Warning -Message '选择无效，请重新输入'
        }
    }

    Write-Host -Object '选择安装的 Office 2021 组件如下:'
    foreach ($Product in $NeedOfficeProducts.GetEnumerator()) {
        if ($Product.Value) {
            Write-Host -Object ''
            Write-Host -Object $OfficeProducts[$Product.Key]
        }
    }
    Write-Host -Object ''
    Write-Host -Object '注意：会卸载当前系统所有已安装的 Office 组件，重新安装上述组件'
    while ($true) {
        Write-Host -Object ''
        $InputOption = Read-Host -Prompt '请选择 (0: 退出安装; 1: 继续安装; 2: 重置所有选择), 按回车键确认'
        if ($null -eq $InputOption -or '' -eq $InputOption) {
            Write-Host -Object ''
            Write-Warning -Message '选择无效，请重新输入'
            continue
        }
        if ('0' -ieq $InputOption) {
            return 0
        }
        if ('1' -ieq $InputOption) {
            return 1
        }
        if ('2' -ieq $InputOption) {
            return 2
        }
        Write-Host -Object ''
        Write-Warning -Message '选择无效，请重新输入'
    }
}

function GetOfficeProductSelect {
    param($ProductName)

    while ($true) {
        Write-Host -Object ''
        $InputOption = Read-Host -Prompt "是否需要安装 $ProductName (0: 不安装; 1: 安装; 2: 重置所有选择), 按回车键确认"
        if ($null -eq $InputOption -or '' -eq $InputOption) {
            Write-Host -Object ''
            Write-Warning -Message '选择无效，请重新输入'
            continue
        }
        if ('0' -ieq $InputOption) {
            return 0
        }
        if ('1' -ieq $InputOption) {
            return 1
        }
        if ('2' -ieq $InputOption) {
            return 2
        }
        Write-Host -Object ''
        Write-Warning -Message '选择无效，请重新输入'
    }
}

function AddSubElement {
    param ($NeedOfficeProducts)

    Add-Content -Path configuration.xml -Value '                <Language ID="MatchOS" />'
    Add-Content -Path configuration.xml -Value '                <Language ID="MatchPreviousMSI" />'
    if (!$NeedOfficeProducts.Contains('Access')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Access" />'
    }
    if (!$NeedOfficeProducts.Contains('Excel')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Excel" />'
    }
    if (!$NeedOfficeProducts.Contains('Lync')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Lync" />'
    }
    if (!$NeedOfficeProducts.Contains('OneDrive')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="OneDrive" />'
    }
    if (!$NeedOfficeProducts.Contains('OneNote')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="OneNote" />'
    }
    if (!$NeedOfficeProducts.Contains('Outlook')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Outlook" />'
    }
    if (!$NeedOfficeProducts.Contains('PowerPoint')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="PowerPoint" />'
    }
    if (!$NeedOfficeProducts.Contains('Publisher')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Publisher" />'
    }
    if (!$NeedOfficeProducts.Contains('Teams')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Teams" />'
    }
    if (!$NeedOfficeProducts.Contains('Word')) {
        Add-Content -Path configuration.xml -Value '                <ExcludeApp ID="Word" />'
    }
}

function CreateOfficeDeploymentFile {
    param ($NeedOfficeProducts)

    Set-Location -Path "$PSScriptRoot"
    if (Test-Path -Path configuration.xml -PathType Leaf) {
        Remove-Item -Path configuration.xml -Force
    }

    $OfficeClientEdition = '64'
    $SystemInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    if (!$SystemInfo.OSArchitecture.Contains('64')) {
        $OfficeClientEdition = '32'
    }

    Add-Content -Path configuration.xml -Value '<Configuration>'
    Add-Content -Path configuration.xml -Value ("    <Add OfficeClientEdition=`"$OfficeClientEdition`"" `
            + " Channel=`"PerpetualVL2021`" MigrateArch=`"TRUE`">")
    Add-Content -Path configuration.xml -Value ('        <Product ID="ProPlus2021Volume"' `
            + ' PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">')
    AddSubElement -NeedOfficeProducts $NeedOfficeProducts
    Add-Content -Path configuration.xml -Value '        </Product>'
    if ($NeedOfficeProducts.Contains('Visio')) {
        Add-Content -Path configuration.xml -Value ''
        Add-Content -Path configuration.xml -Value ('        <Product ID="VisioPro2021Volume"' `
                + ' PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">')
        AddSubElement -NeedOfficeProducts $NeedOfficeProducts
        Add-Content -Path configuration.xml -Value '        </Product>'
    }
    if ($NeedOfficeProducts.Contains('Project')) {
        Add-Content -Path configuration.xml -Value ''
        Add-Content -Path configuration.xml -Value ('        <Product ID="ProjectPro2021Volume"' `
                + ' PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">')
        AddSubElement -NeedOfficeProducts $NeedOfficeProducts
        Add-Content -Path configuration.xml -Value '        </Product>'
    }
    Add-Content -Path configuration.xml -Value '    </Add>'
    Add-Content -Path configuration.xml -Value ''
    Add-Content -Path configuration.xml -Value '    <Property Name="SharedComputerLicensing" Value="0" />'
    Add-Content -Path configuration.xml -Value '    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'
    Add-Content -Path configuration.xml -Value '    <Property Name="DeviceBasedLicensing" Value="0" />'
    Add-Content -Path configuration.xml -Value '    <Property Name="SCLCacheOverride" Value="0" />'
    Add-Content -Path configuration.xml -Value '    <Property Name="AUTOACTIVATE" Value="1" />'
    Add-Content -Path configuration.xml -Value '    <Updates Enabled="TRUE" />'
    Add-Content -Path configuration.xml -Value '    <RemoveMSI />'
    Add-Content -Path configuration.xml -Value '    <AppSettings>'
    Add-Content -Path configuration.xml -Value ('        <User Key="software\microsoft\office\16.0\common" ' `
            + 'Name="sendcustomerdata" Value="0" Type="REG_DWORD" App="office16" Id="L_Sendcustomerdata" />')
    Add-Content -Path configuration.xml -Value ('        <User Key="software\microsoft\office\16.0\common\graphics" ' `
            + 'Name="disablehardwareacceleration" Value="1" Type="REG_DWORD" App="office16" ' `
            + 'Id="L_DoNotUseHardwareAcceleration" />')
    Add-Content -Path configuration.xml -Value '    </AppSettings>'
    Add-Content -Path configuration.xml -Value '    <Display Level="None" AcceptEULA="TRUE" />'
    Add-Content -Path configuration.xml -Value '</Configuration>'
}

function ActiveOffice {

    Clear-Host
    Write-Host -Object ''
    Write-Host -Object '正在检测是否安装 Office 2021 批量授权版，请勿关闭此窗口'

    $OsppPath = ''
    if (Test-Path -Path "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS" -PathType Leaf) {
        $OsppPath = "$env:ProgramFiles\Microsoft Office\Office16\OSPP.VBS"
    }
    elseif (Test-Path -Path "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS" -PathType Leaf) {
        $OsppPath = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OSPP.VBS"
    }
    if (!$OsppPath) {
        Write-Host -Object ''
        Write-Warning -Message 'OSPP.VBS 文件不存在，未安装 Office 2021 批量授权版，无法激活'
        return
    }

    $OfficeActiveInfo = GetOfficeActiveInfo -OsppPath $OsppPath
    if (!$OfficeActiveInfo['IsVolume']) {
        Write-Host -Object ''
        Write-Warning -Message '未安装 Office 2021 批量授权版，无法激活'
        return
    }

    Write-Host -Object ''
    if ($OfficeActiveInfo['IsActive']) {
        Write-Host -Object ($OfficeActiveInfo['Name'] + ' 已激活, 激活有效期至 ' + $OfficeActiveInfo['ActiveEndTime'])
    }
    elseif ($OfficeActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($OfficeActiveInfo['Name'] + ' 激活有效期至 ' + $OfficeActiveInfo['ActiveEndTime'])
    }
    else {
        Write-Host -Object ($OfficeActiveInfo['Name'] + ' 未激活')
    }

    $ValidKms = GetValidKmsServer -KmsHost $OfficeActiveInfo['KmsHost'] -KmsIp '' -OsppPath $OsppPath
    if (!$ValidKms) {
        CScript //Nologo "$OsppPath" /remhst | Out-Null
        return
    }

    Write-Host -Object ''
    Write-Host -Object '开始激活 Office 2021 批量授权版'

    $NewActiveInfo = GetOfficeActiveInfo -OsppPath $OsppPath
    Write-Host -Object ''
    if ($NewActiveInfo['IsActive'] -and $OfficeActiveInfo['ActiveEndTime'] -ne $NewActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($NewActiveInfo['Name'] + ' 激活成功, 激活有效期至 ' + $NewActiveInfo['ActiveEndTime'])
    }
    elseif ($NewActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($NewActiveInfo['Name'] + ' 激活有效期至 ' + $NewActiveInfo['ActiveEndTime'])
    }
    else {
        Write-Warning -Message ($NewActiveInfo['Name'] + ' 激活失败')
    }
}

function ActiveWindows {

    Clear-Host
    Write-Host -Object ''
    Write-Host -Object '正在检测 Windows 版本，请勿关闭此窗口'

    $SystemInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    if (!$SystemInfo.Caption.Contains('10') -and !$SystemInfo.Caption.Contains('11')) {
        Write-Host -Object ''
        Write-Warning -Message "不支持激活 $($SystemInfo.Caption)"
        return
    }

    if (!(Test-Path -Path "$env:windir\System32\slmgr.vbs" -PathType Leaf)) {
        Write-Host -Object ''
        Write-Warning -Message "$env:windir\System32\slmgr.vbs 文件不存在，无法激活 $($SystemInfo.Caption)"
        return
    }

    $WindowsActiveInfo = GetWindowsActiveInfo
    Write-Host -Object ''
    if ($WindowsActiveInfo['IsActive']) {
        Write-Host -Object ($SystemInfo.Caption + ' 已激活, 激活有效期至 ' + $WindowsActiveInfo['ActiveEndTime'])
    }
    elseif ($WindowsActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($SystemInfo.Caption + ' 激活有效期至 ' + $WindowsActiveInfo['ActiveEndTime'])
    }
    else {
        Write-Host -Object ($SystemInfo.Caption + ' 未激活')
    }

    $Gvlk = GetWindowsGvlk -WindowsActiveInfo $WindowsActiveInfo
    if ($null -eq $Gvlk) {
        return
    }

    if ('' -ne $Gvlk) {
        Write-Host -Object ''
        Write-Host -Object "安装产品密钥: $Gvlk"
        Write-Host -Object ''
        CScript //Nologo "$env:windir\System32\slmgr.vbs" /ipk $Gvlk
    }

    $ValidKms = GetValidKmsServer -KmsHost $WindowsActiveInfo['KmsHost'] -KmsIp $WindowsActiveInfo['KmsIp'] `
        -OsppPath $null
    if (!$ValidKms) {
        CScript //Nologo "$env:windir\System32\slmgr.vbs" /ckms | Out-Null
        return
    }

    Write-Host -Object ''
    Write-Host -Object "开始激活 $($SystemInfo.Caption)"

    $NewActiveInfo = GetWindowsActiveInfo
    Write-Host -Object ''
    if ($NewActiveInfo['IsActive'] -and $WindowsActiveInfo['ActiveEndTime'] -ne $NewActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($SystemInfo.Caption + ' 激活成功, 激活有效期至 ' + $NewActiveInfo['ActiveEndTime'])
    }
    elseif ($NewActiveInfo['ActiveEndTime']) {
        Write-Host -Object ($SystemInfo.Caption + ' 激活有效期至 ' + $NewActiveInfo['ActiveEndTime'])
    }
    else {
        Write-Warning -Message ($SystemInfo.Caption + ' 激活失败')
    }
}

function InstallOffice {

    Clear-Host

    $OfficeProducts = [ordered]@{
        'Word'       = 'Word';
        'Excel'      = 'Excel';
        'PowerPoint' = 'PowerPoint';
        'Outlook'    = 'Outlook';
        'OneNote'    = 'OneNote';
        'OneDrive'   = 'OneDrive';
        'Visio'      = 'Visio';
        'Project'    = 'Project';
        'Teams'      = 'Teams';
        'Lync'       = 'Skype';
        'Access'     = 'Access';
        'Publisher'  = 'Publisher'
    }

    $NeedOfficeProducts = [ordered]@{}
    while ($true) {
        Clear-Host
        $Reset = $false
        foreach ($Product in $OfficeProducts.GetEnumerator()) {
            $Reset = $false
            $Select = GetOfficeProductSelect -ProductName $Product.Value
            if (1 -eq $Select) {
                $NeedOfficeProducts.Add($Product.Key, $true)
            }
            elseif (2 -eq $Select) {
                $NeedOfficeProducts = [ordered]@{}
                $Reset = $true
                break
            }
        }
        if (!$Reset) {
            $Select = ConfirmOfficeProducts -NeedOfficeProducts $NeedOfficeProducts -OfficeProducts $OfficeProducts
            if (0 -eq $Select) {
                return
            }
            elseif (1 -eq $Select) {
                break
            }
            $NeedOfficeProducts = [ordered]@{}
        }
    }

    CreateOfficeDeploymentFile -NeedOfficeProducts $NeedOfficeProducts

    Set-Location -Path "$PSScriptRoot"

    Write-Host -Object ''
    Write-Host -Object '正在下载 Office 2021 批量授权版安装文件，耗时较长，请勿关闭此窗口'
    .\setup.exe /download configuration.xml
    Write-Host -Object ''
    Write-Host -Object 'Office 2021 批量授权版安装文件下载成功'

    Write-Host -Object ''
    Write-Host -Object '正在安装 Office 2021 批量授权版，耗时较长，请勿关闭此窗口'
    .\setup.exe /configure configuration.xml
    Write-Host -Object ''
    Write-Host -Object 'Office 2021 批量授权版安装完成'

    ActiveOffice
}

function CleanFile {

    Clear-Host
    Set-Location -Path "$PSScriptRoot"

    if (Test-Path -Path Office -PathType Container) {
        Remove-Item -Path Office -Recurse -Force
    }
    if (Test-Path -Path configuration.xml -PathType Leaf) {
        Remove-Item -Path configuration.xml -Force
    }

    Write-Host -Object ''
    Write-Host -Object 'Office 安装文件缓存清理完成'
}

function CreateShortcut {
    param ($Type)

    Clear-Host

    $TargetPath = [System.Environment]::GetFolderPath([Environment+SpecialFolder]::Programs) + '\KmsTool.lnk'
    if ($Type -eq 1) {
        $TargetPath = [System.Environment]::GetFolderPath([Environment+SpecialFolder]::Desktop) + '\KmsTool.lnk'
    }

    if (Test-Path -Path "$TargetPath" -PathType Leaf) {
        Remove-Item -Path "$TargetPath" -Force
    }

    $WScriptShell = New-Object -ComObject 'WScript.Shell'
    $Shortcut = $WScriptShell.CreateShortcut("$TargetPath")
    $Shortcut.TargetPath = "$PSScriptRoot\KmsTool.cmd"
    $Shortcut.WindowStyle = 1
    $Shortcut.WorkingDirectory = "$PSScriptRoot"
    $Shortcut.Save()

    Write-Host -Object ''
    Write-Host -Object '快捷方式创建完成'
}

function MainMenu {

    Clear-Host

    Write-Host -Object ''
    Write-Host -Object "=====> KmsTool v$VersionInfo https://github.com/dsx42/KmsTool <====="
    Write-Host -Object ''
    Write-Host -Object '================'
    Write-Host -Object '选择要进行的操作'
    Write-Host -Object '================'
    Write-Host -Object ''
    Write-Host -Object '1: 安装 Office 2021 批量授权版'
    Write-Host -Object ''
    Write-Host -Object '2: 激活 Office 2021 批量授权版'
    Write-Host -Object ''
    Write-Host -Object '3: 激活 Windows 10/11 批量授权版'
    Write-Host -Object ''
    Write-Host -Object '4: 清理 Office 安装文件缓存'
    Write-Host -Object ''
    Write-Host -Object '5: 为本工具创建桌面快捷方式'
    Write-Host -Object ''
    Write-Host -Object '6: 为本工具创建开始菜单快捷方式'
    Write-Host -Object ''
    Write-Host -Object 'q: 退出'

    $InputOption = 'q'
    while ($true) {
        Write-Host -Object ''
        $InputOption = Read-Host -Prompt '请输入选择的序号，按回车键确认'
        if ($null -eq $InputOption -or '' -eq $InputOption) {
            Write-Host -Object ''
            Write-Warning -Message '选择无效，请重新输入'
            continue
        }
        if ('q' -ieq $InputOption -or '1' -ieq $InputOption -or '2' -ieq $InputOption -or '3' -ieq $InputOption `
                -or '4' -ieq $InputOption -or '5' -ieq $InputOption -or '6' -ieq $InputOption) {
            break
        }
        Write-Host -Object ''
        Write-Warning -Message '选择无效，请重新输入'
    }

    if ('q' -eq $InputOption) {
        [System.Environment]::Exit(0)
    }
    if ('1' -eq $InputOption) {
        InstallOffice
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
    if ('2' -eq $InputOption) {
        ActiveOffice
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
    if ('3' -eq $InputOption) {
        ActiveWindows
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
    if ('4' -eq $InputOption) {
        CleanFile
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
    if ('5' -eq $InputOption) {
        CreateShortcut -Type 1
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
    if ('6' -eq $InputOption) {
        CreateShortcut -Type 2
        Write-Host -Object ''
        Read-Host -Prompt '按确认键返回主菜单'
        MainMenu
    }
}

$VersionInfo = GetVertion

if ($Version) {
    return $VersionInfo
}

RequireAdmin

$PSDefaultParameterValues['*:Encoding'] = 'utf8'
$ProgressPreference = 'SilentlyContinue'
$Host.UI.RawUI.WindowTitle = "KmsTool v$VersionInfo"
Set-Location -Path $PSScriptRoot

$kmsServers = @(
    'kms.03k.org',
    'kms.cangshui.net',
    'skms.netnr.eu.org'
)

MainMenu
