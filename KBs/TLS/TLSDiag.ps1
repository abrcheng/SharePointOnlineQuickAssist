param(
    [switch]
    $Resolve
)

$OSVersion = [environment]::OSVersion.Version

$script:TLSDiagResult = New-Object PSObject
$TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLS12Supported -Value $false

$script:RequireReboot = $false

$O365SupportedCipherSuites = (
    "TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384", 
    "TLS_ECDHE_RSA_WITH_AES_128_GCM_SHA256", 
    "TLS_DHE_RSA_WITH_AES_256_GCM_SHA384", 
    "TLS_DHE_RSA_WITH_AES_128_GCM_SHA256"
);

function IsNull($obj, $def)
{
    if ($obj -ne $null)
    {
        return $obj
    }
    else
    {
        return $def
    }
}

function DiagLegacyOSTLSSetting()
{
    $result = $true
    $key = Get-Item "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
    if ($key -eq $null)
    {
       $TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLS12ClientRegKeyExists -Value $false
        if ($Resolve)
        {
            $key = New-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Force
            if ($key)
            {
                $TLSDiagResult.TLS12ClientRegKeyExists = $true
                $script:RequireReboot = $true
            }
            else
            {
                $result = $false
            }
        }
    }

    if ($key.GetValue("DisabledByDefault") -ne 0)
    {
        $TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLS12ClientDisabledByDefault -Value  (IsNull -obj $key.GetValue("DisabledByDefault") -def "Not Configured")
        if ($Resolve)
        {
            if (New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "DisabledByDefault" -Value 0 -PropertyType DWORD -Force)
            {
                TLSDiagResult.TLS12ClientDisabledByDefault = 0;
                $script:RequireReboot = $true
            }
            else
            {
                $result = $false
            }
        }
        else
        {
            $result = $false
        }

    }

    if ($key.GetValue("Enabled") -eq 0)
    {
        $TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLS12ClientEnabled -Value (IsNull -obj $key.GetValue("Enabled") -def "Not Configured")
        if ($Resolve)
        {
            if (Remove-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client" -Name "Enabled")
            {
                $TLSDiagResult.TLS12ClientEnabled = $null;
                $script:RequireReboot = $true
            }
            else
            {
                $result = $false
            }

        }
        else
        {
            $result = $false
        }

    }
    return $result
}

function DiagDotNetTLSSetting()
{
    $result = $true

    $key = Get-Item "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727"
    if ($key -ne $null)
    {
        if ($key.GetValue("SystemDefaultTlsVersions") -ne 1)
        {
            $TLSDiagResult | Add-Member -MemberType NoteProperty -Name DotNet35SystemDefaultTlsVersions -Value (IsNull -obj $key.GetValue("SystemDefaultTlsVersions") -def "Not Configured")
            if ($Resolve)
            {
                if (New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v2.0.50727" -Name "SystemDefaultTlsVersions" -Value 1 -PropertyType DWORD -Force)
                {
                    $TLSDiagResult.DotNet35SystemDefaultTlsVersions = 1;
                    $script:RequireReboot = $true
                }
                else
                {
                    $result = $false
                }

            }
            else
            {
                $result = $false
            }

        }
    }

    $key = Get-Item "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727"
    if ($key -ne $null)
    {
        if ($key.GetValue("SystemDefaultTlsVersions") -ne 1)
        {
            $TLSDiagResult | Add-Member -MemberType NoteProperty -Name DotNet35Wow6432SystemDefaultTlsVersions -Value (IsNull -obj $key.GetValue("SystemDefaultTlsVersions") -def "Not Configured")
            if ($Resolve)
            {
                if (New-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v2.0.50727" -Name "SystemDefaultTlsVersions" -Value 1 -PropertyType DWORD -Force)
                {
                    $TLSDiagResult.DotNet35Wow6432SystemDefaultTlsVersions = 1;
                    $script:RequireReboot = $true
                }
                else
                {
                    $result = $false
                }
            }
            else
            {
                $result = $false
            }


        }
    }

    $key = Get-Item "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319"
    if ($key -ne $null)
    {
        if ($key.GetValue("SchUseStrongCrypto") -ne 1)
        {
            $TLSDiagResult | Add-Member -MemberType NoteProperty -Name DotNet40SchUseStrongCrypto -Value (IsNull -obj $key.GetValue("SchUseStrongCrypto") -def "Not Configured")
            if ($Resolve)
            {
                if (New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -PropertyType DWORD -Force)
                {
                    $TLSDiagResult.DotNet40SchUseStrongCrypto = 1;
                    $script:RequireReboot = $true
                }
                else
                {
                    $result = $false
                }
            }
            else
            {
                $result = $false
            }

        }
    }

    $key = Get-Item "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319"
    if ($key -ne $null)
    {
        if ($key.GetValue("SchUseStrongCrypto") -ne 1)
        {
            $TLSDiagResult | Add-Member -MemberType NoteProperty -Name DotNet40Wow6432SchUseStrongCrypto -Value (IsNull -obj $key.GetValue("SchUseStrongCrypto") -def "Not Configured")
            if ($Resolve)
            {
                if (New-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\v4.0.30319" -Name "SchUseStrongCrypto" -Value 1 -PropertyType DWORD -Force)
                {
                    $TLSDiagResult.DotNet40Wow6432SchUseStrongCrypto = 1;
                    $script:RequireReboot = $true
                }
                else
                {
                    $result = $false
                }
            }
            else
            {
                $result = $false
            }
        }
    }
    return $result
}

function DiagTlsCipherSuite($IsLegacy = $false)
{
    $TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLSCipherSuiteSupported -Value $false

    $CipherList = $null;
    if ($IsLegacy)
    {
        $key = get-item "HKLM:\SYSTEM\CurrentControlSet\Control\Cryptography\Configuration\Local\Default\00010002"
        $CipherList = $key.GetValue("Functions").Split(",");
    }
    else
    {
        $CipherList = (Get-TlsCipherSuite).Name;
    }
    $CipherList |% {
        foreach ($cs in $O365SupportedCipherSuites)
        {
            if ($cs -eq $_)
            {
                $TLSDiagResult.TLSCipherSuiteSupported = $true
            }
        }
    }

    if (!$TLSDiagResult.TLSCipherSuiteSupported)
    {
        if ($IsLegacy)
        {
            Write-Host "Cipher Suite Order configuration might be required!" -Foreground Green
            Write-Host " 1. Open the Group Policy Editor (Either gpedit.msc or gpmc.msc)" -Foreground Green
            Write-host " 2. Computer Configuration - Administrative Templates - Network - SSL Configuration Settings" -Foreground Green
            Write-Host " 3. Open SSL Cipher Suite Order" -Foreground Green
            Write-Host " 4. Enable either of the following supported TLS 1.2 CipherSuites in the top order." -Foreground Green
            $O365SupportedCipherSuites |% { Write-Host ("  " + $_)  -Foreground Green}
            Write-Host " Refer https://docs.microsoft.com/en-us/sharepoint/troubleshoot/administration/authentication-errors-tls12-support " -Foreground Green
        }

        if ($Resolve)
        {
            if (!$IsLegacy)
            {
                Enable-TlsCipherSuite -Name "TLS_ECDHE_RSA_WITH_AES_256_GCM_SHA384" -Position 0
            }
        }
    }


}



Write-Host "----------------------------------------------" -ForegroundColor Green

switch($OSVersion.Major)
{
    6 {
        switch($OSVersion.Minor)
        {
            0 {
                #2008/Vista
                Write-Host "Windows 2008 / Vista" -ForegroundColor Green
                Write-Host " Prerequisite : https://support.microsoft.com/en-us/help/3154517/" -ForegroundColor Green
                $TLSDiagResult.TLS12Supported = DiagLegacyOSTLSSetting
                break;
            } 
            1 {
                # 2008R2/7
                Write-Host "Windows 2008R2 / 7" -ForegroundColor Green
                Write-Host " Prerequisite : https://support.microsoft.com/en-us/help/3154518/" -ForegroundColor Green
                $TLSDiagResult.TLS12Supported = DiagLegacyOSTLSSetting
                break;
            } 
            2 {
                # 2012/8
                Write-Host "Windows 2012 / 8" -ForegroundColor Green
                Write-Host " Prerequisite : https://support.microsoft.com/en-us/help/3154519/" -ForegroundColor Green
                $TLSDiagResult.TLS12Supported = $true;
                break;
            } 
            3 {
                #2012R2/8.1
                Write-Host "Windows 2012R2 / 8.1" -ForegroundColor Green
                Write-Host " Prerequisite : https://support.microsoft.com/en-us/help/3154520/" -ForegroundColor Green
                $TLSDiagResult.TLS12Supported = $true;
                break;
            } 
            else
            {
                throw "This script is not supported in this machine."
            }
        }
        DiagTlsCipherSuite -IsLegacy $true
        break;
    }
    10 {
        Write-Host "Windows 10 : Only if using LTSC, please check the following." -ForegroundColor Green
        Write-Host " v1507 (LTSC)" -ForegroundColor Green
        Write-Host "  Prerequisite : https://support.microsoft.com/en-us/help/4001772/" -ForegroundColor Green
        Write-Host " v1607 (LTSC) " -ForegroundColor Green
        Write-Host "  Option 1 : https://support.microsoft.com/en-us/help/4004253/" -ForegroundColor Green
        Write-Host "  Option 2 : https://support.microsoft.com/en-us/help/4004227/" -ForegroundColor Green

        $TLSDiagResult.TLS12Supported = $true;

        DiagTlsCipherSuite

        break;
    }
    else
    {
        throw "This script is not supported in this machine."
        return;
    }
}
Write-Host "----------------------------------------------" -ForegroundColor Green


$TLSDiagResult | Add-Member -MemberType NoteProperty -Name TLSDotNetConfigured -Value $false
$DotNetresult = (DiagDotNetTLSSetting);
$TLSDiagResult.TLSDotnetConfigured = $DotNetResult

$PSSupported = $false
foreach ($oktls in ("Tls12", "SystemDefault", "0"))
{
    if ([System.Net.ServicePointManager]::SecurityProtocol -like "*" + $oktls + "*")
    {
        $PSSupported = $true
    }
}

if (!$PSSupported)
{
    $TLSDiagResult | Add-Member -MemberType NoteProperty -Name PowerShellSecurityProtocol -Value ([System.Net.ServicePointManager]::SecurityProtocol.ToString())
    if ($Resolve)
    {
        Write-Host "Please consider specifying the following code in your PowerShell Solution." -Foreground Green
        Write-Host " [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SolutionProtocolType]::Tls12 " -Foreground Green
    }
}

$TLSDiagResult | fl

if ($script:RequireReboot)
{
    Write-Host "Machine reboot might be required..." -Foreground Red
    Write-Host

}

