<#

.INSTRUCTIONS
    Read the README.txt file.
    Seriously.

.DESCRIPTION
    Pull in CSV of all Courtroom configuration details. Updates Display Project/Program, and Updates IP Tables.

.INPUTS
    Complete CSV with all Courtroom configuraiton details placed in local directory with script.

.NOTES
    Version: 1.01.00
    Creation Date: 2/1/2022
    Purpose/Change: Adding Addtional Logging
#>



$DEBUG = $true

$User = "ftr_admin"
$Pass = "Fortherecord123!"

# This value needs to be copied while the script is running. It doesn't work if you run the script, and then try to reference the value from console
$ScriptPath = $PSScriptRoot + '\'
$DataFileName = "AZMC_CourtroomData.csv"


Write-Host -ForegroundColor Cyan @"

FTR's Magic Crestron Configuration Script

"@

# MyShell
Write-Host -ForegroundColor Yellow "Hi!`nI'm Jonks, your friendly neighborhood PowerShell script.`nI'll be your guide today.`n`n`n"


function fatal
{
    Read-Host -Prompt "`n`nPress any key to exit"
    exit
}
# Error handling
function err([string]$err)
{
    Write-Host $err
}
function err([int]$i)
{
    switch($i)
    {
        1 
        { 
            Write-Host -ForegroundColor Yellow 
            (
                "Hey!`n" + 
                "This script requires software called 'PSCrestron' (the PowerShellCrestron module).`n" +
                "I believe you have a good internet connection right now, but something else went wrong when I tried to download and install this software.`n" +
                
                "If you get this message again, try to manually download and install PSCrestron. Drop this into a web browser:`n" +
                "https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe`n" +
                "... and run the file, following any prompts to install the software.`n`n" +

                "If you're still having problems, you can email bcampagnola@fortherecord.com, with 'HALP!' in the subject line." 
            ) 
            fatal
        } 
        2 
        {
            Write-Host -ForegroundColor Yellow 
            (
                "Hey!`n" +
                "This script requires software called 'PSCrestron' (the PowerShellCrestron module).`n" +
                "I attempted to download the software, but it appears that this computer doesn't have an internet connection.`n`n" +

                "Try gettin' some internets, and then run the script again.`n`n" +

                "If you believe this is not the problem, there might be something wrong with the permissions on your machine, or you're having delusions of connectivity.`n`n`n" +


                "If you're online but still having problems, you can email bcampagnola@fortherecord.com, with 'HALP!' in the subject line." 
            )        
            fatal        
        }
        3
        {
            Write-Host -ForegroundColor Yellow 
            (
                "Hm.." + 
                "I couldn't find that file.`n" + 
                "Does this path look right?`n`n" +        
                
                "$ScriptPath" + "$DatafileName`n`n" + 

                "If not, try the options to edit the script path and .csv file name values.`n`n"
            )    
        }
    }
}


# Import Libs

function tryImport
{
    try
    {
        Import-Module PSCrestron -ErrorAction Stop
        return 0
    }
    catch 
    { 
        return 1 
    }
}


$src = ""
$dst = ""
$err = tryImport

if($err -ne 0)
{
    $src = 'https://sdkcon78221.crestron.com/downloads/EDK/EDK_Setup_1.0.5.3.exe'
    $dst = $env:temp + '\EDK_Setup_1.0.5.3.exe'

    if(Test-Connection "8.8.8.8" -Count 1 -Quiet)
    {
        Invoke-RestMethod $src -OutFile $dst -TimeoutSec 8
        Start-Process -Wait -FilePath "$dst" -ArgumentList "/S" -PassThru
    }
    else
    {
        erre 2
    }

    $err = tryImport
    if($err -ne 0)
    {
        err 1
    }
}






# Stopwatch Feature
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# Object for Results Data
$DeviceResultsData = @() | Out-GridView -Title "Device Status Results"




function ProcessorConfig([string]$ProcIP, [string[]]$AVBEntry, [string]$LPZPath, [string]$WyreIP)
{
    try
    {
        # Clear Connect Error
        $ConnectError = " "

        # Clear Status
        $Status = " "
        
        # New Data Object
        $DeviceResultItem = New-Object PSObject
        Write-Host -f Green "Proc Config Setup for: $ProcIP"
        
        # Connect to Device via SSH
        $SessionID = Open-CrestronSession -Device $ProcIP -Secure -Username $User -Password $Pass -ErrorAction SilentlyContinue

        # Hostname
        $hostnameResponce = Invoke-CrestronSession $SessionID "hostname"
        $deviceHostname = [regex]::Match($hostnameResponce, "(?<=Host\sName:\s)[\w-]+").value
        Write-Host -f Green "Working on => $deviceHostname`n`n`n`n"
        
        # Send Crestron Program to Processor
        Send-CrestronProgram -ShowProgress -Device $ProcIP -LocalFile $LPZPath -Secure -Username $User -Password $Pass

        # Configure Processor IP Tabele
        
        Invoke-CrestronSession $SessionID -Command "AddP 11 $WyreIP"
        $IPID = 0x30
        foreach ($AVBIP in $AVBEntry)
        {
            $AddPeer = "AddP {0:X} {1}" -f $IPID, $AVBIP
            Invoke-CrestronSession $SessionID -Command "$AddPeer"
            $IPID += 1
        }
    
        # Reset Program
        Invoke-CrestronSession $SessionID 'progreset -p:01'
        
        # Close Session
        Close-CrestronSession $SessionID
        
        #Status to Log and Console
        $Status = "Processor Config Success"
        Write-Host -ForegroundColor Green "`n - Processor Config Success`n`n`n`n"
      
    }

    catch
    {
        $deviceHostname = " "
        $ConnectError = "Connection Attempts Unsuccessful"
        Write-Host -f Red "`n $d - Unable to Configure Processor`n`n`n`n"
    }

    #Current Date/Time
    $time = (get-date)
    #Build Table
    # Table Coulumn 1 - Time
    $DeviceResultItem | Add-Member -Name "Time" -MemberType NoteProperty -Value $time
    # Table Coulumn 2 - IP Address
    $DeviceResultItem | Add-Member -Name "IP Address" -MemberType NoteProperty -Value $ProcIP
    # Table Coulumn 3 - Hostname
    $DeviceResultItem | Add-Member -Name "Hostname" -MemberType NoteProperty -Value $deviceHostname
    # Table Coulumn 4 - Error
    $DeviceResultItem | Add-Member -Name "Error" -MemberType NoteProperty -Value $ConnectError
    # Table Coulumn 5 - Status
    $DeviceResultItem | Add-Member -Name "Status" -MemberType NoteProperty -Value $Status    
    # Add line to the report
    $DeviceResultsData += $DeviceResultItem
 
    #Append results to Processor Change Results Document + Log 
    $DeviceResultsData | Export-Csv -Path "$Path\Proc Config Results.csv" -NoTypeInformation -append
    $DeviceResultsData | Export-Csv -Path "$Path\Proc Config Results Log.csv" -NoTypeInformation -append

    # Total time of script
    # $stopwatch

}

function PanelConfig([string]$Panel, [string]$VTZPath, [string]$ProcIP, [int]$IPID)
{
    try
    {
        # Clear Connect Error
        $ConnectError = " "

        # Clear Status
        $Status = " "
        
        # New Data Object
        $DeviceResultItem = New-Object PSObject
        Write-Host -f Green "Panel Config Setup for: $Panel"
        
        # Connect to Device via SSH
        $SessionID = Open-CrestronSession -Device $Panel -Secure -Username $User -Password $Pass -ErrorAction SilentlyContinue

        # Hostname
        $hostnameResponce = Invoke-CrestronSession $SessionID "hostname"
        $deviceHostname = [regex]::Match($hostnameResponce, "(?<=Host\sName:\s)[\w-]+").value
        Write-Host -f Green "Working on => $deviceHostname`n`n`n`n"
        
        # Send Display Project
        Send-CrestronProject -ShowProgress -Device $Panel -LocalFile $VTZPath -Secure -Username $User -Password $Pass
        
        # IP Table Setup
        Write-Host '>>> Setting up Panel IP Table '+ $Panel
        $SessionID = Open-CrestronSession -Device $Panel -Secure -Username $User -Password $Pass -ErrorAction SilentlyContinue
        $AddMaster = "addm {0:X} {1}" -f $IPID, $ProcIP
        Invoke-CrestronSession $SessionID -Command "$AddMaster"
        
        # Close Session
        Close-CrestronSession $SessionID
        
        # Print Success to Console
        $Status = "Panel Config Success"
        Write-Host -ForegroundColor Green "`n - Panel Config Success`n`n`n`n"
    }

    catch
    {
        # Write Error Log 
        $deviceHostname =" "
        $ConnectError = "Connection Attempts Unsuccessful"
        Write-Host -f Red "`n $d - Unable to Configure Panel`n`n`n`n"
    }

    # Current Date/Time
    $time = (get-date)
    # Build Table
    # Table Coulumn 1 - Time
    $DeviceResultItem | Add-Member -Name "Time" -MemberType NoteProperty -Value $time
    # Table Coulumn 2 - IP Address
    $DeviceResultItem | Add-Member -Name "IP Address" -MemberType NoteProperty -Value $Panel
    # Table Coulumn 3 - Hostname
    $DeviceResultItem | Add-Member -Name "Hostname" -MemberType NoteProperty -Value $deviceHostname
    # Table Coulumn 4 - Error
    $DeviceResultItem | Add-Member -Name "Error" -MemberType NoteProperty -Value $ConnectError
    # Table Coulumn 5 - Status
    $DeviceResultItem | Add-Member -Name "Status" -MemberType NoteProperty -Value $Status
    # Add line to the report
    $DeviceResultsData += $DeviceResultItem
 
    # Append results to Panel Config Change Results Document + Log 
    $DeviceResultsData | Export-Csv -Path "$Path\Panel Config Results.csv" -NoTypeInformation -append
    $DeviceResultsData | Export-Csv -Path "$Path\Panel Config Results Log.csv" -NoTypeInformation -append

    # Total time of script
    # $stopwatch

}

class Courtroom
{
    [int]$Index
    [string]$CommentLine
    [string]$RoomName
    [string]$FacilityName
    [string]$Subnet
    [string]$Processor_IP
    [string]$FileName_LPZ
    [string[]]$Panel_IP
    [string[]]$FileName_VTZ
    [string]$ReporterWebSvc_IP
    [string]$Wyrestorm_IP
    [string[]]$FixedCam_IP
    [string[]]$DSP_IP
    [string[]]$RecorderSvr_IP
    [string]$DVD_IP
    [string]$Audicue_IP
    [string[]]$PTZCam_IP

    [int]$IPID

}




# Import Device Data

function importFile
{
    $sheet = ""
    $fileName = $ScriptPath + $DataFileName
    
    if([System.IO.File]::Exists("$fileName"))
    {
        $sheet = Import-csv $fileName
        return $sheet
    }
    else
    {
        err 3
    }
}

function parseFile($sheet)
{
    $rooms = @{}
    $i = 1
    

    foreach($row in $sheet)
    {
        # $i is initialized to 1. As this value will be the dict key, and we want that to match the .csv line number (for convenience),
        #     we will start the loop by adding 1.
        #     Ergo, the first dict entry will be     (2, $c) 
     
        $i += 1
        Write-Host $1

        $c = new-object -TypeName Courtroom
 
        $c.Index = $i           
        $c.CommentLine = $row | Select-object -ExpandProperty Ignore_Line

        $c.RoomName, $c.FacilityName = $row | Select-object -ExpandProperty Room_Name, Facility_Name
        #$c.FacilityName = $row | select-object -ExpandProperty Facility_Name

        $c.Subnet = $row | select-object -ExpandProperty Subnet_Address

        $c.Processor_IP = $row | Select-object -ExpandProperty Processor_IP
        $c.FileName_LPZ = $row | Select-object -ExpandProperty FileName_LPZ
        $c.Panel_IP = $row | Select-object -ExpandProperty Panel_IP
        $c.Panel_IP = $c.Panel_IP[0].split('~')
        $c.FileName_VTZ = $row | Select-object -ExpandProperty FileName_VTZ
        $c.FileName_VTZ = $c.FileName_VTZ[0].split('~')

        $c.ReporterWebSvc_IP = $row | Select-object -ExpandProperty IP_ReporterWebSvc

        $c.Wyrestorm_IP = $row | Select-Object -ExpandProperty IP_WyrestormCtrl

        $c.FixedCam_IP = $row | Select-object -ExpandProperty IP_FixedCams
        $c.FixedCam_IP = $c.FixedCam_IP[0].split('~')
            
        $c.DSP_IP = $row | Select-object -ExpandProperty IP_DSPs
        $c.DSP_IP = $c.DSP_IP[0].split('~')

        $c.RecorderSvr_IP = $row | Select-object -ExpandProperty IP_Recorders
        $c.RecorderSvr_IP = $c.RecorderSvr_IP[0].split('~')

        $c.DVD_IP = $row | Select-Object -ExpandProperty IP_DVDPlayer

        $c.Audicue_IP = $row | Select-Object -ExpandProperty IP_AudicueGW

        $c.PTZCam_IP = $row | Select-Object -ExpandProperty IP_PTZCams 
        $c.PTZCam_IP = $c.PTZCam_IP[0].split('~')

        $rooms.Add($c.Index, $c)
    }
}

# SIG # Begin signature block
# MIIR2wYJKoZIhvcNAQcCoIIRzDCCEcgCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+bGkPtvvLrtJwSSopJtD11BS
# dECggg1BMIIDBjCCAe6gAwIBAgIQbF2ntJf28JdI27Ov5cN1bzANBgkqhkiG9w0B
# AQsFADAbMRkwFwYDVQQDDBBBVEEgQXV0aGVudGljb2RlMB4XDTIyMDIxMDEwMzI1
# OFoXDTIzMDIxMDEwNTI1OFowGzEZMBcGA1UEAwwQQVRBIEF1dGhlbnRpY29kZTCC
# ASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAOJ+Y5UdHSUAp7OWQxQNQRo5
# hAzfTgx7L47ZeEJf4DnFKTYdWooVOxPVHH+O+Q8hsMm07Tdm/iW90z7I1lH9lcXb
# KodsjJdGDywcQGCPCgrTanvUewvuKYg2rhr5ZF6k1HZ/Au/JSo7uHm6ACEP2siEy
# RDWmC2WaQePnaGiuLm7FtE9N6WnWeLSO/LnAcSxfsx0LJsRudMIPY0ar/91B7QwV
# +aE6qQIaK7JcpJx+M2XrhVz6VcwqvhncOWxWnceYYU3F0e3KoaJn+5tAd2dpylnZ
# BF4+koxZiwyhJXSfA7eFGbuKc71/lIM+d9PP0qhxHVub9VZAErMdpfsc1Q5aIW0C
# AwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMB0G
# A1UdDgQWBBS1eCnZPenhpTfSPIac3LPlyPMjPTANBgkqhkiG9w0BAQsFAAOCAQEA
# Q3ZU6weRZAFqd9O/P7V3pZXgkw1tTrFlQV5wOj+tMm9xAEYdqDxCbEpYClJ2SIEJ
# v3SLfSEWdXrq/7S07cmegpS536dOw61tS3pV4bKGTMSY72zrNSloQCXEqT9et7gV
# XBXm/pWGBb82074XEbyB+YSEJsCexQ2h0EXuNObcQ1whXUHfrb5/iIMgq+JmABVR
# YWJEVHbATQHaTWi9JZ4o4uAbrHVPT7qd0sMl1oCwqcwCkHMQXxAkSCe26xUxJ2gl
# UI/BdY9QBApSU2om9C4XVSmN7kabQbBCVEzL5KZTfunpE6MDixhTwaJh4S25pOfh
# fDCVnk/mU+gtDPktEDXw6DCCBP4wggPmoAMCAQICEA1CSuC+Ooj/YEAhzhQA8N0w
# DQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0
# IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTExMC8GA1UEAxMoRGlnaUNl
# cnQgU0hBMiBBc3N1cmVkIElEIFRpbWVzdGFtcGluZyBDQTAeFw0yMTAxMDEwMDAw
# MDBaFw0zMTAxMDYwMDAwMDBaMEgxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjEgMB4GA1UEAxMXRGlnaUNlcnQgVGltZXN0YW1wIDIwMjEwggEi
# MA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDC5mGEZ8WK9Q0IpEXKY2tR1zoR
# Qr0KdXVNlLQMULUmEP4dyG+RawyW5xpcSO9E5b+bYc0VkWJauP9nC5xj/TZqgfop
# +N0rcIXeAhjzeG28ffnHbQk9vmp2h+mKvfiEXR52yeTGdnY6U9HR01o2j8aj4S8b
# Ordh1nPsTm0zinxdRS1LsVDmQTo3VobckyON91Al6GTm3dOPL1e1hyDrDo4s1SPa
# 9E14RuMDgzEpSlwMMYpKjIjF9zBa+RSvFV9sQ0kJ/SYjU/aNY+gaq1uxHTDCm2mC
# tNv8VlS8H6GHq756WwogL0sJyZWnjbL61mOLTqVyHO6fegFz+BnW/g1JhL0BAgMB
# AAGjggG4MIIBtDAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUB
# Af8EDDAKBggrBgEFBQcDCDBBBgNVHSAEOjA4MDYGCWCGSAGG/WwHATApMCcGCCsG
# AQUFBwIBFhtodHRwOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwHwYDVR0jBBgwFoAU
# 9LbhIB3+Ka7S5GGlsqIlssgXNW4wHQYDVR0OBBYEFDZEho6kurBmvrwoLR1ENt3j
# anq8MHEGA1UdHwRqMGgwMqAwoC6GLGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9z
# aGEyLWFzc3VyZWQtdHMuY3JsMDKgMKAuhixodHRwOi8vY3JsNC5kaWdpY2VydC5j
# b20vc2hhMi1hc3N1cmVkLXRzLmNybDCBhQYIKwYBBQUHAQEEeTB3MCQGCCsGAQUF
# BzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wTwYIKwYBBQUHMAKGQ2h0dHA6
# Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFNIQTJBc3N1cmVkSURUaW1l
# c3RhbXBpbmdDQS5jcnQwDQYJKoZIhvcNAQELBQADggEBAEgc3LXpmiO85xrnIA6O
# Z0b9QnJRdAojR6OrktIlxHBZvhSg5SeBpU0UFRkHefDRBMOG2Tu9/kQCZk3taaQP
# 9rhwz2Lo9VFKeHk2eie38+dSn5On7UOee+e03UEiifuHokYDTvz0/rdkd2NfI1Jp
# g4L6GlPtkMyNoRdzDfTzZTlwS/Oc1np72gy8PTLQG8v1Yfx1CAB2vIEO+MDhXM/E
# EXLnG2RJ2CKadRVC9S0yOIHa9GCiurRS+1zgYSQlT7LfySmoc0NR2r1j1h9bm/cu
# G08THfdKDXF+l7f0P4TrweOjSaH6zqe/Vs+6WXZhiV9+p7SOZ3j5NpjhyyjaW4em
# ii8wggUxMIIEGaADAgECAhAKoSXW1jIbfkHkBdo2l8IVMA0GCSqGSIb3DQEBCwUA
# MGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsT
# EHd3dy5kaWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQg
# Um9vdCBDQTAeFw0xNjAxMDcxMjAwMDBaFw0zMTAxMDcxMjAwMDBaMHIxCzAJBgNV
# BAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdp
# Y2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBUaW1l
# c3RhbXBpbmcgQ0EwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQC90DLu
# S82Pf92puoKZxTlUKFe2I0rEDgdFM1EQfdD5fU1ofue2oPSNs4jkl79jIZCYvxO8
# V9PD4X4I1moUADj3Lh477sym9jJZ/l9lP+Cb6+NGRwYaVX4LJ37AovWg4N4iPw7/
# fpX786O6Ij4YrBHk8JkDbTuFfAnT7l3ImgtU46gJcWvgzyIQD3XPcXJOCq3fQDpc
# t1HhoXkUxk0kIzBdvOw8YGqsLwfM/fDqR9mIUF79Zm5WYScpiYRR5oLnRlD9lCos
# p+R1PrqYD4R/nzEU1q3V8mTLex4F0IQZchfxFwbvPc3WTe8GQv2iUypPhR3EHTyv
# z9qsEPXdrKzpVv+TAgMBAAGjggHOMIIByjAdBgNVHQ4EFgQU9LbhIB3+Ka7S5GGl
# sqIlssgXNW4wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wEgYDVR0T
# AQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUH
# AwgweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaG
# NGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcmwwUAYDVR0gBEkwRzA4BgpghkgBhv1sAAIEMCowKAYI
# KwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRpZ2ljZXJ0LmNvbS9DUFMwCwYJYIZIAYb9
# bAcBMA0GCSqGSIb3DQEBCwUAA4IBAQBxlRLpUYdWac3v3dp8qmN6s3jPBjdAhO9L
# hL/KzwMC/cWnww4gQiyvd/MrHwwhWiq3BTQdaq6Z+CeiZr8JqmDfdqQ6kw/4stHY
# fBli6F6CJR7Euhx7LCHi1lssFDVDBGiy23UC4HLHmNY8ZOUfSBAYX4k4YU1iRiSH
# Y4yRUiyvKYnleB/WCxSlgNcSR3CzddWThZN+tpJn+1Nhiaj1a5bA9FhpDXzIAbG5
# KHW3mWOFIoxhynmUfln8jA/jb7UBJrZspe6HUSHkWGCbugwtK22ixH67xCUrRwII
# fEmuE7bhfEJCKMYYVs9BNLZmXbZ0e/VWMyIvIjayS6JKldj1po5SMYIEBDCCBAAC
# AQEwLzAbMRkwFwYDVQQDDBBBVEEgQXV0aGVudGljb2RlAhBsXae0l/bwl0jbs6/l
# w3VvMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqG
# SIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3
# AgEVMCMGCSqGSIb3DQEJBDEWBBRHkOYSGaFjV2Ca14I5M7uExjIx0zANBgkqhkiG
# 9w0BAQEFAASCAQCUaRA5R5oEosQ/zWUJ52kKYYk1qPgC1fEz4E3Ir6LXj/Y7Saxf
# cbXIw7hac2PmYB7ROCfhCCw24TZslUNfgw5lPho4s4GN6Fuj2crNrp6YQR3bKZTQ
# SdhrmTk7+i5mMAHKURR6qBEpcCYXOD7s+KnVGUX1/8Dquk+eHmqtsZy2GIzwtYCN
# I1PURUWstSJpgrSnyYdaoEMIibx3ofPQCc2eO6oZYY/7U+HLdcNXnjzSw2hb2AYD
# EgDvzN5YofCtZBt+PEru5NasZkhlH90CcMC1D5XeEmLbhpugbk2JVN46B3hrDGX+
# F7bh5AaQ59S5dBGjyemFrr6YLwaI3X4ds3a+oYICMDCCAiwGCSqGSIb3DQEJBjGC
# Ah0wggIZAgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJ
# bmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAvBgNVBAMTKERpZ2lDZXJ0
# IFNIQTIgQXNzdXJlZCBJRCBUaW1lc3RhbXBpbmcgQ0ECEA1CSuC+Ooj/YEAhzhQA
# 8N0wDQYJYIZIAWUDBAIBBQCgaTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwG
# CSqGSIb3DQEJBTEPFw0yMjAyMTAxMDUzMTFaMC8GCSqGSIb3DQEJBDEiBCAhZt4M
# 8WxWsjhDyD2AF1FNe83oB3FLaQ/QTwrZvYVSwjANBgkqhkiG9w0BAQEFAASCAQAi
# p0e4k3cJl/ww0t4c3g8awzNZoCqOVlBDh71x/U5GG9y6A/dByybMYNTkz3VnwJ5P
# e8JTQHDITkssB5MZXL65UHXDPWdeFNv+/To0q30/QLr4CVkkLYfOgyWq1FxaqcxO
# wDlboMrJBI9iRpqQb8xoKpaSYN6SEMS4MmGIqOPAqH1uC/66qaCQ2mcTO9GhBpuH
# kngZdXgBODcb00SISTh8taxoVnJxKlIpHqHYl1mv8Bn96sf6RYtQjTenanMVcBIN
# Lxout/EhNxyVU/ZmtNjPCBRn+/nlXVlicOdqdiWweRIKbznt2n9k3eEwT+bV01G8
# /R/D5HiEZ/7JdOqjx/B8
# SIG # End signature block
