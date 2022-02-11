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



$global:DEBUG = $true

$global:User = "ftr_admin"
$global:Pass = "Fortherecord123!"

# This value needs to be copied while the script is running. It doesn't work if you run the script, and then try to reference the value from console
$global:ScriptPath = $PSScriptRoot + '\'
$global:DataFileName = "AZMC_CourtroomData.csv"
$global:SelectedFile = ""

Write-Host -ForegroundColor Cyan @"

FTR's Magic Crestron Configuration Script

"@

# MyShell
Write-Host -ForegroundColor Yellow "Hi!`nI'm Jonks, your friendly neighborhood PowerShell script.`n`n`n"


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
Add-Type -AssemblyName System.Windows.Forms

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


# Init Menu

class Menu
{
    $L1 = @(  "`t1) Log Files & Reports`n" + 
              "`t2) Spreadsheet`n" +
              "`t3) Device Config`n" + 
              "`t4) Script Settings"   )

    $L2_1 = @( 
                )
    $L2_2 = @( 
                )
    $L2_3 = @( 
                )
    $L2_4 = @( 
                )
    $L2_5 = @( 
                )


}

function menu
{
    $menu
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

}




# Import Device Data
#$global:rooms = @{}
#$global:roomsByName = @{}


function importFile([string]$fileName)
{
   
    if([System.IO.File]::Exists("$fileName"))
    {
        $global:sheet = Import-csv $fileName
        $sheetLen = $sheet | Measure-Object | Select-Object -ExpandProperty Count
        Write-Host -ForegroundColor Yellow ("Ok, I imported the file $fileName.`nThere are $sheetLen rooms in the list.`n`n")
    }
    else
    {
        err 3
        return ""
    }

    $global:rooms = @{}
    $global:roomsByName = @{}

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

        $c.RoomName = $row | Select-object -ExpandProperty Room_Name
        $c.FacilityName = $row | select-object -ExpandProperty Facility_Name

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

        $rooms[$c.Index] = $c
        $roomsByName[$c.RoomName] = $rooms[$c.Index]
    }
}

function getFileName
{
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = $ScriptPath }
    $null = $FileBrowser.ShowDialog()
    $fileName = $FileBrowser | Select-Object -ExpandProperty FileName
    Write-Host $fileName
    return ($fileName)    
}

function selectAndImport
{  
    importFile (getFileName)
}

