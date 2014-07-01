[CmdletBinding()]
Param
(
        [Parameter(Mandatory=$True,
            ValueFromPipeline=$True)]            
        [Alias('computer','cn')]
        [String[]]$ComputerName,

        [Alias('xp')]
        [String]$XPAntivirus = "\\192.168.20.98\C$\mseinstall.exe",

        [Alias('w7')]
        [String]$Win7Antivirus = "\\192.168.20.98\C$\SCEPInstall.exe",

        [String]$LocalPath = "C$\Temp\Antivirus"
    )

Begin
{
    Write-Debug "Creating output array"
    $Output = @()
    Write-Debug "Beginning loop through computer list"
}

Process
{
    ForEach ($comp in $ComputerName)
    {
        Write-Debug "Processing $comp"
        $temp = New-Object PSObject
        $temp | Add-Member -MemberType NoteProperty -Name ID -Value $comp
        $temp | Add-Member -MemberType NoteProperty -Name Name -Value $null
        $temp | Add-Member -MemberType NoteProperty -Name OS -Value $null
        $temp | Add-Member -MemberType NoteProperty -Name Installed -Value $null
        $temp | Add-Member -MemberType NoteProperty -Name Code -Value $null

        #Check if online
        Write-Debug "Testing if $comp online."
        Try
        {
            $online = Test-Connection $comp -Quiet -ErrorAction Stop
        }
        Catch
        {
            $online = $false
        }

        Write-Debug "Online: $online"

        If ($online -eq $false)
        {
            Write-Debug "Computer Offline"
            $temp.Name = "Offline"
            $temp.OS = "Offline"
            $temp.Installed = "Offline"
            $temp.Code = 90
            Write-Debug $temp
            Continue
        }

        Try
        {
            Write-Debug "Getting hostname"
            $CompSys = GWMI -Class Win32_ComputerSystem -ComputerName $comp -ErrorAction Stop
            $temp.Name = $($CompSys.Name)
            Write-Debug "Hostname: $($CompSys.Name)"
        }
        Catch
        {
            Write-Debug "Unable to determine hostname"
            $temp.Name = "Unknown"
        }

        Try
        {
            $OSInfo = GWMI -Class Win32_OperatingSystem -ComputerName $comp -ErrorAction Stop
            Write-Debug "OS Version: $($OSInfo.Version)"
        }
        Catch
        {
            Write-Debug "Failed to get OS info"
            $temp.OS = "Unknown"
            $temp.Installed = "Skipped"
            $temp.Code = 900
            Write-Debug $temp
            $Output += $temp
            Continue
        }
            
        Switch ($OSInfo.Version[0])
        {
            '6' #Windows 7
            {
                Write-Debug "Windows 7"
                $temp.OS = "Windows 7"
                $av = $Win7Antivirus
                $avPath = "$LocalPath\SCEPInstall.exe"
                $argList = "/s", "/q"

                #Check if already installed
                Try
                {
                    $PreCheck = GWMI -cn $comp -query "SELECT * FROM Win32_Product WHERE Name LIKE '%Endpoint Protection%'" -ErrorAction Stop
                    If ($PreCheck.InstallState -eq 5)
                    {
                        Write-Debug "SCEP already installed"
                        $temp.Installed = "Installed"
                        $temp.Code = 0
                    }
                    Else { Write-Debug "Not installed yet" }
                }
                Catch { Write-Debug "PreCheck failed" }
            }
                
            '5' #Windows XP
            {
                Write-Debug "Windows XP"
                $temp.OS = "Windows XP"
                $av = $XPAntivirus
                $avPath = "$LocalPath\mseinstall.exe"
                $argList = "/s", "/runwgacheck", "/o"

                #Check if already installed
                Try
                {
                    $PreCheck = GWMI -cn $comp -query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Security Client%'" -ErrorAction Stop
                    If ($PreCheck.InstallState -eq 5)
                    {
                        Write-Debug "MSE already installed"
                        $temp.Installed = "Installed"
                        $temp.Code = 0
                    }
                    Else { Write-Debug "Not installed yet" }
                }
                Catch { Write-Debug "PreCheck failed" }
            }
                
            Default #Unexpected OS
            {
                Write-Debug "Unexpected OS"
                $temp.OS = "Unknown"
                $temp.Installed = "Skipped"
                $temp.Code = 999
            }
        }

        #Skip if already installed or OS Unknown
        If (($temp.Installed -eq "Installed") -or ($temp.OS -eq "Unknown"))
        {
            $Output += $temp
            Write-Debug $temp
            Continue
        }
        
        #Copy file to local machine
        Try
        {
            Write-Debug "Copying installer to local machine: $comp"
            Write-Debug "AV Source Path: $av"
            Write-Debug "AV Local Path: \\$comp\$LocalPath"

            $ErrCode = 0

            #Create Temp directory
            Write-Debug "Creating C:\Temp directory"
            Try { $discard = New-Item -Path "\\$comp\c$\" -Name "Temp" -ItemType Directory -ErrorAction Stop }  Catch { $ErrCode += 1 }
            #Create Antivirus sub-directory
            Write-Debug "Creating C:\Temp\Antivirus directory"
            Try { $discard = New-Item -Path "\\$comp\c$\Temp\" -Name "Antivirus" -ItemType Directory -ErrorAction Stop } Catch { $ErrCode += 2}

            $discard = Copy-Item -Path $av -Destination "\\$comp\$LocalPath" -ErrorAction Stop
        }
        Catch
        {
            Write-Debug "Unable to copy to destination"
            $temp.Installed = "Failed"
            $temp.Code = "1.$ErrCode"
            $Output += $temp
            Write-Debug $temp
            Continue
        }

        #Initiate remote install
        Try
        {
            Write-Debug "Initiating remote install"
            $InstallString = "\\$comp\$avPath $argList"
            Write-Debug "InstallString: $InstallString"

            $Process = Invoke-WmiMethod -ComputerName $comp -Class Win32_Process -Name Create -ArgumentList $InstallString

            If ($Process.ReturnValue -eq 0)
            {
                $temp.Installed = "Installing"
                $temp.Code = 0
                Write-Debug "Remote install started succesfully"
            }
            Else
            {
                $temp.Installed = "Failed"
                $temp.Code = "Error: $($Process.ReturnValue)"
                Write-Debug "Remote install failed. Error code: $($Process.ReturnValue)"
            }
        }
        Catch
        {
            Write-Debug "Remote install failed"
            $temp.Installed = "Failed"
            $temp.Code = 2
        }

        $Output += $temp
        Write-Debug $temp
    }
}

End
{
    ForEach ($machine in $Output)
    {
        #If installer started successfully go back and see that software install completed
        Write-Debug "Checking installation success"
        If ($machine.Installed -eq "Installing")
        {
            If ($machine.OS -eq "Windows 7")
            {
                Try
                {
                    $InstallCheck = GWMI -cn $($machine.ID) -query "SELECT * FROM Win32_Product WHERE Name LIKE '%Endpoint Protection%'" -ErrorAction Stop
                }
                Catch { Write-Debug "Install Check failed" }

                If ($InstallCheck.InstallState -eq 5)
                {
                    $machine.Installed = "Success"
                    Write-Debug "Install Success"
                }
                Else
                {
                    $machine.Installed = "Failed"
                    $machine.Code = "State: $($InstallCheck.InstallState)"
                    Write-Debug "Install failure"
                }
            }
            ElseIf ($machine.OS -eq "Windows XP")
            {
                Try
                {
                $InstallCheck = GWMI -cn $($machine.ID) -query "SELECT * FROM Win32_Product WHERE Name LIKE 'Microsoft Security Client%'" -ErrorAction Stop
                }
                Catch { Write-Debug "Install Check failed" }

                If ($InstallCheck.InstallState -eq 5)
                {
                    $machine.Installed = "Success"
                    Write-Debug "Install Success"
                }
                ElseIf ($InstallCheck -ne $null)
                {
                    $machine.Installed = "Failed"
                    $machine.Code = "State: $($InstallCheck.InstallState)"
                    Write-Debug "Install failure"
                }
                Else
                {
                    $machine.Installed = "Failed"
                    $machine.Code = 99
                    Write-Debug "Install failure"
                }
            }
        }
    }
    Return $Output
}