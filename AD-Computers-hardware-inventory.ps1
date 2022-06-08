$Computers = Get-ADComputer -filter * | Select-Object -ExpandProperty Name

foreach ($Computer in $Computers) {
    if(!(Test-Connection -Cn $computer -BufferSize 16 -Count 1 -ea 0 -quiet)) {
            write-host "cannot reach $computer offline" -f red
    }
    else {
        $ComputersInfoList = @()
        try {
            $SN=Get-WmiObject win32_bios -ComputerName $Computer  -ErrorAction Stop
            $CPU=Get-WmiObject –class Win32_processor -ComputerName $computer -ErrorAction Stop   
            $AdInfo=get-adcomputer $computer -properties Name,Lastlogondate,operatingsystem,description -ErrorAction Stop
            $RAM="{0} GB" -f ((Get-WmiObject Win32_PhysicalMemory -ComputerName $computer | Measure-Object Capacity  -Sum).Sum / 1GB)
            $DISK="{0} GB" -f [math]::Round((Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Measure-Object Size -Sum).Sum /1GB)
            $x = gwmi win32_computersystem -ComputerName $computer | select Manufacturer,@{Name = "Model";Expression = {if (($_.model -eq "$null")  ) {'Virtual'} Else {$_.model}}},username -ErrorAction Stop
            $usr_temp = $x.username.substring(9)
            $usr = Get-ADUser -Filter {SamAccountName -eq $usr_temp}
            $temp = New-Object PSObject -Property @{
                SerialNumber = $SN.SerialNumber
                ComputerName = $AdInfo.name
                Description=$AdInfo.description
                Manufacturer=$x.Manufacturer
                Model=$x.Model
                ProcessorName=($CPU.name | Out-String).Trim()
                Ram=$RAM
                Disk=$DISK
                OperatingSystem=$AdInfo.operatingsystem
                LastLogonDate=$AdInfo.lastlogondate
                LoggedinUser=$x.username
                UserFullName=$usr.Name
                }
                $ComputersInfoList += $temp
        }
        catch {
            "Error communicating with $computer" 
        }
        $ComputersInfoList | select Computername,Manufacturer,Model,SerialNumber,Description,OperatingSystem,Ram,Disk,ProcessorName,LoggedinUser,UserFullName,LastLogondate | export-csv -Append C:\Users\rs-a\Desktop\AD-inventory.csv -nti -Encoding UTF8
    }
