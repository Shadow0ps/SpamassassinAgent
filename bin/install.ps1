
# Check and see if already installed
if((Test-Path -Path "C:\Program Files (x86)\SpamAssassin\" )){
    $message  = 'Warning'
    $question = "A SpamAssassin installation has been detected already. If you continue all files (including settings) will be overwritten.
    
Are you sure you want to proceed?"

    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))

    $decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
    if ($decision -eq 0) {

        # Stop spamd if it exists
        Stop-Service spamd -ErrorAction Ignore

        # Delete spamd if it exists
        $service = Get-WmiObject -Class Win32_Service -Filter "Name='spamd'"
        if($service) {
            $service.delete()
        }

        Remove-Item -Recurse -Force "C:\Program Files (x86)\SpamAssassin\"

        schtasks.exe /delete /tn "SpamAssassin AutoUpdate" /F

    } else {
        Exit
    }
}

# Download
Invoke-WebRequest http://www.jam-software.com/spamassassin/SpamAssassinForWindows.zip -OutFile C:\Windows\Temp\SpamAssassinForWindows.zip

# Create directory if it doesn't exist
New-Item "C:\Program Files (x86)\SpamAssassin\" -type Directory

# Unzip
$shell = new-object -com shell.application
$zip = $shell.NameSpace(“C:\Windows\Temp\SpamAssassinForWindows.zip”)
foreach($item in $zip.items())
{
    $shell.Namespace("C:\Program Files (x86)\SpamAssassin\").copyhere($item)
}

#TODO: Get latest version of default configs from the github repo

# Create the service
New-Service -BinaryPathName "C:\Windows\System32\srvany.exe" -Name spamd -DisplayName "SpamAssassin Daemon" 
New-Item -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd -Name "Parameters"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "Application" -PropertyType STRING -Value "C:\Program Files (x86)\SpamAssassin\spamd.exe"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "AppDirectory" -PropertyType STRING -Value "C:\Program Files (x86)\SpamAssassin\"
New-ItemProperty -Path HKLM:\SYSTEM\CurrentControlSet\services\spamd\Parameters -Name "AppParameters" -PropertyType STRING -Value "-x -l -s spamd.log"

# Create the system task to run the update script every night
schtasks.exe /create /tn "SpamAssassin AutoUpdate" /tr "'C:\Program Files (x86)\SpamAssassin\sa-update.bat'" /sc DAILY /st 02:00 /RU SYSTEM /RL HIGHEST /v1
schtasks.exe /run /tn "SpamAssassin AutoUpdate"

# wait about 30 seconds for the update to complete
Start-Sleep -Seconds 30

# Download the SpamAssassin config file
Invoke-WebRequest https://raw.githubusercontent.com/jmdevince/SpamassassinAgent/master/contrib/spamassassin/local.cf -OutFile "C:\Program Files (x86)\SpamAssassin\etc\spamassassin\local.cf"

# Start the spamd service
Start-Service spamd

# Now let's download and install Spamassasssin Transport Agent, Ask what version of exchange they are using
$Version = ""

[int]$xMenuChoiceA = 0
while ( $xMenuChoiceA -lt 1 -or $xMenuChoiceA -gt 4 ){
Write-Host "Please select the version of Microsoft Exchange Server you are using..."
Write-host "1. 2016 Service Pack 0"
Write-host "2. 2013 Service Pack 1"
Write-host "3. 2010 Service pack 3"
Write-host "4. 2010 Service Pack 2"
[Int]$xMenuChoiceA = read-host "Please enter an option 1 to 4..." }
Switch( $xMenuChoiceA ){
  1{$version = "2016SP0"}
  2{$version = "2013SP1"}
  3{$version = "2010SP3"}
  4{$version = "2010SP2"}
default{Write-Host "Invalid Selection"}
}

# Create Directory to save to
New-Item "C:\CustomAgents\" -type Directory

# Create Data Directory
New-Item "C:\CustomAgents\SpamassassinAgentData" -type Directory

# Download the proper DLL
Invoke-WebRequest https://raw.githubusercontent.com/jmdevince/SpamassassinAgent/master/bin/$version/SpamassassinAgent.dll -OutFile "C:\CustomAgents\SpamassassinAgent.dll"

# Download the XML configuration
Invoke-WebRequest https://raw.githubusercontent.com/jmdevince/SpamassassinAgent/master/etc/SpamassassinConfig.xml -OutFile "C:\CustomAgents\SpamassassinAgentData\SpamassassinConfig.xml"

# Connect to the exchange Server
. 'C:\Program Files\Microsoft\Exchange Server\V14\bin\RemoteExchange.ps1'
Connect-ExchangeServer -auto

# Install the Transport Agent
Install-TransportAgent -Name "SpamAssassin Agent" -AssemblyPath C:\CustomAgents\SpamassassinAgent.dll -TransportAgentFactory SpamassassinAgent.SpamassassinAgentFactory
Enable-TransportAgent "Spamassassin Agent"
Set-TransportAgent "Spamassassin Agent" -Priority 3

# Restart 
Restart-Service MSExchangeTransport