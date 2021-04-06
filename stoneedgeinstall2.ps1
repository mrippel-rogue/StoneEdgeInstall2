########################Rogue Fitness Stone Edge Powershell Installer########################
##AUTHOR: Jarrett Allman
##DATE: 3/26/2021
##DESCRIPTION: This script is a simple automated deployment routine to install Stone Edge on a Rogue Fitness PC. This automates the deployment of the ODBC connection,
##Ghostscript, and other dependencies as they are currently written.

####Create logfile for installer, output to .\logfile.txt
Start-Transcript -path C:\temp\logfile.txt -Append -NoClobber -IncludeInvocationHeader -Verbose

####Prompt user for input
Write-Host 'Welcome to the Rogue StoneEdge PowerShell installer.'
Write-Host 'For any feedback, please e-mail jallman@roguefitness.com'
Read-Host 'Press any key to continue...'

Write-Host -ForegroundColor green 'Beginning dependency checks...'

####Checks if Access is installed.
$AccessPath = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.exe"
$AccessInstalled = Test-Path $AccessPath

If($AccessInstalled -ne 'True'){
    Write-Host -ForegroundColor red -BackgroundColor white 'MS Access 2010 not installed. Access must be installed to launch and license StoneEdge. Installation will not proceed.'
    pause
    exit
} ElseIf ($GSInstalled -eq 'True')
{
    Write-Host 'Microsoft Access 2010 found, proceeding...'
}

###Silently installs DataCap files if it has not already been installed.
$DatacapPath = C:\Windows\DatacapControls\dsiEMVX.ocx
$DatacapPathInstalled = Test-Path $DatacapPath

If($DatacapPathInstalled -ne 'True'){
    Write-Host 'Installing Datacap filesinstalled...'
    Copy-Item -Path "\\fileserver\public\_RogueIT\Stoneedge V7.024\dsiEMVUS-130-Install20190320-W8.exe" -Destination "C:\Windows\DatacapControls" -Recurse
} ElseIF ($DatacapInstalled -eq 'True')
{
   Write-Host -ForegroundColor green 'Datacap already installed...'
}
###Silently installs GhostScript if it has not already been installed.
$GSPath = "C:\Program Files (x86)\gs\gs9.16\bin\gswin32.exe"
$GSInstalled = Test-Path $GSPath

If($GSInstalled -ne 'True'){
    Write-Host 'Installing GhostScript files...'
    Copy-Item -Path "\\fileserver\public\_RogueIT\Stoneedge V7.024\gs9.16" -Destination "C:\Program Files (x86)\gs\gs9.16" -Recurse
} ElseIf ($GSInstalled -eq 'True')
{
    Write-Host -ForegroundColor green 'GhostScript already installed...'
}

###Silently installs the SQL native client if it has not already been installed.
$SQLNCPath = "C:\Windows\System32\sqlncli10.dll"
$SQLNCInstalled = Test-Path $SQLNCPath

If($SQLNCInstalled -ne 'True'){
    Write-Host 'Installing SQL Native Client 10.0...'
    Start-Process msiexec.exe -Wait -ArgumentList '/i "\\fileserver\public\_RogueIT\SQLNativeClient\sqlnclix64.msi" /qn'
} ElseIf ($SQLNCInstalled -eq 'True')
{
    Write-Host -ForegroundColor green 'SQL Native Client already installed...'
}

###Silently install StoneEdge if it does not already exist.
$SEPath = "C:\StoneEdge\SEOrdMan.MDB"
$SEInstalled = Test-Path $SEPath

If($SEInstalled -ne 'True'){
    Write-Host 'StoneEdge not found, installing. Please wait...'
    Start-Process msiexec.exe -Wait -ArgumentList '/i "\\fileserver\public\_RogueIT\Stoneedge V7.024\StoneEdgeInstaller.msi" /qn'
} ElseIf ($SEInstalled -eq 'True')
{
    Write-Host -ForegroundColor green 'StoneEdge already installed...'
}

Write-Host 'Beginning ODBC connector configuration...'

####Prompt user to select locale.
$NA = New-Object System.Management.Automation.Host.ChoiceDescription '&North America', 'North America'
$EU = New-Object System.Management.Automation.Host.ChoiceDescription '&Europe', 'Europe'
$AU = New-Object System.Management.Automation.Host.ChoiceDescription '&Australia', 'Australia'
$RE = New-Object System.Management.Automation.Host.ChoiceDescription '&Retail', 'NA Retail'

$options = [System.Management.Automation.Host.ChoiceDescription[]]($NA, $EU, $AU, $RE)

$title = 'Locale Selection'
$message = 'What locale are you installing StoneEdge for?'
$result = $host.ui.PromptForChoice($title, $message, $options, 0)
    switch ($result)
    {
      0 { 'NA' }
      1 { 'EU' }
      2 { 'AU' }
      3 { 'Retail' }
    }


####Prompt for NA environment selection.
If ($result -eq '0')
{
    $NAprod = New-Object System.Management.Automation.Host.ChoiceDescription '&Production', 'Production'
    $NAdev = New-Object System.Management.Automation.Host.ChoiceDescription '&Development', 'Development'
    $NAtest1 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &1'
    $NAtest2 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &2'
    $NAtest3 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &3'
    $NAtest4 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &4'

    $NAoptions = [System.Management.Automation.Host.ChoiceDescription[]]($NAprod, $NAdev, $NAtest1, $NAtest2, $NAtest3, $NAtest4)

    $NAtitle = 'Environment'
    $NAmessage = 'What NA environment are you installing StoneEdge for?'
    $NAresult = $host.ui.PromptForChoice($NAtitle, $NAmessage, $NAoptions, 0)
    switch ($NAresult)
    {
      0 { 'Prod' }
      1 { 'Dev' }
      2 { 'Test 1' }
      3 { 'Test 2' }
      4 { 'Test 3' }
      5 { 'Test 4' }
    }
}
####Prompt for EU environment selection.
ElseIf ($result -eq '1')
{
    $EUprod = New-Object System.Management.Automation.Host.ChoiceDescription '&Production'
    $EUdev = New-Object System.Management.Automation.Host.ChoiceDescription '&Development'
    $EUtest1 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &1'
    $EUtest2 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &2'
    $EUtest3 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &3'
    $EUtest4 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &4'

    $EUoptions = [System.Management.Automation.Host.ChoiceDescription[]]($EUprod, $EUdev, $EUtest1, $EUtest2, $EUtest3, $EUtest4)

    $EUtitle = 'Environment'
    $EUmessage = 'What EU environment are you installing StoneEdge for?'
    $EUresult = $host.ui.PromptForChoice($EUtitle, $EUmessage, $EUoptions, 0)
    switch ($EUresult)
    {
      0 { 'Prod' }
      1 { 'Dev' }
      2 { 'Test 1' }
      3 { 'Test 2' }
      4 { 'Test 3' }
      5 { 'Test 4' }
    }
}
####Prompt for AU environment selection.
ElseIf ($result -eq '2')
{
    $AUprod = New-Object System.Management.Automation.Host.ChoiceDescription '&Production'
    $AUdev = New-Object System.Management.Automation.Host.ChoiceDescription '&Development'
    $AUtest1 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &1'
    $AUtest2 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &2'
    $AUtest3 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &3'
    $AUtest4 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &4'

    $AUoptions = [System.Management.Automation.Host.ChoiceDescription[]]($AUprod, $AUdev, $AUtest1, $AUtest2, $AUtest3, $AUtest4)

    $AUtitle = 'Environment'
    $AUmessage = 'What AU environment are you installing StoneEdge for?'
    $AUresult = $host.ui.PromptForChoice($AUtitle, $AUmessage, $AUoptions, 0)
    switch ($AUresult)
    {
      0 { 'Prod' }
      1 { 'Dev' }
      2 { 'Test 1' }
      3 { 'Test 2' }
      4 { 'Test 3' }
      5 { 'Test 4' }
    }
}
####Prompt for retail environment selection.
ElseIf ($result -eq '3')
{
    $Retailprod = New-Object System.Management.Automation.Host.ChoiceDescription '&Production'
    $Retaildev = New-Object System.Management.Automation.Host.ChoiceDescription '&Development'
    $Retailtest1 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &1'
    $Retailtest2 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &2'
    $Retailtest3 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &3'
    $Retailtest4 = New-Object System.Management.Automation.Host.ChoiceDescription 'Test &4'

    $Retailoptions = [System.Management.Automation.Host.ChoiceDescription[]]($Retailprod, $Retaildev, $Retailtest1, $Retailtest2, $Retailtest3, $Retailtest4)

    $Retailtitle = 'Environment'
    $Retailmessage = 'What Retail environment are you installing StoneEdge for?'
    $Retailresult = $host.ui.PromptForChoice($Retailtitle, $Retailmessage, $Retailoptions, 0)
    switch ($Retailresult)
    {
      0 { 'Prod' }
      1 { 'Dev' }
      2 { 'Test 1' }
      3 { 'Test 2' }
      4 { 'Test 3' }
      5 { 'Test 4' }
    }
}

###ODBC connector boolean logic. Sets up the appropriate ODBC connector based on the user's selection(s).
If ($result -eq '0' -And $NAresult -eq '0') ###Install for NA Prod
{
    Write-Host 'Installing for NA Prod'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-PROD.reg
}
ElseIf ($result -eq '0' -And $NAresult -eq '1')
{
    Write-Host 'Installing for NA Dev'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-DEV.reg
}
ElseIf ($result -eq '0' -And $NAresult -eq '2')
{
    Write-Host 'Installing for NA Test 1'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-TEST1.reg
}
ElseIf ($result -eq '0' -And $NAresult -eq '3')
{
    Write-Host 'Installing for NA Test 2'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-TEST2.reg
}
ElseIf ($result -eq '0' -And $NAresult -eq '4')
{
    Write-Host 'Installing for NA Test 3'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-TEST3.reg
}
ElseIf ($result -eq '0' -And $NAresult -eq '5')
{
    Write-Host 'Installing for NA Test 4'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-NA-TEST4.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '0')
{
    Write-Host 'Installing for EU Prod'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-PROD.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '1')
{
    Write-Host 'Installing for EU Dev'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-DEV.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '2')
{
    Write-Host 'Installing for EU Test 1'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-TEST1.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '3')
{
    Write-Host 'Installing for EU Test 2'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-TEST2.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '4')
{
    Write-Host 'Installing for EU Test 3'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-TEST3.reg
}
ElseIf ($result -eq '1' -And $EUresult -eq '5')
{
    Write-Host 'Installing for EU Test 4'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-EU-TEST4.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '0')
{
    Write-Host 'Installing for AU Prod'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-PROD.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '1')
{
    Write-Host 'Installing for AU Dev'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-DEV.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '2')
{
    Write-Host 'Installing for AU Test 1'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-TEST1.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '3')
{
    Write-Host 'Installing for AU Test 2'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-TEST2.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '4')
{
    Write-Host 'Installing for AU Test 3'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-TEST3.reg
}
ElseIf ($result -eq '2' -And $AUresult -eq '5')
{
    Write-Host 'Installing for AU Test 4'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-AU-TEST4.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '0')
{
    Write-Host 'Installing for Retail Prod'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-PROD.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '1')
{
    Write-Host 'Installing for Retail Dev'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-DEV.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '2')
{
    Write-Host 'Installing for Retail Test 1'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-TEST1.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '3')
{
    Write-Host 'Installing for Retail Test 2'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-TEST2.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '4')
{
    Write-Host 'Installing for Retail Test 3'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-TEST3.reg
}
ElseIf ($result -eq '3' -And $Retailresult -eq '5')
{
    Write-Host 'Installing for Retail Test 4'
    reg import \\fileserver\public\_RogueIT\SE_ODBC_REG_KEYS\SE-RETAIL-TEST4.reg
}
Else
{
    Write-Host 'No Valid Selection Made'
    exit
}
###Silently installs StoneEdge Golden file if it has not already been installed.
$GoldenImage = C:\StoneEdge\!SE-Golden-updated-04012021.txt
$GoldenImageInstalled = Test-Path $GoldenImage

If($GoldenImageInstalled -eq 'True'){
    Write-Host 'Copying GoldenImage file...'
    Copy-Item -Path "\\fileserver\public\_RoguIT\Stoneedge V7.-24\!Golden-04-01-2021" -Destination "C;\StoneEdge" -Recurse

}ElseIf ($GoldenImageInstalled -eq 'True')
{
Write-Host -ForegroundColor green 'GoldenImage already installed...'
}

Write-Host 'Press any key to launch StoneEdge and complete installation.'
pause
notepad.exe \\fileserver\public\_RogueIT\Stoneedge V7.024\LICENSE.txt

###Checks for the existence of MS Access
$AccessPath = "C:\Program Files (x86)\Microsoft Office\Office14\MSACCESS.exe"
$AccessInstalled = Test-Path $AccessPath

If($AccessInstalled -ne 'True'){
    Read-Host 'MS Access 2010 not installed. Access must be installed to launch and license StoneEdge.'
    pause
} ElseIf ($GSInstalled -eq 'True')
{
    Write-Host 'Launching StoneEdge to complete licensing and setup...'
    C:\StoneEdge\SEOrdMan.MDB
}

Read-Host 'Install complete, press any key to continue...'

Stop-Transcript
