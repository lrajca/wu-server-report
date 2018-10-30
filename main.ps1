#Full Canonical name of Server OU's to query
$bases = ""

#Path to Credentials stored as a CliXML File
$credentialPath = ""

#End variable declaration


$Credentials = Import-Clixml -Path $credentialPath

$noConnect = @('')
$arr = @('')
$Count = 0
$Iterations = 0
$Computers = @('')


Foreach ($base in $bases){ 

$Computers += Get-AdComputer -Filter *  -SearchBase $base -Properties OperatingSystem

}

$Script = {


$Session = New-Object -ComObject Microsoft.Update.Session            
$Searcher = $Session.CreateUpdateSearcher()         
$HistoryCount = $Searcher.GetTotalHistoryCount()                    
$searchSession = ($Searcher.QueryHistory(0,$HistoryCount) | Sort-Object -Descending:$true -Property Date | Select-Object -Property * -First 1)

$Result = ''
Switch ($searchSession.ResultCode)
{
	0 { $Result = 'Not Started'}
	1 { $Result = 'In Progress' }
	2 { $Result = 'Succeeded ' }
	3 { $Result = 'Succeeded With Errors' }
	4 { $Result = 'Failed' }
	5 { $Result = 'Aborted' }
	default { $Result = '' }
	
}

$RegServer = Get-ItemProperty HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate | Select-Object -Expand WUServer 
$Server = ($RegServer.split('-'))[1].ToUpper()

    
$HotFix = Get-HotFix
$ComName = $env:ComputerName
#$IP = (Resolve-DNSName ($ComName) | Select-Object -Expand IpAddress)[1]
#$ServerType = (Get-WmiObject win32_operatingsystem).caption
$HotFix = $HotFix.HotFixID[-1]
$InstalledOn = ((Get-HotFix).InstalledOn[-1])


New-Object PSObject -Property @{

'Server' = $ComName; 
'IP Address' = Resolve-DNSName ($ComName) -Type A | Select-Object -Expand IpAddress
"Operating System" = (Get-WmiObject win32_operatingsystem).caption; 
"Update Name" = $searchSession.Title; 
"Update Result" = $Result;
"Last Update" = ($searchSession.Date).ToShortDateString(); 
"Restart Pending" = (Test-Path 'HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired');
"WSUS Server" = $Server;}

}

$Output = @('')



Foreach ($Computer in $Computers){

    $Com = $Computer.Name
    $ComCount = ($Computers.Count)
    If (Test-Connection $Computer.name -Count 1){

    Try {
        $Output = Invoke-Command -ComputerName $Computer.name -Credential $Credentials -ScriptBlock $Script -ErrorAction Stop | Select-Object 'Server', 'IP Address', 'Operating System', 'Last Update', 'Update Name', 'Update Result', 'Restart Pending', 'WSUS Server' | Export-CSV -Append .\UpdateReport4.csv
        } 
        
   Catch {
        $noConnect += $Com}
        }

    Else {
    $Com = $Computer.Name
    $noConnect += $Com
    Write-Host "Could not connect to: $Com"
    }

    $Count += 1
    
    Write-Progress -Activity "Building Report..." -Status "$Count/$ComCount Servers" -PercentComplete ($Count/$ComCount*100)
}


    