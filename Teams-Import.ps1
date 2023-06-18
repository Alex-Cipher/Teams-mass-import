<#
Add-Type -assemblyname Microsoft.VisualBasic
Add-Type -assemblyname system.windows.forms 
Add-Type –assemblyName WindowsBase
#>

Connect-MicrosoftTeams 

#$file = Read-Host "Bitte vollstaendigen zur Datei angeben"
$file = Read-Host "full path to your csv file (like C:\teams\team.csv)"
Write-Host -ForegroundColor Yellow $file
$data = Import-Csv -Path $file -Delimiter ";" 
$data | Format-Table

$startTime = Get-Date -DisplayHint Date


foreach ($row in $data) {

    
    $emailstrg = $row.TeamName -replace '\s+', ''
    $emailstrg = $emailstrg.Replace(',','').Replace('-','')

    if(-not [string]::IsNullOrWhiteSpace($row.TeamName)){
        $GroupID = (New-Team -DisplayName $row.TeamName -MailNickName $emailstrg -Visibility Private).GroupID
    } else {
        break
    }
}

#Write-Host -ForegroundColor Green $row.TeamName.Count " Team(s) wurde(n) erstellt" 

foreach ($row in $data){

    $emailstrg = $row.TeamName -replace '\s+', ''
    $emailstrg = $emailstrg.Replace(',','').Replace('-','')

    if(-not [string]::IsNullOrWhiteSpace($row.TeamName)){
        Write-Host -ForegroundColor Magenta "Team created: " $row.TeamName
        Write-Host -ForegroundColor Magenta "E-Mail address: " $emailstrg"@blabla.tld"
    } else {
        break
    }
}

foreach ($row in $data) {

    if(-not [string]::IsNullOrWhiteSpace($row.Channel)){
        New-TeamChannel -GroupId $GroupID -DisplayName $row.Channel -MembershipType $row.Type
    } else {
        break
    }
    
    Write-Host -ForegroundColor Green "channel: " $row.Channel
    #Write-Progress -Activity "Kanäle werden erstellt" -Status "Komplett:" -PercentComplete -1 -SecondsRemaining -1 -CurrentOperation "Kanal wird erstellt"
}

foreach ($row in $data) {

    if(-not [string]::IsNullOrWhiteSpace($row.User)){
        Add-TeamUser -GroupId $GroupId -User $row.User -Role $row.Role
    } else {
        break
    }
        
    Write-Host -ForegroundColor Green   "user: " $row.User  
}

foreach ($row in $data) {

    Add-TeamChannelUser -GroupId $GroupID -DisplayName $row.ChannelName -User $row.ChannelUser

    Write-Host -ForegroundColor Green "channel: " $row.ChannelName "   user: " $row.ChannelUser  
}

foreach ($row in $data) {

    
    if($row.ChannelRole.Contains("Member")){
        continue
    } elseif($row.ChannelRole.Contains("owner")){
        Add-TeamChannelUser -GroupId $GroupID -DisplayName $row.ChannelName -User $row.ChannelUser -Role $row.ChannelRole
    } else {
        break
    }
    

    Write-Host -ForegroundColor Green "channel:  " $row.ChannelName "   user:  " $row.ChannelUser  " role:  "  $row.ChannelRole
}
<#
foreach ($row in $data) {

    Get-TeamChannelUser -GroupID $GroupID -DisplayName $row.Channel

}
#>

$endTime = Get-Date -DisplayHint Date
$totalTime = $endTime - $startTime
Write-Host -ForegroundColor Magenta "duration for import: " $totalTime

Disconnect-MicrosoftTeams

#[System.Windows.Forms.MessageBox]::Show("Bitte Enter drücken um das Fenster zu schliessen","Import komplett", [System.Windows.Forms.MessageBoxButtons]::OK)
Read-Host "Bitte Enter drücken um das Fenster zu schliessen"