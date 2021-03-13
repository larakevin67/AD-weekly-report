$time = (Get-Date).Adddays(-(90)) 
$filt = {LastLogonTimeStamp -lt $time -and enabled -eq $true}
$Computers90 = Get-ADComputer -Filter $filt -properties * 
$Users90 = Get-ADUser -Filter $filt -properties * 
$DisabledUsers = Search-ADAccount -AccountDisabled -UsersOnly | Where-Object {$_.distinguishedname -notlike "*OU=Disabled Users,DC=BGO,DC=local*"} 
$DisabledComputers = Search-ADAccount -AccountDisabled -ComputersOnly | Where-Object {$_.distinguishedname -notlike "*OU=Disabled Computers,DC=BGO,DC=local*"} 

$smtp = "smtp server"
$Password = '**********'
$User = "sample@email.com"
$SecurePassword = $Password | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($User, $SecurePassword)
$to = "email"
$from = "email"
$subject = "AD Weekly Report"

foreach($Computer in $Computers90){
    $name = $Computer.Name 
    $passex = $Computer.Ipv4address
    $passnevex = $Computer.OperatingSystem
    $wcreat = $Computer.whenCreated
    $passlaset = $Computer.WhenChanged
    $lalog = $Computer.LastLogonDate
    $pinging = Test-Connection $name -Count 1 -BufferSize 2 -Quiet -ErrorAction SilentlyContinue
    $dataRow =
    "
    </tr><td>$name</td>
    <td>$passex</td>
    <td>$passnevex</td>
    <td>$wcreat</td>
    <td>$passlaset</td>
    <td>$lalog</td>
    <td>$pinging</td>
    "
    $Computerreport += $datarow  
    }

foreach($user in $Users90){
    $name = $user.Name 
    $display = $user.SamAccountName
    $upn = $user.UserPrincipalName
    $descript = $user.Description
    $wcreat = $user.Created
    $lalog = $user.LastLogonDate 
    $passlaset = $user.PasswordLastSet
    $notexp = $user.PasswordNeverExpires
    $dataRow =
    "
    </tr><td>$name</td>
    <td>$display</td>
    <td>$upn</td>
    <td>$descript</td>
    <td>$wcreat</td>
    <td>$lalog</td>
    <td>$passlaset</td>
    <td>$notexp</td>
    "
    $userreport += $datarow
    
    }

foreach ($Disableduser in $Disabledusers) {
    $dname = $Disableduser.Name
    $type = $Disableduser.objectclass
    $upn = $Disableduser.distinguishedname
    $dataRow =
    "
    </tr><td>$dname</td>
    <td>$type</td>
    <td>$upn</td>
    "
    $disableduserreport += $datarow
    
    }
foreach ($DisabledComputer in $DisabledComputers) {
    $dname = $DisabledComputer.Name
    $type = $DisabledComputer.objectclass
    $upn = $DisabledComputer.distinguishedname
    $dataRow =
    "
    </tr><td>$dname</td>
    <td>$type</td>
    <td>$upn</td>
    "
    $disabledcompreport += $datarow
    }

$html = 
"<html><style>{font-family: Arial; font-size: 13pt;}TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt;}
TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}TD{border: 1px solid black; padding: 5px; }</style>
<h2>Users not Logged in 90+ Days</h2><table>
<tr><th>Name</th>
<th>SamAccountName</th>
<th>UPN</th>
<th>Description</th>
<th>When Created</th>
<th>Last Login</th>
<th>Password Last Set</th>
<th>Password Never Expires</th>
</tr>$userreport</table><tr>

<html><style>{font-family: Arial; font-size: 13pt;}TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt;}
TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}TD{border: 1px solid black; padding: 5px; }</style>
<h2>Computers not Logged in 90+ Days</h2><table>
<tr><th>Name</th>
<th>IP Address</th>
<th>Operating System</th>
<th>When Created</th>
<th>Password Last Set</th>
<th>Last Login</th>
<th>Online</th>
</tr>$Computerreport</table><tr>

<html><style>{font-family: Arial; font-size: 13pt;}TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt;}
TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}TD{border: 1px solid black; padding: 5px; }</style>
<h2>Disabled User Accounts</h2><table>
<tr><th>Name</th>
<th>Object Type</th>
<th>Orginal OU</th>
</tr>$disableduserreport</table><tr>

<html><style>{font-family: Arial; font-size: 13pt;}TABLE{border: 1px solid black; border-collapse: collapse; font-size:12pt;}
TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}TD{border: 1px solid black; padding: 5px; }</style>
<h2>Disabled Computer Accounts</h2><table>
<tr><th>Name</th>
<th>Object Type</th>
<th>Orginal OU</th>
</tr>$disabledcompreport</table><tr>"

send-MailMessage -SmtpServer $smtp -To $to -From $from -Subject $subject -Body $html -usessl -Credential $cred -Port 587
