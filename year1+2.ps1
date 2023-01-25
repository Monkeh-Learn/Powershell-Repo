#initialise functions

function Test-ADOU 
{
    param($Name)
    $Exist = $true
    $TestOU = "OU= $Name ,DC=contoso,DC=com"
    try {
        Get-ADOrganizationalUnit -Identity $TestOU -ErrorAction Stop | Out-Null
    }
    catch {
        $Exist = $false
    }
    Write-Output -InputObject $Exist
}
Test-ADOU

function Test-ADUser 
{
    param($Name)
    $Exist = $true
    try {
        Get-ADUser -Identity $Name -ErrorAction Stop | Out-Null
    }
    catch {
        $Exist = $false
    }
    Write-Output -InputObject $Exist
}
Test-ADUser

function File-Picker
{
    Add-Type -AssemblyName system.windows.forms

    $File = New-Object System.Windows.Forms.OpenFileDialog

    $File.InitialDirectory = "C:\Users\Administrator.CONTOSO\Documents\CSV Project"

    $File.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.csv"

    $result = $File.ShowDialog()

    



     $File.FileName
}

#Begin Script

#Import CSV with users, through pop-up dialogue - problem; doesnt pop up in front

$NewUsers = $null

while ($NewUsers -eq $null) 
{
    Write-Verbose -Message (' Pick a file, please ' ) -Verbose
    $NewUsers = Import-Csv -Path (File-Picker)
}


$RootDomain = (Get-ADDomain).distinguishedname
$UPNSUFFIX = (Get-ADDomain).forest



foreach ($User in $Newusers)
{
    $Description = $User.DESCRIPTION
    $FirstName = $User.FIRSTNAME
    $LastName = $User.LASTNAME
    $DisplayName = $FirstName + $LastName
    $SamAccountName = $User.FIRSTNAME[0] + $User.LASTNAME
    $OUnit = "OU=$Description,DC=Contoso,DC=Com"
    $Password = (ConvertTo-SecureString -AsPlainText 'P@ssw.rd1234' -Force)
    $OU = 'OU='+$Description +','+$RootDomain 
    $UPN = $SamAccountName +'@'+$UPNSUFFIX

    $UserY1 = Get-ADUser -Identity $SamAccountName
    $OldOU = $userY1.distinguishedname.split(',')[1]

    if (Test-ADOU -name $Description) 
    {
        Write-Verbose -Message ( $Description+ ' Exists' ) -Verbose
    }
    else 
    {
        New-ADOrganizationalUnit -name $Description -ProtectedFromAccidentalDeletion $false
    }   

    if (Test-ADUser -name $SamAccountName) 
    {
        Write-Verbose -Message ( $SamAccountName+ ' Exists' ) -Verbose

        if ($OldOU -eq 'OU='+ $description) {
            Write-Verbose -Message ( 'No Action Required' ) -Verbose
        }
        else {
            Get-ADUser -Identity $SamAccountName | Move-ADObject -TargetPath $OU
        }        
    }
    else 
    {
        New-ADUser -Name $SamAccountName -UserPrincipalName $UPN -SamAccountName $SamAccountName -Enabled $true -AccountPassword $Password -GivenName $Item.FIRSTNAME -Surname $Item.LASTNAME -Path $OU -DisplayName $DisplayName    
        Write-Verbose -Message ("creating user " + $SamAccountName) -Verbose
    }  

}
