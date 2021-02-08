# Script d'ajout d'utilisateurs avec un fichier .csv

$Users = Import-Csv -Delimiter "," -Path "C:\Utilisateurs.csv"            
foreach ($User in $Users)            
{            
    $Displayname = $User.Prenom + " " + $User.Nom            
    $UserFirstname = $User.Prenom            
    $UserLastname = $User.Nom
    $Prenom =  $UserFirstname.ToLower()
    $Nom =  $UserLastName.ToLower()   
    $OU = "OU=" + $User.Departement + "," + "dc=Caslex,dc=local"      
    $SAM = $Prenom+$Nom            
    $UPN = $Prenom + "." + $Nom + "@" + "Caslex.local"
    echo $UPN         
    $Departement = $User.Departement     
    $Password = $User.MotDePasse       
    New-ADOrganizationalUnit -Name $User.Departement -Path "DC=POTELAND,DC=local" -ProtectedFromAccidentalDeletion $false
    New-ADUser -Name "$Displayname" -DisplayName "$Displayname" -SamAccountName $SAM -UserPrincipalName $UPN -EmailAddress $UPN -GivenName $UserFirstname -Surname $UserLastname -Description $Description -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true -Path "$OU" -ChangePasswordAtLogon $false –PasswordNeverExpires $true     
# La ligne 21 sert à ajouter un script d'ouverture de session. Ajouter un # au début de la ligne afin d'annuler ses effets sur le script.
    Get-ADUser -Filter * -SearchBase $OU | Set-ADUser –scriptPath “\\NOM-DU-SERVEUR\$Departement.bat”
    New-ADGroup -Name $User.Departement -Description $User.Departement -DisplayName $User.Departement -Path $OU -SamAccountName $User.Departement -GroupScope Global
}
# FIN

