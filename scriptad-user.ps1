# Script à executer en tant qu'administrateur

Import-Module ActiveDirectory

$date=Get-Date -Format "dd-MM-yyyy"

$a=0
$server="192.168.0.1"
$fichier=Read-Host "Nom du fichier csv ? (ex: etudiants.csv)"


# Cette fonction convertie la promotion du fichier CSV vers le groupe dans l'AD
#
# Exemple un étudiant en ESSCA/ISSE B2 sera convertie vers le groupe ISEE B2
Function PromoCSV2PromoAD{
    
    if($promoCSV.Contains("Ingesup") -or $promoCSV.Contains("INGESUP")){
        if($promoCSV.Contains("B1")){$promo="Ingesup B1"} ; if($promoCSV.Contains("B2")){$promo="Ingesup B2"} ; if($promoCSV.Contains("B3")){$promo="Ingesup B3"}
        if($promoCSV.Contains("M1") -or $promoCSV.Contains("MAST1") ){$promo="Ingesup MAST1"} ; if($promoCSV.Contains("M2") -or $promoCSV.Contains("MAST2") ){$promo="Ingesup MAST2"}
    }

    elseif($promoCSV.Contains("ISEE") -or $promoCSV.Contains("Isee") -or $promoCSV.Contains("ESSCA")){
        if($promoCSV.Contains("B1")){$promo="ISEE B1"} ; if($promoCSV.Contains("B2")){$promo="ISEE B2"} ; if($promoCSV.Contains("B3")){$promo="ISEE B3"}
        if($promoCSV.Contains("M1") -or $promoCSV.Contains("MAST1") ){$promo="ISEE MAST1"} ; if($promoCSV.Contains("M2") -or $promoCSV.Contains("MAST2") ){$promo="ISEE MAST2"}
    }

    elseif($promoCSV.Contains("LIMART") -or $promoCSV.Contains("Limart") -or $promoCSV.Contains("Lim'Art") -or $promoCSV.Contains("LIM'ART")){
        if($promoCSV.Contains("B1")){$promo="LIMART B1"} ; if($promoCSV.Contains("B2")){$promo="LIMART B2"} ; if($promoCSV.Contains("B3")){$promo="LIMART B3"}
    }



    return $promo
}


# Cette fonction récupère le groupe actuel de l'étudiant dans l'AD
Function Get-Group{

    if($($ingesupB1 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="Ingesup B1"}
    elseif($($ingesupB2 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="Ingesup B2"}
    elseif($($ingesupB3 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="Ingesup B3"}
    elseif($($ingesupMast1 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="Ingesup MAST1"}
    elseif($($ingesupMast2 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="Ingesup MAST2"}

    elseif($($ISEEB1 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="ISEE B1"}
    elseif($($ISEEB2 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="ISEE B2"}
    elseif($($ISEEB3 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="ISEE B3"}
    elseif($($ISEEMast1 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="ISEE MAST1"}
    elseif($($ISEEMast2 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="ISEE MAST2"}

    elseif($($LIMARTB1 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="LIMART B1"}
    elseif($($LIMARTB2| where {($_ | select * | out-string) -match "$samAccount"})){$groupe="LIMART B2"}
    elseif($($LIMARTB3 | where {($_ | select * | out-string) -match "$samAccount"})){$groupe="LIMART B3"}


    return $groupe
}



if(Test-Path $fichier){

    $csv=Import-Csv -Delimiter ";" -Path $fichier

    # Récupération des étudiants
    $ingesupB1=Get-ADGroupMember -Identity "Ingesup B1" ; $ingesupB2=Get-ADGroupMember -Identity "Ingesup B2" ; $ingesupB3=Get-ADGroupMember -Identity "Ingesup B3"
    $ingesupMast1=Get-ADGroupMember -Identity "Ingesup MAST1" ;$ingesupMast2=Get-ADGroupMember -Identity "Ingesup MAST2"

    $ISEEB1=Get-ADGroupMember -Identity "ISEE B1" ; $ISEEB2=Get-ADGroupMember -Identity "ISEE B2" ; $ISEEB3=Get-ADGroupMember -Identity "ISEE B3"
    $ISEEMast1=Get-ADGroupMember -Identity "ISEE MAST1" ; $ISEEMast2=Get-ADGroupMember -Identity "ISEE MAST2"

    $LIMARTB1=Get-ADGroupMember -Identity "LIMART B1" ; $LIMARTB2=Get-ADGroupMember -Identity "LIMART B2" ; $LIMARTB3=Get-ADGroupMember -Identity "LIMART B3"
    # Fin de la récupération


    foreach($line in $csv){
        Remove-Variable test,prenom2 -ErrorAction SilentlyContinue

        $mail=$line.'Adresse e-mail'
        $nom=$line.Nom
        $prenom=$line.Prenom
        $promoCSV=$line.Promotions

        $promo=PromoCSV2PromoAD

        Write-Host " mail : $mail , nom : $nom , prenom : $prenom , promo : $promoCSV"

        if($nom.Contains(" ")){ $nom=$nom -replace '\s',''}   
        
        if($nom.Contains(" ")){ $nom=$nom.Replace(" ","")}
        if($nom.Contains("'")){ $nom=$nom.Replace("'","") }
        if($nom.Contains("É")){ $nom=$nom.Replace("É","E") }

        if($prenom.Contains(" ")){ $prenom2=$prenom.Replace(" ","")}
        if($prenom.Contains("'")){ $prenom2=$prenom.Replace("'","") }
        if($prenom.Contains("é")){ $prenom2=$prenom.Replace("é","e") }
        if($prenom.Contains("è")){ $prenom2=$prenom.Replace("è","e") }
        if($prenom.Contains("É")){ $prenom2=$prenom.Replace("É","E") }
        if($prenom.Contains("ë")){ $prenom2=$prenom.Replace("ë","e") }

        if($prenom.Contains("à")){ $prenom2=$prenom.Replace("à","a") }
        if($prenom.Contains("â")){ $prenom2=$prenom.Replace("â","a") }

        if($prenom.Contains("ï")){ $prenom2=$prenom.Replace("ï","i") }
        if($prenom.Contains("î")){ $prenom2=$prenom.Replace("î","i") }

        if($prenom.Contains("-")){ $prenom2=$prenom.Replace("-","") }
        if($prenom.Contains("'")){ $prenom2=$prenom.Replace("'","") }


        if(!$mail){
            if(!$prenom2){
                $prenom2=$prenom
            }

            $mail="$prenom2"+"."+"$nom"+"@ynov.com"
            $mail=$mail.Replace(" ","")

            $mail=$mail.ToLower()
        }


        $samAccount=$($($prenom[0])+"$nom").ToLower()


        if($promoCSV.contains("Lyon")){
            

            # L'etudiant n'est pas nouveau dans l'école
            if($(Get-ADUser "$samAccount")){
                $a=$a+1
                Write-Host -ForegroundColor Cyan "Utilisateur présent dans l'AD : $nom | $prenom | $samAccount | $mail | $promo"
                <#
                # Récupération du groupe actuel de l'étudiant
                $AncienGroupe=Get-Group

                if($AncienGroupe -ne $promo){
                    Write-Host "Changement de classe : $AncienGroupe -> $promo"
                    
                    # Suppression de l'étudiant du groupe
                    Add-Content -Value "L'utilisateur $samAccount a été supprimé du groupe : $AncienGroupe" -Path C:\scripts\logs\ajout-etudiants-$date.txt
                    Remove-ADGroupMember -Identity "$AncienGroupe" -Member "$samAccount" -Confirm:$false
                    
                    # Ajout de l'étudiant dans le nouveau groupe
                    Add-Content -Value "L'utilisateur $samAccount a été ajouté dans le groupe : $promo" -Path C:\scripts\logs\ajout-etudiants-$date.txt
                    Add-ADGroupMember -Server $server "$promo" -Members "$samAccount"
                    Add-Content -Value " " -Path C:\scripts\logs\ajout-etudiants-$date.txt

                }
                else{
                    Write-Host -ForegroundColor DarkMagenta "Redoublement pour : $prenom | $nom  "
                    Add-Content -Value "Redoublement pour : $prenom | $nom  " -Path C:\scripts\logs\ajout-etudiants-$date.txt
                }

                #>
            }
            else{
                
                Write-Host -ForegroundColor Red "L'utilisateur n'est pas présent dans l'AD $prenom | $nom | $samAccount | $mail | $promo"
                
                #New-ADUser -Server $server -SamAccountName "$samAccount" -EmailAddress "$mail" -Name "$nom $prenom" -GivenName "$prenom" -Surname "$nom" -UserPrincipalName "$samAccount" -Path "OU=Etudiants,OU=Campus_LYON,DC=ynovlyon,DC=fr" -PasswordNotRequired 1 -Enabled $true -ChangePasswordAtLogon 0
                #Add-ADGroupMember -Server $server "$promo" -Members "$samAccount"

                #Add-Content -Value "L'utilisateur $prenom | $nom | $mail | $promo a été créé dans l'Active Directory" -Path C:\scripts\logs\ajout-etudiants-$date.txt
            }

        }
    }





    <# Déplacement et désactivation des anciens étudiants dans l'OU A_Supprimer
    $fichier=Import-Csv .\etudiants.csv -Delimiter ";"
    $i=0

    $users=Get-ADUser -Filter * -Properties * -SearchBase "OU=Etudiants,OU=Campus_LYON,DC=ynovlyon,DC=fr"


    foreach($user in $users){
        $i=$i+1
        $prenom=$user.GivenName
        $nom=$user.Surname


        if(!$($fichier | ?{$_.Nom -eq "$nom" -and $_.Prenom -eq "$prenom"})){
            Write-Host -ForegroundColor Red "$prenom $nom KO"
            #Disable-ADAccount $user
            #Move-ADObject $user -TargetPath "OU=A_Supprimer,OU=Campus_LYON,DC=ynovlyon,DC=fr"
            Add-Content -Value "Déplacement de $prenom $nom" -Path
        }




        Clear-Variable prenom,nom,test

    }

    Write-Host "`n$i Utilisateurs"

    #>

}
else{
    Write-Host -ForegroundColor Yellow "Le fichier $fichier n'a pas été trouvé"
}