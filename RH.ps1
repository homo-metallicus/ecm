#
# RH.ps1 (C) 2022 @homo-metallicus (Romain DECLE)
# https://github.com/homo-metallicus/moOde-mountstatus
#
# This Program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3, or (at your option)
# any later version.
#
# This Program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <http://www.gnu.org/licenses/>.
#

Set-ExecutionPolicy -force Unrestricted

$SOURCE="D:\SCAN\RH"
$DESTINATION="D:\SHARES\COMMUN\RESSOURCES HUMAINES"
$CORRESPONDANCE=[ordered]@{}
$CORRESPONDANCE["adm"]=[ordered]@{ "name" = "Administratif" ; "1" = "CNI" ; "2" = "PC" ; "3" = "Attestations" ; "4" = "RIB" ; "5" = "Etat civil" ; "6" = "Diplômes" ; "7" = "CV" ; "8" = "Casier judiciaire" ; "9" = "Notifications" }
$CORRESPONDANCE["ctr"]=[ordered]@{ "name" = "Contrat" ; "1" = "Contrats et avenants" ; "2" = "Arrêtés" ; "3" = "Fiches de notation" ; "4" = "Entretiens pro" ; "5" = "CET" ; "6" = "Rapports inspection" }
$CORRESPONDANCE["med"]=[ordered]@{ "name" = "Médical" ; "1" = "Visites aptitude" ; "2" = "Arrêts" ; "3" = "Visites poste de travail" }
$CORRESPONDANCE["form"]=[ordered]@{ "name" = "Formation" ; "1" = "Stages de formation" }
$CORRESPONDANCE["abs"]=[ordered]@{ "name" = "Absence" ; "1" = "Autorisations absence" ; "2" = "Congés enfant malade" ; "3" = "Réunions" ; "4" = "Absences non autorisées" }
$ANNEE=get-date -Format yyyy
$JOUR=get-date -Format yyyyMMdd
$i=0
$j=0
$ErrorActionPreference="SilentlyContinue"

 Function CreateFolder($path){
    $global:foldPath=$null
    foreach($foldername in $path.split("\"))
    {
        $global:foldPath+=($foldername+"\")
        if(!(Test-Path $global:foldPath))
        {
            New-Item -ItemType Directory -Path $global:foldPath
            echo ""
            echo "Création du dossier $global:foldPath"
        }
    }
}

Stop-Transcript | out-null

$ErrorActionPreference = "Continue"

Start-Transcript -path $SOURCE\Scan_RH_$JOUR.txt -Force

echo ""
echo "Copie des scans"

$SRC = "$SOURCE\*.pdf"
$filesPath = Get-ChildItem $SRC
ForEach($file in $filesPath){
    $fileName = $file.Name
    $shortedFileName = $fileName.Split('.')[0]; #remove the extension
    $newFileName = $fileName.Split('.')[0]+"_"+$JOUR+"."+$fileName.Split('.')[1]
    $arr = $shortedFileName.Split('-')

    #extract info
    $nom = $arr[0]
    $mainDirectory = $arr[1]
    $subDirectory = $arr[2]

    $directory = $CORRESPONDANCE[$mainDirectory]["name"]
    $subdirectory = $CORRESPONDANCE[$mainDirectory][$subDirectory]

    $DST = "$DESTINATION\$nom\$directory\$subdirectory"
    echo "Copie de $SOURCE\$fileName vers $DST\$newFileName"
    try {
        CreateFolder $DST
        Copy-Item -Force "$SOURCE\$fileName" "$DST\$newFileName"
        Remove-Item -Force $SOURCE\$fileName
        Write-Host 'OK' -fore green
        $i++
    } catch {
        Write-Host 'ERROR' -fore red
        $j++
    }
}

echo "Fichiers copiés : $i"

if ( $j -gt 0 ) {
    echo "Erreurs de copie : $j"
}

echo ""

Stop-Transcript

exit
