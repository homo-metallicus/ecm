#
# LGA.ps1 (C) 2022 @homo-metallicus (Romain DECLE)
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

$SOURCE="D:\SCAN\lycee-admin\LGA-Compta"
$DESTINATION="D:\SHARES\EPL\COMMUN EPL\7-GESTION FINANCIERE et PAIES\5-COCWINELLE"
$CENTRES=[ordered]@{ "00" = "EPL" ; "01" = "LYCEE" ; "02" = "EXPLOITATION AGRICOLE" ; "03" = "CFA" ; "04" = "CFPPA" ; "05" = "EXPLOITATION HORTICOLE" }
$ANNEE=get-date -Format yyyy
$POSTES=[ordered]@{ "CA" = "CERTIFICAT ADMINISTRATIF" ; "D" = "DEPENSES" ; "OR" = "ORD REDUC RECETTES" ; "OV" = "ORD VERS DEPENSES" ; "P" = "PENSIONS" ; "R" = "RECETTES" ; "RD" = "REIMPUTATION DEPENSES" ; "RJ" = "REJETS" ; "RR" = "REIMPUTATION RECETTES" ; "S" = "SALAIRES" }
$OPERATIONS=[ordered]@{ "D" = "DEPENSES" ; "OR" = "ORD REDUC RECETTES" ; "OV" = "ORD VERS DEPENSES" ; "R" = "RECETTES" ; "RJ" = "REJETS" }
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

#Function MoveFiles($src,$dst){
#    echo $src." : ".$dst."
#    #Get-ChildItem $src | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $dst" ; try { Copy-Item -Force $SOURCE\$PDF $dst ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
#}

Stop-Transcript | out-null

$ErrorActionPreference = "Continue"

Start-Transcript -path $SOURCE\Scan_LGA_$JOUR.txt -Force

echo ""
echo "Copie des scans"

foreach($c in $CENTRES.Keys)
{
    $CENTRE = $($CENTRES.item($c))
    foreach($p in $POSTES.Keys)
    {
        $POSTE = $($POSTES.item($p))
        if ($POSTE -like "PENSIONS")
        {
            foreach($o in $OPERATIONS.Keys)
            {
                $OPERATION = $($OPERATIONS.item($o))
                $SRC = "$SOURCE\$c-$p-$o-*.pdf"
                $DST = "$DESTINATION\ANNEE $ANNEE\$c-$CENTRE\$POSTE\$OPERATION\"
                CreateFolder $DST
                Get-ChildItem $SRC | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $DST" ; try { Copy-Item -Force $SOURCE\$PDF $DST ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
                #MoveFiles($SRC,$DST)
            }
        }
        else
        {
            $SRC = "$SOURCE\$c-$p-*.pdf"
            $DST = "$DESTINATION\ANNEE $ANNEE\$c-$CENTRE\$POSTE\"
            CreateFolder $DST
            Get-ChildItem $SRC | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $DST" ; try { Copy-Item -Force $SOURCE\$PDF $DST ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
            #MoveFiles($SRC,$DST)
        }
        #Get-ChildItem $SRC | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $DST" ; try { Copy-Item -Force $SOURCE\$PDF $DST ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
    }
}

echo ""
echo "Copie des contrats"

foreach($c in $CENTRES.Keys)
{
    $CENTRE = $($CENTRES.item($c))
    $SRC_CT = "$SOURCE\CT-$c-*.pdf"
    $DST_CT = "$DESTINATION\CONTRATS\$c-$CENTRE\"
    CreateFolder $DST_CT
    Get-ChildItem $SRC_CT | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $DST_CT" ; try { Copy-Item -Force $SOURCE\$PDF $DST_CT ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
    #MoveFiles($SRC_CT,$DST_CT)
}

echo ""
echo "Copie des RIB"

$SRC_RIB = "$SOURCE\RIB-*.pdf"
$DST_RIB = "$DESTINATION\RIB\"
CreateFolder $DST_RIB
Get-ChildItem $SRC_RIB | % -Process { $PDF = $_.Name ; echo "Copie de $PDF vers $DST_RIB" ; try { Copy-Item -Force $SOURCE\$PDF $DST_RIB ; Remove-Item -Force $SOURCE\$PDF ; Write-Host 'OK' -fore green ; $i++ } catch { Write-Host 'ERROR' -fore red ; $j++ } ; echo "" }
#MoveFiles($SRC_RIB,$DST_RIB)

echo "Fichiers copiés : $i"

if ( $j -gt 0 )
{
    echo "Erreurs de copie : $j"
}

echo ""

Stop-Transcript

exit