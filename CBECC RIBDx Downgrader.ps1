Add-Type -AssemblyName System.Windows.Forms
function Select-File {
  param([string]$Directory = $PWD)

  $dialog = [System.Windows.Forms.OpenFileDialog]::new()

  $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
  $dialog.RestoreDirectory = $true
  $dialog.Filter = 'CBECCFile (*.ribd*x)|*.ribd*x'

  $result = $dialog.ShowDialog()

  if($result -eq [System.Windows.Forms.DialogResult]::OK){
    return $dialog.FileName
  }
}

$path = Select-File
$FolderPath = Split-Path -Path $path
$extensionType = (Get-ChildItem $path).Extension
Write-Output $FolderPath
Write-Output $path

Switch ($extensionType){
    ".ribd19x" {
        $nodestoremove = @(
            "//SoftwareVersionDetail","//StandardsVersion","//UseCommunitySolar","//CommunitySolarProject","//InCmntySlrProjTerritory",
            "//PVSizeOption","//PVWPwrElec", "//PVWSolarAccess", "//ReducedPVReq","//ReducedPVReqExcept", "//ReducedPVReqValue", "//SimStandaloneBatt",
            "//BattDRNumRankedDays", "//SFamCompactDistrib", "//SFamCDSpecFixDists", "//SFamUserSpecCmpctDist", "//SFamADUType", "//UseDefaultACH5",
            "//ProjInclKitche", "//TakeCntrlElecDHWSlrPVCred", "//SFamADUNumBedrooms", "//SFamADUArea", "//SFamMaxVertDist", "//IAQSupInletsAccessible", "//IAQHasFID",
            "//CSERpt_HtPumpHtg", "//StructPanel2Layer", "//StructPanelLayer", "//EffMetric", "//HSPF2", "//SEER2", "//EER2", "//VCHPDucts", "//VCHPCertAutoFan",
            "//IsLessThan25Ft", "//NewSmConsInstElecExcept","//CentralRecircType","//CHPWHSysDescrip","//CHPWHIntegPkgType","//CHPWHCompType","//CHPWHNEEABrand",
            "//CHPWHNEEAModel","//CHPWHComModel","//CHPWHTankLoc","//CHPWHSrcAirLoc","//CHPWHLoopTankType","//CHPWHLoopTankCompType",
            "//CHPWHLoopTankNEEABrand","//CHPWHLoopTankNEEAModel","//CHPWHLoopTankComModel","//CDServesMBathFix","//CDServesKitFix","//CDServesThirdFix",
            "//TotShowersServed", "//TotBathsServed", "//DWHRSysShowers", "//DWHRSysFeedHeater", "//HPWHModel", "//HPWHComModel", "//SFamDwellingType", "//EUseSummary",
            "//SFamDwellingType", "//UseDefaultACH50", "//ProjInclKitchen", "//BypassMessageBoxes", "//CompactDistrib", "//DWHRSysConfig", "//HPWHCategory",
            "//DWHRSysTakeCold"
        )

        if(Test-Path $path){
            Get-ChildItem $path | Foreach {Copy-item $_ "$FolderPath\$($_.Name -replace ".ribd19x", ".ribd16x")" -Force
                $NewFile = "$FolderPath\$($_.Name -replace ".ribd19x", ".ribd16x")"
            }

            $xml = [xml](Get-Content $NewFile)
            foreach($removeme in $nodestoremove){
                $xml.SelectNodes($removeme)
                $nodes = $xml.SelectNodes($removeme)
                foreach($node in $nodes){$node.ParentNode.RemoveChild($node)}
            }

            $OldNode = $xml.SelectNodes('//PVWCalFlexInstType')
            $index = 0
            foreach($node in $OldNode){
                $text = $node.InnerText
                $NewNode = $xml.CreateElement('PVWCalFlexInstall')
                $NewNode.SetAttribute("index", $index.ToString())
                $NewNode.InnerText = $text
                $node.ParentNode.AppendChild($NewNode)
                $node.ParentNode.RemoveChild($node)
                $index = $index + 1
            }

            $insulNode = $xml.SelectNodes('//InsulConsQuality')
            foreach($node in $insulNode){
                if ($node.InnerText = "Yes"){
                    $node.InnerText = "Improved"
                }else{
                    $node.InnerText = "Standard"
                }
            }
        
            $roofNodes = $xml.SelectNodes('//RoofingLayer')
            foreach($node in $roofnodes){
                if ($node.InnerText = "10 PSF (RoofTileAirGap)"){
                    $node.InnerText = "10 PSF (RoofTile)"
                }
            }

            $HPtext = $xml.SelectNodes('//HeaterElementType')
            foreach($node in $HPtext){
                if ($node.InnerText = "Heat Pump"){
                    $DeleteNode = $xml.SelectSingleNode('HPWHCategory')
                    $NewNode = $xml.CreateElement('HPWH_NEEARated')
                    $NewNode.InnerText = "1"
                    $node.ParentNode.AppendChild($NewNode)
                    foreach($delnode in $DeleteNode){$delnode.ParentNode.RemoveChild($delnode)}
                }
            }

            $HRV = $xml.SelectNodes('//IAQFan')
            foreach($node in $HRV){
                if ($node.IAQFanType = "Balanced"){
                    $NewNode = $xml.CreateElement('IAQRecovEffect')
                    $NewNode.InnerText = $node.SensRecovEff
                    $node.AppendChild($NewNode)
                    $DeleteNodes = @('//IncludesRecov', "//SensRecovEff", "//AdjSensRecovEff", "//SetPerfByHVIInterp")
                    foreach($removeme in $DeleteNodes){
                        $xml.SelectNodes($removeme)
                        $nodes = $xml.SelectNodes($removeme)
                        foreach($node in $nodes){$node.ParentNode.RemoveChild($node)}
                    }
                }
            }

            $nodes = $xml.SelectNodes("//RulesetFilename");
            foreach($node in $nodes) {
                $node.SetAttribute("file", "CA Res 2016.bin");
            }

            $xml.Save($NewFile)
            #[System.Windows.MessageBox]::Show('File Converted and saved here:' + $NewFile)
        }
    }
    
    ".ribd22x" {
        $nodestoremove = @(
            "//StandardsVersion", "//IsDualFuel", "//HasHPLockout"
        )
        
        if(Test-Path $path){
            Get-ChildItem $path | Foreach {Copy-item $_ "$FolderPath\$($_.Name -replace ".ribd22x", ".ribd19x")" -Force
                $NewFile = "$FolderPath\$($_.Name -replace ".ribd22x", ".ribd19x")"
            }
            
            $xml = [xml](Get-Content $NewFile)
            foreach($removeme in $nodestoremove){
                $xml.SelectNodes($removeme)
                $nodes = $xml.SelectNodes($removeme)
                foreach($node in $nodes){$node.ParentNode.RemoveChild($node)}
            }

            $nodes = $xml.SelectNodes("//RulesetFilename");
            foreach($node in $nodes) {
                $node.SetAttribute("file", "CA Res 2019.bin");
            }

            $xml.Save($NewFile)
            #[System.Windows.MessageBox]::Show('File Converted and saved here:' + $NewFile)
        }

    }

    default {
        #[System.Windows.MessageBox]::Show('Please choose the correct file type.')
    }
}
