using namespace System
using namespace System.IO
using namespace System.Collections
using namespace System.Collections.Generic
############################################################################################################################
enum eDir {
  bin
  lib
  etc
  Result
}
###----------------------------------------------------------------------------------------------------------------------###
class Moddir
{
  [string]$bin
  [string]$lib
  [string]$etc
  [string]$result
  [List[string]]$penv = @( [eDir]::bin, [eDir]::lib, [eDir]::etc, [eDir]::Result)
  Moddir( [string]$lcBase )
  { ### 引数指定が無い場合はカレントディレクトリ設定
    if ( [string]::IsNullOrEmpty($lcBase) ) {
      $lcBase = [string]($(Get-Location)).ProviderPath
    }
    $this.bin    = "$($lcBase)\$($This.penv[0])"
    $this.lib    = "$($lcBase)\$($This.penv[1])"
    $this.etc    = "$($lcBase)\$($This.penv[2])"
    $this.result = "$($lcBase)\$($This.penv[3])"
  }
}
###----------------------------------------------------------------------------------------------------------------------###
class DefDataIni
{
  [string]$Prog
  [string]$inFile
  [string]$outFile
  [string]$errFile
  [string]$leafDir
  [string]$targetDir
  [string]$DataLimtTime
  [string]$DataLimtSize
  [string]$DataExtend
  DefDataIni([string]$gpre)
  {
    $idt = Get-Content "$gpre" | ConvertFrom-StringData
    $This.Prog         = "$($idt.TargetPorg)"
    $This.inFile       = "$($idt.TargetIfile)"
    $This.outFile      = "$($idt.TargetOfile)"
    $This.errFile      = "$($idt.TargetEfile)"
    $This.targetDir    = "\\$($idt.BaseDrive)\$($idt.BaseDir)\$($idt.Branches)\$($idt.LeavesDir)"
    $This.leafDir      = "\$($idt.LeavesDir)"
    $This.DataLimtTime = "$($idt.DefDate) $($idt.DefTime)"
    $This.DataLimtSize = "$($idt.DefSize)"
    $This.DataExtend   = "$($idt.DefExtens)"
  }
}
###----------------------------------------------------------------------------------------------------------------------###
class EnvCls
{
  [Moddir]$md
  [DefDataIni]$ddi
  EnvCls( [Moddir]$imd, [DefDataIni]$kddi )
  {
    $This.md  = $imd
    $This.ddi = $kddi
  }
}
############################################################################################################################
Function Get-ProgEnvironment([string]$reargs)
{
  Write-Host " Start  := $(Get-Date)"
  $iniFile = "DefineData.ini"
  $gpe     = [Moddir]::New( $reargs )
  $inidt   = [DefDataIni]::New("$($gpe.etc)\$iniFile")
  return [EnvCls]::New( $gpe, $inidt)
}
############################################################################################################################
Function ConvertFrom-Scaffolding( [object[]]$OutData )
{
  $OutDir, $StdInPut, $StdOutPut, $StdError = $OutData
  IF ( ! ( Test-Path -Path "$($OutDir)" ) )  {
    New-Item -ItemType "Directory" -Path "$($OutDir)"                 | Out-Null
    New-Item -ItemType "File"      -Path "$($OutDir)\$($StdInPut)"    | Out-Null
    New-Item -ItemType "File"      -Path "$($OutDir)\$($StdOutPut)"   | Out-Null
    New-Item -ItemType "File"      -Path "$($OutDir)\$($StdError)"    | Out-Null
  } ELSE {
    IF ( ! ( Test-Path -Path "$($OutDir)\$($StdInPut)" ) ) {
      New-Item -ItemType "File" -Path "$($OutDir)\$($StdInPut)"  | Out-Null
    }
    IF ( ! ( Test-Path -Path "$($OutDir)\$($StdOutPut)" ) ) {
      New-Item -ItemType "File" -Path "$($OutDir)\$($StdOutPut)" | Out-Null
    }
    IF ( ! ( Test-Path -Path "$($OutDir)\$($StdError)" ) ) {
      New-Item -ItemType "File" -Path "$($OutDir)\$($StdError)"  | Out-Null
    }
  }
}
############################################################################################################################
Function Expand-chkdata( [object]$pre )
{
  "$($pre.ddi.leafDir)"
  "$($pre.ddi.targetDir)"
  "$($pre.ddi.DataLimtTime)"
  "$($pre.ddi.DataLimtSize)"
  "$($pre.ddi.DataExtend)"
  ""
  "$($pre.ddi.Prog)"
  ""
  ConvertFrom-Scaffolding @( "$($pre.md.result)", "$($pre.ddi.inFile)", "$($pre.ddi.outFile)", "$($pre.ddi.errfile)" )
}
############################################################################################################################
Function Enter-Terminator
{
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  Write-Host " Finish := $(Get-Date)"
}
############################################################################################################################
Function Enter-NasCheck([object[]]$reargs) ### ラッパーモジュール
{
  #######　事前処理（定義データセット）
  $progEnv = Get-ProgEnvironment $reargs

  #######　設定データ確認
  Expand-chkdata $progEnv
  $processOptions = @{
      FilePath               = "$($progEnv.md.bin)\$($progEnv.ddi.Prog)"
      ArgumentList           = " dummy "
      RedirectStandardInput  = "$($progEnv.md.result)\$($progEnv.ddi.inFile)"
      RedirectStandardOutput = "$($progEnv.md.result)\$($progEnv.ddi.outFile)"
      RedirectStandardError  = "$($progEnv.md.result)\$($progEnv.ddi.errFile)"
  }
  #######　本ラッパーモジュールから本体起動（時間中断化ため２部構成）
  $psid = Start-Process @processOptions -PassThru  -WindowStyle Hidden

#  Watch-Timer $($WrpDef.Wrapers.TimeOut) $($obj) "$($processOptions.RedirectStandardError)"
  #######　終了処理（動的メモリ開放）
  Enter-Terminator
}
############################################################################################################################
#powershell -executionpolicy bypass -Command #${ Enter-Main $args }
#Start-Transcript  -OutputDirectory .
Set-PSBreakpoint NasCheck.PS1 126

Enter-NasCheck $args

Get-PSBreakpoint | Remove-PSBreakpoint
#Stop-Transcript
############################################################################################################################
<########
[AppDomain]::CurrentDomain.GetAssemblies() | % { $_.GetName().Name } | Sort-Object
########>
