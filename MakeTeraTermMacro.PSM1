using namespace System
using namespace System.IO
using namespace System.Text
using module    ".\ListLoginSheet.PSM1"
####################################################################################
function CommonCmdDetail ( [StreamWriter]$fp, [object[]] $EncDat )
{
  $number,  $SendCmd,  $WaitCmd = $EncDat
  [StreamWriter]$fp.WriteLine("sendln '$($SendCmd)'")
  [StreamWriter]$fp.WriteLine("wait $($WaitCmd)")
}
####################################################################################
function TeratermEncDataDetail ( [StreamWriter]$fp, [object[]] $EncDat )
{
  $number,  $User,  $encpass = $EncDat
  [StreamWriter]$fp.WriteLine("$($User)=$($encpass)")
}
####################################################################################
function CommonCmd ( [string]$CCMD, [object[,]] $EncDat )
{
  $RStart = 1
  $LyrCtgy = $CCMD
  $CommonDir = "$([DefFiles]::Common)"
  $uFile = "UpDate01"
  $CmdfileName = "$($PSScriptRoot)\$($CommonDir)\$($LyrCtgy)$($uFile).TTL"
  if ( !( Test-Path -Path $CmdfileName ) )
  {
    $fl = New-Object StreamWriter($CmdfileName, $false, [Encoding]::GetEncoding("sjis"))
    $RStart..$EncDat.GetLength(0) | %{ $Row = $PSItem 
      if ( $Row -ne 1 ){
        CommonCmdDetail $fl $([string]$EncDat[$Row, 1]) , $([string]$EncDat[$Row, 2]) ,$([string]$EncDat[$Row, 3])
      }
    }
    $fl.Close()
  }
  return $CmdfileName
}
####################################################################################
function TeratermEncData ( [string]$CCMD, [object[,]] $EncDat )
{
  $RStart = 1
  $PASDTDir  = $([DefFiles]::PASDAT)
  $PasfileName = "$($PSScriptRoot)\$($PASDTDir)\PASSWORD.DAT"

  if ( !( Test-Path -Path $PasfileName ) )
  {
    $fl = New-Object StreamWriter($PasfileName, $false, [Encoding]::GetEncoding("sjis"))
    $RStart..$EncDat.GetLength(0) | %{ $Row = $PSItem
      if ( $Row -eq 1 ){
        $fl.WriteLine("[Password]")
      }
      if ( $Row -gt 1 ){
        TeratermEncDataDetail $fl $([string]$EncDat[$Row, 1]) , $([string]$EncDat[$Row, 2]) ,$([string]$EncDat[$Row, 3])
      }
    }
    $fl.Close()
  }
  return $PasfileName
}
####################################################################################
function MakeTeraTermMacro ( [string]$passFile, [string[]]$IncFileName, [string[]]$params )
{
  $categy, $hostnm, $ipa, $user = $params

  $MacroDir  = $([DefFiles]::TetaTermMacro)
  $CommonDir = "..\$([DefFiles]::Common)"
  $fileName  = "$($PSScriptRoot)\$($MacroDir)\$($categy)_$($hostnm).TTL"
  if ( !( Test-Path -Path $fileName ) )
  {
    $fl = New-Object StreamWriter($fileName, $false, [Encoding]::GetEncoding("sjis"))
    $fl.WriteLine("USERNAME = '$user'")
    $fl.WriteLine("PASSFILE = '$passFile'")
    $fl.WriteLine("LOGSPATH = '$($PSScriptRoot)\$([DefFiles]::LOG)\'")
    $fl.WriteLine("HOSTADDR = '$ipa'")
    $fl.WriteLine("TITLE = '$hostnm'")
    $fl.WriteLine("getpassword PASSFILE USERNAME PASSWORD")
    $fl.WriteLine("COMMAND = HOSTADDR")
    $auterparam = switch($categy){"LB"{"challenge"}Default{"password"}}
    $fl.WriteLine("strconcat COMMAND ':22 /ssh /2 /auth=$($auterparam) /user=$user /passwd=' ")
    $fl.WriteLine("strconcat COMMAND PASSWORD")
    $fl.WriteLine("getdate LOG_NAME '$($categy)_$($hostnm)_&h_%Y%m%d_%H%M%S.log'")
    $fl.WriteLine("FULLPATH = LOGSPATH")
    $fl.WriteLine("strconcat FULLPATH LOG_NAME")
    $fl.WriteLine("connect COMMAND")
    $fl.WriteLine("if result <> 2 then ")
    $fl.WriteLine("  messagebox 'It could not be connected.' 'Connection Error' ")
    $fl.WriteLine("  end")
    $fl.WriteLine("endif")
    $fl.WriteLine("settitle TITLE")
    $fl.WriteLine("logopen FULLPATH")
    $fl.WriteLine("timeout = 300")
    $fl.WriteLine("wait '$' '#'")
    $fl.WriteLine("if result = 0 then")
    $fl.WriteLine("  end")
    $fl.WriteLine("endif")
    $fl.WriteLine("include '$($IncFileName.Where({$PSItem.Contains(($categy).SubString(0,2))}).ForEach({$PSItem}))'")
    $fl.WriteLine("end")
    $fl.Close()
  }
  return $fileName
}
####################################################################################
