using namespace System.Collections.Generic
$exBs = New-Object -ComObject Excel.Application
$defBook = "\\vdinas03\FolderRedirect\NSS\fbs_matsumoto_kazuhi\home\fbsmatsu\develop\AutoLogin\DAT\ServerLoginList.xlsx" 
$RefBook = $exBs.Workbooks.Open( $defBook )
$TargetTable = $RefBook.Sheets( "DefineLoginSheet" ).ListObjects( "DefineLogin")
$V2     = $TargetTable.Range.Value2
$HeadRw = $TargetTable.HeaderRowRange.Row
$HeadCnt= $TargetTable.HeaderRowRange.Count
$Header = $TargetTable.HeaderRowRange
$ColNum = $TargetTable.Listcolumns.Count
$RowNum = $TargetTable.Listrows.Count
$DatBdy = $TargetTable.DataBodyRange
$DatBdy2= $TargetTable.DataBodyRange.Value2
# 1..$HeadCnt | %{   Write-Output("$($Header[$PSItem].Value2)") }
enum TTMS
{
  No        = 1
  Category  = 2
  hostname  = 3
  IPAddress = 4
  LoginUser = 5
}
class DefineLoginSheet
{
  [string]$No
  [string]$Category
  [string]$hostname
  [string]$IPAddress
  [string]$LoginUser
  DefineLoginSheet( )  {}
  DefineLoginSheet( [string]$No, [string]$Category, [string]$hostname, [string]$IPAddress, [string]$LoginUser )
  {
    $this.No        = $No
    $this.Category  = $Category
    $this.hostname  = $hostname
    $this.IPAddress = $IPAddress
    $this.LoginUser = $LoginUser
  }
  [string] ToString2() {
    return "No:=$($this.No),Category:=$($this.Category),hostname:=$($this.hostname),IPAddress:=$($this.IPAddress),LoginUser:=$($this.LoginUser)"
  }
}
class LoginSheet : List[DefineLoginSheet]
{
  [int]$ColNum
  [int]$RowNum
  LoginSheet ( )  {  }
  LoginSheet ( [int]$ColNum, [int]$RowNum, [object]$LstTbl3 )
  {
    $this.ColNum = $ColNum
    $this.RowNum = $RowNum
    1..$this.RowNum | %{
        $this.Add( [DefineLoginSheet]::new( "$($LstTbl3.Get( $_ ,[TTMS]::No))",
                                            "$($LstTbl3.Get( $_ ,[TTMS]::Category))",
                                            "$($LstTbl3.Get( $_ ,[TTMS]::hostname))",
                                            "$($LstTbl3.Get( $_ ,[TTMS]::IPAddress))",
                                            "$($LstTbl3.Get( $_ ,[TTMS]::LoginUser))")
        )
    }
  }
  [string]ToString()
  {
    $ret = 0..($this.RowNum - 1) | %{
      Write-Output ( "$($this[$_].ToString2())" )
    }
    return "$($ret)"
  }
}

$DLS = [LoginSheet]::new($ColNum, $RowNum, $DatBdy2)
$DLS | Format-Table
$DLS.ToString()

$RefBook.Close()
$exBs.Quit()
gv | ? Value -is [__ComObject] | clv
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
