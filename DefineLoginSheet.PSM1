using namespace System.Collections.Generic
####################################################################################
class DefineLoginSheet
{
  [string]$No
  [string]$Category
  [string]$hostname
  [string]$IPAddress
  [string]$LoginUser
  DefineLoginSheet( ) {  }
  DefineLoginSheet( [string]$No, [string]$Category, [string]$hostname, [string]$IPAddress, [string]$LoginUser )
  {
    $this.No        = $No
    $this.Category  = $Category
    $this.hostname  = $hostname
    $this.IPAddress = $IPAddress
    $this.LoginUser = $LoginUser
  }
}
####################################################################################
