#Requires -Version 3.0
<#
    .SYNOPSIS
    Creates a complete listing of Ports, Protocols and Services for Microsoft Windows systems.
    .DESCRIPTION
    This script provides IT/IA with a tabular output of currently available and/or active Ports, Protocols and Services.
    The output can then be copied and pasted into any document format.
    With the resules window open, press Ctrl A, then Ctrl C and then paste into any document format.
    .PARAMETER
    No parameters are required.
    .INPUTS
    None.
    .OUTPUTS
    No objects are output from this script.  This script generates and displays tables of Ports, Protocols and Services.
    .NOTES
    NAME: PPS.ps1
    VERSION: 1.0.1.7
    AUTHOR: Lee Lepore
    LASTEDIT: October 17, 2019
#>

#  Pop-up Info Message
$wshell = New-Object -ComObject Wscript.Shell
$wshell.popup('Gathering PPS Details, Press OK to Continue',0,'Ports, Protocols and Services',64)

#   Retrieve Firewall Rules
Try
{
  If(!$Enabled)
  {
    $Firewall = (New-Object -ComObject hnetcfg.fwpolicy2).rules
  }
  Else
  {
    $Firewall = (New-Object -ComObject hnetcfg.fwpolicy2).rules | Where-Object -FilterScript {
      $_.enabled -like $Enabled 
    }
  }
               
  If (!$Firewall) 
  {
    Throw 'Failed to pull Firewall Rules'
  }
}
Catch 
{
  Write-Verbose -Message $_.Exception.Message
  Break
}

#   Building array with Firewall Rule info
$FWArray = @()

#   Looping for each Rule
$Firewall | ForEach-Object -Process {
  $Rule = $_
  Switch($Rule.Direction) 
  { 
    '1' 
    {
      $Direction = 'Inbound' 
    } 
    '2' 
    {
      $Direction = 'Outbound' 
    }  
  }
  Switch($Rule.Action)
  {
    '0' 
    {
      $Action = 'Block' 
    } 
    '1' 
    {
      $Action = 'Allow' 
    }  
  }
  Switch($Rule.Profiles)
  {
    '1' 
    {
      $Profile = 'Domain' 
    } 
    '2' 
    {
      $Profile = 'Private' 
    }  
    '4' 
    {
      $Profile = 'Public' 
    } 
    '2147483647' 
    {
      $Profile = 'All' 
    }  
  }
  Switch($Rule.Protocol)
  {
    '6' 
    {
      $Protocol = 'TCP' 
    } 
    '17' 
    {
      $Protocol = 'UDP' 
    }  
    '1' 
    {
      $Protocol = 'ICMPv4' 
    } 
    '58' 
    {
      $Protocol = 'ICMPv6' 
    }  
  }

  #   Create a custom object containing all gathered values
  $Object = New-Object -TypeName PSCustomObject
  $Object | Add-Member -MemberType NoteProperty -Name 'Direction' -Value $Direction
  $Object | Add-Member -MemberType NoteProperty -Name 'Action' -Value $Action
  $Object | Add-Member -MemberType NoteProperty -Name 'Rule Name' -Value $Rule.name
  $Object | Add-Member -MemberType NoteProperty -Name 'Profile' -Value $Profile
  $Object | Add-Member -MemberType NoteProperty -Name 'Enabled' -Value $Rule.Enabled
  $Object | Add-Member -MemberType NoteProperty -Name 'Protocol' -Value $Protocol
  $Object | Add-Member -MemberType NoteProperty -Name 'Local Ports' -Value $Rule.Localports

  #   Add custom object data to the array
  If ($Object.Enabled -eq 'True')  
  {
    $FWArray += $Object
  }
}
If(!$RuleType)
{
  #   Display Ports and Protocols
  $FWArray |
  Sort-Object -Property enabled -Descending |
  Out-GridView -Title ('Ports and Protocols') -PassThru > $null
}
Else
{
  #   Display Ports and Protocols
  $FWArray |
  Where-Object -FilterScript {
    $_.Direction -eq $RuleType
  } |
  Sort-Object -Property direction, enabled |
  Out-GridView -Title ('Ports and Protocols') -PassThru > $null
}

#   Display Services
Get-Service |
Where-Object -FilterScript {
  ($_.StartType -eq 'Automatic') -or ($_.Status -eq 'Running')
} |
Select-Object -Property Name, DisplayName, Status, StartType |
Out-GridView -Title ('Services') -PassThru > $null