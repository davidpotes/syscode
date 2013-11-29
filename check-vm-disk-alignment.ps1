# Get Alignment Script
# dp 9/15/10
# v 0.1
# tweak the $vms variable to adjust as necessary
# 

$cred = get-credential

$vms = get-vm -location prod

foreach ($vm in $vms)

{

$OffsetKB = @{label=”Offset(KB)”;Expression={$_.StartingOffset/1024 -as [decimal] }}
$ModOffset = @{Label = "Mod Offset"; Expression={(($_.startingoffset/1024) % 1024) -as [decimal] }}

$SizeMB = @{label=”Size(MB)”;Expression={$_.Size/1MB -as [int]}} 




Get-WmiObject -Computer $vm -credential $cred -Class "Win32_DiskPartition" | ft SystemName, Name, DiskIndex, $SizeMB, $OffsetKB, $ModOffset 

}
