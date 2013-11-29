param ( $Age = 30)
connect-viserver amon4a.dc.gtnexus.com
$vm = get-vm
$snapshots = get-snapshot -vm $vm
write-host -foregroundcolor red Snapshots found:"
foreach ( $snap in $snapshots )
	write-host "Name: " $snap.name " Size: " $snap.sizeMB "
	
	