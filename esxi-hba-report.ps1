$hbaReport = @()
foreach ($cluster in Get-cluster) {
  foreach ($vmHost in @($cluster | Get-vmhost)) {
    foreach ($hba in @($vmHost | Get-VMHostHba)) {
      $objHba = "" | Select ClusterName,HostName,Pci,Device,Type,Model,Status,Wwpn
      $objHba.ClusterName = $cluster.Name
      $objHba.HostName = $vmhost.Name
      $objHba.Pci = $hba.Pci
      $objHba.Device = $hba.Device
      $objHba.Type = $hba.Type
      $objHba.Model = $hba.Model
      $objHba.Status = $hba.Status
      $objHba.Wwpn = "{0:x}" -f $hba.PortWorldWideName
      $hbaReport += $objHba
    }
  }
}

$hbaReport | Export-Csv HbaReport.csv