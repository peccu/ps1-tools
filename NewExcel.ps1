function new-excel {
  (New-Object -ComObject Excel.Application).Visible = $true
  Read-Host -Prompt "Press Enter to exit (also kills Excel)"
}
new-excel
