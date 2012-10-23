## ================================================
## ==                                            ==
## ==               script version 1.00          ==
## ==                                            ==
## ==                                 Moroz Oleg ==
## ==                                 10/02/2012 ==
## ==                             mod 03/10/2012 ==
## ================================================

$cDir = Get-Location
$wrkdir = $cDir.Path + "\"
$agentpath = "\\contoso.com\NETLOGON\msi\ccmsetup.exe"
$hostlist = $wrkdir + "hosts.csv"
$mplist = $wrkdir + "MPS.csv"
$hsts = Import-Csv $hostlist -Delimiter ";"
$mpcsv = Import-Csv $mplist -Delimiter ";"
Set-Alias psexec $wrkdir"PsExec.exe"
$psoutfile = $wrkdir + "psout.txt"
$scoutfile = $wrkdir + "scout.txt"
$errfile = $wrkdir + "faildlist.csv"
$faildcnt = 0
# comment next string for MP autodetection
#$mppredef="sccm-mp.contoso.com"

if (Test-Path $errfile) { Remove-Item $errfile }
$errtbl = @()

# create MPs hash
$MPS = @{}
foreach ($mp in $mpcsv) {
    $MPS.Add($mp.RD.ToLower(), $mp.MPoint.ToLower())
}

function ConvertTo-Object($hashtable) 
{
   $object = New-Object PSObject
   $hashtable.GetEnumerator() | 
      ForEach-Object { Add-Member -inputObject $object -memberType NoteProperty -name $_.Name -value $_.Value }
   $object
}

function NameFrom-CN($hcn)
{
    $cnparts = $hcn.split(",")
    $cnname = $cnparts[0]
    $hnameparts = $cnname.split("=")
    $hstname = $hnameparts[1]
    $hstname
}

foreach ($l in $hsts) {
    if ($l.DN -ne $null) { $hname = NameFrom-CN $l.DN }
    if ($l.Name -ne $null) { $hname = $l.Name }
    $flProblemClient = 0
    $prNoWait = ""
    Write-Output "Installing SCCM Agent for $hname"
	Write-Output "Check network path..."
	$hostadmdir = "\\" + $hname + "\admin$\"
	if (Test-Path $hostadmdir) {
		$hostdir64 = "\\" + $hname + "\admin$\SysWOW64\"
		if (Test-Path $hostdir64) { $hostdir = "\\" + $hname + "\admin$\ccmsetup" } else { $hostdir = "\\" + $hname + "\admin$\system32\ccmsetup" }
    	Write-Output "Check directory $hostdir exists"
    	if (!(Test-Path $hostdir)) {
        	Write-Output "Create directory $hostdir"
        	New-Item -path $hostdir -itemType "directory"
    	} else {
			Write-Output "Clear directory $hostdir"
			Write-Output "Check if ccmsetup service exists"
			& sc.exe \\$hname query ccmsetup 2>&1 | Out-File $scoutfile
			$scoutput = $(Get-Content -Path $scoutfile)
			if (($scoutput -match "FAILED") -or ($scoutput -match "ошибка")) {
				Write-Output "No service running. Safe to clean directory..."
			} else {
				Write-Output "Service ccmsetup found!!! Installation must be canceled..."
				Write-Output "Stopping service..."
				& sc.exe \\$hname stop ccmsetup
				Start-Sleep -Seconds 10
				Write-Output "Delete service..."
				& sc.exe \\$hname delete ccmsetup
				Start-Sleep -Seconds 5
				Write-Output "Safe to clean directory for now!"
			}
			Write-Output "Clear all files from setup directory"
			$delpath = $hostdir + "\*"
			Remove-Item $delpath -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable rmerr
            if ($rmerr.count -gt 0 ) {
                $rmerrcnt = $rmerr.count
                $flProblemClient = 1
                Write-Host -ForegroundColor "Yellow" "Cleanup complete with $rmerrcnt error(s). But try to continue..."
            }
		}
    	Write-Output "Copy setup file to local directory"
        try {
    	   Copy-Item -Force $agentpath $hostdir -ErrorAction SilentlyContinue
        } catch {
            $flProblemClient = 1
            Write-Host -ForegroundColor "Yellow" "Copy setup files error! But try to continue..."
        }
		Write-Output "Determine correct local path..."
		$pXP = "\\" + $hname + "\c$\windows\"
		$p2K = "\\" + $hname + "\c$\winnt\"
		$p64 = "\\" + $hname + "\c$\windows\syswow64"
		$remotecmd = ""
		if (Test-Path $pXP) {
			$remotecmd = "c:\windows\system32\ccmsetup\ccmsetup.exe"
			Write-Output "Using local path $remotecmd"
		}
		if (Test-Path $p2K) {
			$remotecmd = "c:\winnt\system32\ccmsetup\ccmsetup.exe"
			Write-Output "Using local path $remotecmd"
		}
		if (Test-Path $p64) {
			$remotecmd = "c:\windows\ccmsetup\ccmsetup.exe"
			Write-Output "Using local path $remotecmd"
		}
		if ($remotecmd -eq "") {
			$faildcnt++
			$err = @{Name=$hname; Error="no valid local path"}
			$errtbl += ConvertTo-Object $err
			Write-Output "No valid path found!"
			Write-Output "All done for $hname"
			continue
		}
        Write-Output "Determine management point..."
        $hostmp = ""
        if ($l.RD -ne $null) { $hostmp = $MPS.Get_Item($l.RD) }
        if ($hostmp -ne "") {
            Write-Output "MP: $hostmp"
            $remcmdparams = "/retry:1 /mp:" + $hostmp
        } else {
            Write-Output "MP not found! Try to autodetect MP during setup..."
            $remcmdparams = "/retry:1"
        }
        if ($mppredef -ne $null) {
            Write-Output "MP predefined! Override detected MP with $mppredef"
            $hostmp=$mppredef
        }
        if ($flProblemClient -gt 0) {
            Write-Host -ForegroundColor "Yellow" "Potential problem client. Run setup with no wait parameter"
            $prNoWait = "-d"
        }
        Write-Output "Executing client installation..."
    	& psexec -e $prNoWait -n 90 \\$hname $remotecmd $remcmdparams 2>&1 | Out-File $psoutfile
    	$psoutput = $(Get-Content -Path $psoutfile)
    	if ($psoutput -match "starting") {
        	Write-Output "Remote command started..."
        	if ($psoutput -match "0\.") {
            	Write-Output "Program successed!"
        	} else {
                if ($flProblemClient -gt 0) {
                   Write-Output "Status unknown!"
            	   $faildcnt++
				   $err = @{Name=$hname; Error="status unknown"}
				   $errtbl += ConvertTo-Object $err
                } else {
            	   Write-Output "Program faild!"
            	   $faildcnt++
				   $err = @{Name=$hname; Error="install faild"}
				   $errtbl += ConvertTo-Object $err
                }
        	}
    	} else {
        	Write-Output "Remote command start faild!"
        	$faildcnt++
			$err = @{Name=$hname; Error="remote command start faild"}
			$errtbl += ConvertTo-Object $err
    	}
	} else {
		Write-Output "No connection to host!"
		$faildcnt++
		$err = @{Name=$hname; Error="no connection to host"}
		$errtbl += ConvertTo-Object $err
	}
    if (Test-Path $psoutfile) { Remove-Item $psoutfile }
	if (Test-Path $scoutfile) { Remove-Item $scoutfile }
    Write-Output "All done for $hname"
    Write-Output " "
}

if (!($faildcnt -eq 0)) {
	$errtbl | Export-Csv $errfile -NoTypeInformation
    Write-Output "$faildcnt error(s)! See $errfile for details!"
}

Write-Output "All Done!"
Write-Output "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")