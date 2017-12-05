[void] [System.Reflection.Assembly]::LoadWithPartialName("'System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")

Function DoTesting {
	param(
		$loginUrl,
		$url, 
		$count,
		[switch]$file, 
		[switch]$twoTab
	)
	
	$global:ieproc = (Get-Process -Name iexplore)|? {$_.MainWindowHandle -eq $global:ie.HWND}
	if ($loginUrl -ne ""){
		$global:ie.Navigate($loginUrl)
		IEWaitForPage 500
		$rrr = Read-Host "Please login to the site and press enter"
	}
	if ($file){
		Write-Host "Downloading $($url):"
	}
	else{
		Write-Host "Navigating to $($url):"
	}
	for ($i=0; $i -lt $count; $i++) {
		IENavigateTo $url -file:$file -twoTab:$twoTab
	}
}

Function IEOpen(){
	$global:ie = New-Object -Com "InternetExplorer.Application"
	$global:ie.Visible = $true
}
Function IEClose(){
	if ($global:ie) {
		Write-Host "Closing IE"
		$global:ie.Quit()
	}
	$global:ieproc | Stop-Process -Force -ErrorAction SilentlyContinue
	
	Remove-item "$env:systemroot\system32\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
	Remove-item "$env:systemroot\syswow64\config\systemprofile\appdata\local\microsoft\Windows\temporary internet files\content.ie5\*.*" -Recurse -ErrorAction SilentlyContinue
}
Function IENavigateTo([string] $url, [int] $delayTime = 50, [int]$level = 0, [switch]$file, [switch]$twoTab) {
	if ($url) {
		if ($url.ToUpper().StartsWith("HTTP")) {
			try {
				$global:ie.Navigate("about:blank")

				IEWaitForPage $delayTime
				
				$startTime = [System.DateTime]::Now

				$global:ie.Navigate($url)

				IEWaitForPage $delayTime

				$timespan = [System.DateTime]::Now.Subtract($startTime)
				if (-not $file){
					Write-Host "$($timespan)"
				}
				else {
					[Microsoft.VisualBasic.Interaction]::AppActivate($global:ieproc.Id)
					Start-Sleep -Milliseconds 1000
					[System.Windows.Forms.SendKeys]::Sendwait("{TAB}");
					Start-Sleep -Milliseconds 500
					if ($twoTab)
					{
						[System.Windows.Forms.SendKeys]::Sendwait("{TAB}");
						Start-Sleep -Milliseconds 500
					}
					
					[System.Windows.Forms.SendKeys]::Sendwait("s");
					Start-Sleep -Milliseconds 500
					[System.Windows.Forms.SendKeys]::Sendwait("{ENTER}");  

					$startTime = [System.DateTime]::Now

					$folder = "$env:userprofile\Downloads"
					while ((Get-ChildItem -Path $folder | Where-object {$_.Name.EndsWith(".partial")}).Count -gt 0){
						Start-Sleep -Milliseconds 100
					}

					$timespan = [System.DateTime]::Now.Subtract($startTime)
					Write-Host "$($timespan)"

					[System.Windows.Forms.SendKeys]::Sendwait("{TAB}");
					Start-Sleep -Milliseconds 500
					[System.Windows.Forms.SendKeys]::Sendwait("o{ESC} ");

					IEWaitForPage 1000
				}
			} 
			catch {
				try {
					$pid = $global:ieproc.id
					Write-Host "IE not responding.  Closing process ID $pid"
					$global:ie.Quit()
				}
				catch {}

				$global:ieproc | Stop-Process -Force -ErrorAction SilentlyContinue

				Write-Host "Press any key to exit ..."
				$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
				exit 1
			}
		}
	}
}

Function IEWaitForPage([int] $delayTime = 10) {
	Start-Sleep -Milliseconds $delayTime 
	While ($global:ie.ReadyState -ne 4 -or $global:ie.Busy) { Start-Sleep -Milliseconds $delayTime }

	try
    {
		$spinnerVisible = $false
		$elements = $global:ie.document.body.getElementsByClassName("ms-Spinner-circle");
		foreach ($spinner in $elements) {
			$spinnerVisible = $spinnerVisible -or $spinner.style.visibility -ne "hidden"
		}
	    While ($elements.length -gt 1) { 
		    Start-Sleep -Milliseconds $delayTime 
		    $elements = $global:ie.document.body.getElementsByClassName("ms-Spinner-circle");
		    foreach ($spinner in $elements) {
			    $spinnerVisible = $spinnerVisible -or $spinner.style.visibility -ne "hidden"
		    }
	    }
    }
    catch{
	    $_.Exception
    }
	
}

$global:path = $MyInvocation.MyCommand.Path

IEOpen

DoTesting "https://tenant.sharepoint.com" "https://tenant.sharepoint.com/Sites/Communication" 5
DoTesting "" "https://tenant.sharepoint.com/SitePages/Home.aspx" 5
DoTesting "" "https://tenant.sharepoint.com/Sites/Communication" 5
DoTesting "" "https://tenant.sharepoint.com/SitePages/Home.aspx" 5
DoTesting "" "https://tenant.sharepoint.com/Sites/Communication" 5
DoTesting "" "https://tenant.sharepoint.com/SitePages/Home.aspx" 5
DoTesting "" "https://tenant.sharepoint.com/Sites/Communication" 5
DoTesting "" "https://tenant.sharepoint.com/SitePages/Home.aspx" 5
DoTesting "" "https://tenant.sharepoint.com/Sites/Communication" 5
DoTesting "" "https://tenant.sharepoint.com/SitePages/Home.aspx" 5
DoTesting "" "https://tenant.sharepoint.com/Sites/Communication" 5

IEClose

Write-Host "Done!" 
Write-Host "Press any key to continue ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
