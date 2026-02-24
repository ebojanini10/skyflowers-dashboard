Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Close any popups
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 1
[System.Windows.Forms.SendKeys]::SendWait("{ESC}")
Start-Sleep -Seconds 1

# Open the dashboard
Start-Process "C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\sky_flowers_commercial.html"
Start-Sleep -Seconds 5

# Go to top
[System.Windows.Forms.SendKeys]::SendWait("^{HOME}")
Start-Sleep -Seconds 1

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# Shot 1 - header + KPIs
$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_1.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_1 saved"

# Scroll to bottom sections (Cards I and J)
[System.Windows.Forms.SendKeys]::SendWait("{END}")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_2.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_2 saved"
