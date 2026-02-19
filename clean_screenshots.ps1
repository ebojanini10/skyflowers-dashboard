Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Wait for any installers to finish
Start-Sleep -Seconds 8

# Close any popup with Escape/Enter
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
Start-Sleep -Seconds 2
[System.Windows.Forms.SendKeys]::SendWait("{ESC}")
Start-Sleep -Seconds 1

# Open the dashboard
Start-Process "C:\Users\Eboja\skyflowers-dashboard\sky_flowers_commercial.html"
Start-Sleep -Seconds 5

# Click in the middle of the page to make sure the browser has focus
# Use Ctrl+Home to go to top
[System.Windows.Forms.SendKeys]::SendWait("^{HOME}")
Start-Sleep -Seconds 1

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# Shot 1 - header + KPIs
$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_1.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 1 saved"

# Scroll down with Space (more reliable than PGDN in browsers)
[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_2.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 2 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_3.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 3 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_4.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 4 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_5.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 5 saved"

[System.Windows.Forms.SendKeys]::SendWait("{END}")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\skyflowers-dashboard\final_6.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Final 6 saved"
