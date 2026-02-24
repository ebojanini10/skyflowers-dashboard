Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Open the dashboard fresh
Start-Process "C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\sky_flowers_commercial.html"
Start-Sleep -Seconds 6

# F11 fullscreen to ensure browser fills the screen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
Start-Sleep -Seconds 2

# Go to top
[System.Windows.Forms.SendKeys]::SendWait("^{HOME}")
Start-Sleep -Seconds 1

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# Shot 1 - top (header + KPIs)
$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_1.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_1 saved"

# Scroll down
[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_2.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_2 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_3.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_3 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_4.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_4 saved"

[System.Windows.Forms.SendKeys]::SendWait(" ")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_5.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_5 saved"

# Go to very bottom for cards I and J
[System.Windows.Forms.SendKeys]::SendWait("{END}")
Start-Sleep -Seconds 1

$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\v2_6.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "v2_6 saved"

# Exit fullscreen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
