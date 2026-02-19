Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Press Escape first to close any modals
[System.Windows.Forms.SendKeys]::SendWait("{ESC}")
Start-Sleep -Milliseconds 500

# Open the dashboard fresh - it will get focus
Start-Process "C:\Users\Eboja\skyflowers-dashboard\sky_flowers_commercial.html"
Start-Sleep -Seconds 4

# Screenshot 1 - top of page (KPIs)
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

$bmp1 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g1 = [System.Drawing.Graphics]::FromImage($bmp1)
$g1.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp1.Save('C:\Users\Eboja\skyflowers-dashboard\dash_01_kpis.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g1.Dispose()
$bmp1.Dispose()
Write-Output "Screenshot 1 (KPIs) saved"

# Press F11 for fullscreen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
Start-Sleep -Seconds 2

# Screenshot fullscreen top
$bmp1b = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g1b = [System.Drawing.Graphics]::FromImage($bmp1b)
$g1b.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp1b.Save('C:\Users\Eboja\skyflowers-dashboard\dash_01_fullscreen.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g1b.Dispose()
$bmp1b.Dispose()
Write-Output "Screenshot 1b (fullscreen top) saved"

# Scroll down
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bmp2 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g2 = [System.Drawing.Graphics]::FromImage($bmp2)
$g2.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp2.Save('C:\Users\Eboja\skyflowers-dashboard\dash_02_categories.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g2.Dispose()
$bmp2.Dispose()
Write-Output "Screenshot 2 (categories) saved"

# Scroll down more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bmp3 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g3 = [System.Drawing.Graphics]::FromImage($bmp3)
$g3.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp3.Save('C:\Users\Eboja\skyflowers-dashboard\dash_03_clients.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g3.Dispose()
$bmp3.Dispose()
Write-Output "Screenshot 3 (clients) saved"

# Scroll down more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bmp4 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g4 = [System.Drawing.Graphics]::FromImage($bmp4)
$g4.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp4.Save('C:\Users\Eboja\skyflowers-dashboard\dash_04_products.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g4.Dispose()
$bmp4.Dispose()
Write-Output "Screenshot 4 (products) saved"

# Scroll down more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bmp5 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g5 = [System.Drawing.Graphics]::FromImage($bmp5)
$g5.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp5.Save('C:\Users\Eboja\skyflowers-dashboard\dash_05_weekly.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g5.Dispose()
$bmp5.Dispose()
Write-Output "Screenshot 5 (weekly) saved"

# Scroll down more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bmp6 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g6 = [System.Drawing.Graphics]::FromImage($bmp6)
$g6.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp6.Save('C:\Users\Eboja\skyflowers-dashboard\dash_06_concentration.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g6.Dispose()
$bmp6.Dispose()
Write-Output "Screenshot 6 (concentration) saved"

# Exit fullscreen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
