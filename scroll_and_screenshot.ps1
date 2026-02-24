Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Focus on the browser and scroll down
Start-Sleep -Seconds 1

# Screenshot 1 - scroll down to see more cards
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bitmap.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\commercial_screenshot_2.png', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics.Dispose()
$bitmap.Dispose()
Write-Output "Screenshot 2 saved"

# Screenshot 2 - scroll more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bitmap2 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics2 = [System.Drawing.Graphics]::FromImage($bitmap2)
$graphics2.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bitmap2.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\commercial_screenshot_3.png', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics2.Dispose()
$bitmap2.Dispose()
Write-Output "Screenshot 3 saved"

# Screenshot 3 - scroll more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bitmap3 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics3 = [System.Drawing.Graphics]::FromImage($bitmap3)
$graphics3.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bitmap3.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\commercial_screenshot_4.png', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics3.Dispose()
$bitmap3.Dispose()
Write-Output "Screenshot 4 saved"

# Screenshot 4 - scroll more
[System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
Start-Sleep -Seconds 2

$bitmap4 = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics4 = [System.Drawing.Graphics]::FromImage($bitmap4)
$graphics4.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bitmap4.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\commercial_screenshot_5.png', [System.Drawing.Imaging.ImageFormat]::Png)
$graphics4.Dispose()
$bitmap4.Dispose()
Write-Output "Screenshot 5 saved"
