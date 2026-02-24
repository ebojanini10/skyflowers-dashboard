Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Open dashboard fresh
Start-Process "C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\sky_flowers_commercial.html"
Start-Sleep -Seconds 5

# Go fullscreen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
Start-Sleep -Seconds 2

$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds

# Screenshot 1 - top
$bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$bmp.Save('C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\shot1.png', [System.Drawing.Imaging.ImageFormat]::Png)
$g.Dispose(); $bmp.Dispose()
Write-Output "Shot 1 saved"

# Scroll and screenshot
for ($i = 2; $i -le 6; $i++) {
    [System.Windows.Forms.SendKeys]::SendWait("{PGDN}")
    Start-Sleep -Seconds 1
    $bmp = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
    $bmp.Save("C:\Users\Eboja\clients\sky-flowers\skyflowers-dashboard\shot$i.png", [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
    Write-Output "Shot $i saved"
}

# Exit fullscreen
[System.Windows.Forms.SendKeys]::SendWait("{F11}")
