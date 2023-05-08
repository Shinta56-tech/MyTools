#==================================================================================================================
# PCをスリープさせない
#==================================================================================================================

Add-Type -AssemblyName System.Windows.Forms

# マウスの移動を検知してタイマーをリセットする
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point(0, 0)
while ($true) {
    Start-Sleep -Seconds 180 # 感覚時間　３分
    $currentPosition = [System.Windows.Forms.Cursor]::Position
    if ($lastPosition -eq $currentPosition) {
        # Write-Host "EQ"
        [System.Windows.Forms.SendKeys]::SendWait("^") # Ctrlキーを送信する
    } else {
        # Write-Host "STOP"
    }
    $lastPosition = $currentPosition
}
