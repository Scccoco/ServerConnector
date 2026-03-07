Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Connector Agent Setup'
$form.Size = New-Object System.Drawing.Size(520, 360)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

$font = New-Object System.Drawing.Font('Segoe UI', 10)

function Add-Label($text, $x, $y) {
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $text
    $lbl.Location = New-Object System.Drawing.Point($x, $y)
    $lbl.Size = New-Object System.Drawing.Size(460, 24)
    $lbl.Font = $font
    $form.Controls.Add($lbl)
}

function Add-TextBox($x, $y, $w=460) {
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point($x, $y)
    $tb.Size = New-Object System.Drawing.Size($w, 26)
    $tb.Font = $font
    $form.Controls.Add($tb)
    return $tb
}

Add-Label 'Server URL:' 20 20
$tbServer = Add-TextBox 20 44
$tbServer.Text = 'http://62.113.36.107:8080'

Add-Label 'Device ID:' 20 80
$tbDevice = Add-TextBox 20 104
$tbDevice.Text = ('pc-' + $env:COMPUTERNAME.ToLower())

Add-Label 'Device token:' 20 140
$tbToken = Add-TextBox 20 164

Add-Label 'Heartbeat interval (sec):' 20 200
$tbInterval = Add-TextBox 20 224 120
$tbInterval.Text = '60'

$cbPy = New-Object System.Windows.Forms.CheckBox
$cbPy.Text = 'Install Python automatically if missing'
$cbPy.Location = New-Object System.Drawing.Point(20, 258)
$cbPy.Size = New-Object System.Drawing.Size(460, 24)
$cbPy.Font = $font
$form.Controls.Add($cbPy)

$btnInstall = New-Object System.Windows.Forms.Button
$btnInstall.Text = 'Install'
$btnInstall.Location = New-Object System.Drawing.Point(300, 290)
$btnInstall.Size = New-Object System.Drawing.Size(90, 30)
$btnInstall.Font = $font
$form.Controls.Add($btnInstall)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = 'Cancel'
$btnCancel.Location = New-Object System.Drawing.Point(400, 290)
$btnCancel.Size = New-Object System.Drawing.Size(90, 30)
$btnCancel.Font = $font
$form.Controls.Add($btnCancel)

$btnCancel.Add_Click({ $form.Close() })

$btnInstall.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($tbToken.Text)) {
            [System.Windows.Forms.MessageBox]::Show('Enter device token.', 'Connector Agent') | Out-Null
            return
        }

        $installScript = Join-Path $PSScriptRoot 'install_agent.ps1'
        $params = @{
            ServerUrl = $tbServer.Text.Trim()
            DeviceId = $tbDevice.Text.Trim()
            DeviceToken = $tbToken.Text.Trim()
            HeartbeatSeconds = [int]$tbInterval.Text.Trim()
        }
        if ($cbPy.Checked) {
            $params.InstallPythonIfMissing = $true
        }

        & $installScript @params | Out-Null
        [System.Windows.Forms.MessageBox]::Show('Installation complete. Desktop shortcut created.', 'Connector Agent') | Out-Null
        $form.Close()
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Install error: $($_.Exception.Message)", 'Connector Agent') | Out-Null
    }
})

[void]$form.ShowDialog()
