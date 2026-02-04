Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Set-Form {
    param(
        [string]$Title,
        [int]$Width,
        [int]$Length
    )

    $BaseWindow = New-Object System.Windows.Forms.Form
    $BaseWindow.Text = $Title
    $BaseWindow.Size = New-Object System.Drawing.Size($Width, $Length)
    $BaseWindow.TopMost = $true
    $BaseWindow.StartPosition = "CenterScreen"

    return $BaseWindow
}

function Set-Label {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y
    )

    $BaseLabel = New-Object System.Windows.Forms.Label
    $BaseLabel.Text = $Text
    $BaseLabel.Location = New-Object System.Drawing.Point($X, $Y)
    $BaseLabel.AutoSize = $true

    return $BaseLabel
}

function Set-Button {
    param(
        [string]$Text,
        [int]$X,
        [int]$Y
    )

    $BaseButton = New-Object System.Windows.Forms.Button
    $BaseButton.Location = New-Object System.Drawing.Point($X, $Y)
    $BaseButton.Text = $Text

    return $BaseButton
}


$Window = Set-Form -Title "Reset Teams Cache" -Width 300 -Length 150
$Prompt = Set-Label -Text "Your Teams cache needs to be reset." -x 50 -y 25
$ApproveButton = Set-Button -Text "Allow" -X 50 -Y 50
$DenyButton = Set-Button -Text "Deny" -X 150 -Y 50

$ApproveButton.Add_Click({
    Write-Host "User allowed Teams cache reset."
    
    #Kill Teams Processes
    Write-Host "Killing Teams process..."
    $TeamsProcessIds = (Get-Process | Where-Object -Property ProcessName -eq ms-teams | Select-Object id).id
    Foreach ($ProcessId in $TeamsProcessIds) {
        Stop-Process -id $ProcessId -Force
    }

    #Time Buffer
    Start-Sleep -Seconds 2

    #Clear Cache
    try {
        Write-Host "Clearing cache files..."
        Remove-Item "C:\Users\$env:USERNAME\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\" -Confirm:$false -Force -Recurse -ErrorAction Stop

        Write-Host "Launching Teams..."
        Start-Process "msteams:"
        $Window.Close()
    } catch {
        Write-Warning "Failed to clear Teams cache."
        Write-Warning $_.Exception.Message
        exit 1
    }
})

$DenyButton.Add_Click({
    Write-Host "User declined Teams cache reset."
    $Window.Close()
})


$Window.Controls.Add($ApproveButton)
$Window.Controls.Add($DenyButton)
$Window.Controls.Add($Prompt)

$Window.ShowDialog()

exit 0