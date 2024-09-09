Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$form = New-Object System.Windows.Forms.Form
$form.Text = "SSL Certificate Checker GUI"
$form.Size = New-Object System.Drawing.Size(600, 450)
$form.StartPosition = 'CenterScreen'

$scriptBlock = {
    param([string]$FQDN)

    function DisableSSLValidation { 
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    function CheckSSLCertificate {
        param([string]$FQDN)

        DisableSSLValidation
        
        $details = [pscustomobject]@{
            FQDN       = $FQDN
            Issuer     = 'N/A'
            IssueDate  = 'N/A'
            ExpiryDate = 'N/A'
            Thumbprint = 'N/A'
            Status     = 'Error'
        }

        try {
            $uri = New-Object System.UriBuilder("https://$FQDN")
            $hostname = $uri.Host
            $port = if ($uri.Port -ne 443) { $uri.Port } else { 443 }

            $tcpClient = New-Object System.Net.Sockets.TcpClient
            $tcpClient.Connect($hostname, $port)

            $sslStream = New-Object System.Net.Security.SslStream($tcpClient.GetStream(), $false, ({$true}))

            $sslStream.AuthenticateAsClient($hostname)

            $certificate = $sslStream.RemoteCertificate

            if ($certificate -ne $null) {
                $details.Issuer = $certificate.Issuer
                $details.IssueDate = [datetime]::Parse($certificate.GetEffectiveDateString()).ToString('yyyy-MM-dd')
                $details.ExpiryDate = [datetime]::Parse($certificate.GetExpirationDateString()).ToString('yyyy-MM-dd')
                $details.Thumbprint = $certificate.GetCertHashString()
                $validDate = [datetime]::Parse($certificate.GetExpirationDateString())

                $details.Status = if ($validDate -gt [datetime]::Now) { 'Valid' } else { 'Expired' }
            }

            $sslStream.Close()
            $tcpClient.Close()
        } catch {
            $details.Status = "Error: $_"
        }

        return $details
    }

    CheckSSLCertificate -FQDN $FQDN
}

# GUI components setup
$browseInputButton = New-Object System.Windows.Forms.Button
$browseInputButton.Location = New-Object System.Drawing.Point(10, 10)
$browseInputButton.Size = New-Object System.Drawing.Size(150, 23)
$browseInputButton.Text = "Browse Input CSV"
$browseInputButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $inputTextbox.Text = $openFileDialog.FileName
    }
})
$form.Controls.Add($browseInputButton)

$browseOutputButton = New-Object System.Windows.Forms.Button
$browseOutputButton.Location = New-Object System.Drawing.Point(10, 40)
$browseOutputButton.Size = New-Object System.Drawing.Size(150, 23)
$browseOutputButton.Text = "Browse Output CSV"
$browseOutputButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
    $saveFileDialog.Filter = "CSV files (*.csv)|*.csv"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $outputTextbox.Text = $saveFileDialog.FileName
    }
})
$form.Controls.Add($browseOutputButton)

$inputTextbox = New-Object System.Windows.Forms.TextBox
$inputTextbox.Location = New-Object System.Drawing.Point(170, 10)
$inputTextbox.Size = New-Object System.Drawing.Size(310, 23)
$form.Controls.Add($inputTextbox)

$outputTextbox = New-Object System.Windows.Forms.TextBox
$outputTextbox.Location = New-Object System.Drawing.Point(170, 40)
$outputTextbox.Size = New-Object System.Drawing.Size(310, 23)
$form.Controls.Add($outputTextbox)

$verboseCheckbox = New-Object System.Windows.Forms.CheckBox
$verboseCheckbox.Location = New-Object System.Drawing.Point(10, 100)
$verboseCheckbox.Size = New-Object System.Drawing.Size(470, 23)
$verboseCheckbox.Text = "Enable Verbose Output"
$form.Controls.Add($verboseCheckbox)

$verboseOutputTextbox = New-Object System.Windows.Forms.TextBox
$verboseOutputTextbox.Location = New-Object System.Drawing.Point(10, 130)
$verboseOutputTextbox.Size = New-Object System.Drawing.Size(570, 200)
$verboseOutputTextbox.Multiline = $true
$verboseOutputTextbox.ScrollBars = "Vertical"
$form.Controls.Add($verboseOutputTextbox)

$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Location = New-Object System.Drawing.Point(10, 70)
$checkButton.Size = New-Object System.Drawing.Size(470, 23)
$checkButton.Text = "Check SSL Certificates and Save Results"
$checkButton.Add_Click({
    # Asynchronous processing using Runspaces
    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.Open()
    $pipeline = $runspace.CreatePipeline()

    $script = {
        param($inputPath, $outputPath, $verboseChecked, $verboseBox, $scriptBlock)

        $domains = Import-Csv -Path $inputPath
        $results = @()

        foreach ($domain in $domains) {
            $cleanDomain = $domain.FQDN -replace "^https://", "" -replace "^http://", ""
            $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $cleanDomain
            
            if ($job | Wait-Job -Timeout 2) {
                $status = Receive-Job -Job $job
            } else {
                Stop-Job -Job $job
                $status = [pscustomobject]@{
                    FQDN       = $cleanDomain
                    Issuer     = 'N/A'
                    IssueDate  = 'N/A'
                    ExpiryDate = 'N/A'
                    Thumbprint = 'N/A'
                    Status     = 'Error: Timeout'
                }
            }
            Remove-Job -Job $job

            $status.PSObject.Properties.Remove("PSComputerName")
            $status.PSObject.Properties.Remove("RunspaceId")
            $status.PSObject.Properties.Remove("PSShowComputerName")

            $results += $status

            if ($verboseChecked -eq $true) {
                $entry = "FQDN: $($status.FQDN)`r`nIssuer: $($status.Issuer)`r`nIssue Date: $($status.IssueDate)`r`nExpiry Date: $($status.ExpiryDate)`r`nThumbprint: $($status.Thumbprint)`r`nStatus: $($status.Status)`r`n`r`n"
                $verboseBox.BeginInvoke([System.Windows.Forms.MethodInvoker]{
                    $verboseBox.AppendText($entry)
                })
            }
        }

        $results | Export-Csv -Path $outputPath -NoTypeInformation
        [System.Windows.Forms.MessageBox]::Show("SSL certificate check completed and results saved")
    }

    $pipeline.Commands.AddScript($script)
    $pipeline.Commands[0].Parameters.Add("inputPath", $inputTextbox.Text)
    $pipeline.Commands[0].Parameters.Add("outputPath", $outputTextbox.Text)
    $pipeline.Commands[0].Parameters.Add("verboseChecked", $verboseCheckbox.Checked)
    $pipeline.Commands[0].Parameters.Add("verboseBox", $verboseOutputTextbox)
    $pipeline.Commands[0].Parameters.Add("scriptBlock", $scriptBlock)
    $pipeline.InvokeAsync()
})
$form.Controls.Add($checkButton)

# Author Credit Label
$authorLabel = New-Object System.Windows.Forms.Label
$authorLabel.Location = New-Object System.Drawing.Point(10, 340)
$authorLabel.Size = New-Object System.Drawing.Size(570, 25)
$authorLabel.Text = "Written by GPT-4o, in spite of the utterly useless assistance of a human meatbag. 
The input CSV column must be labeled FQDN, and port numbers are supported (e.g., domain.com.au:8443)."
$form.Controls.Add($authorLabel)

$form.ShowDialog()

# Reset SSL certificate validation after the script runs
[System.Net.ServicePointManager]::ServerCertificateValidationCallback = $null
Pause