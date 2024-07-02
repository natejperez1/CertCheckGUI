Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -TypeDefinition @"
using System;
using System.Threading.Tasks;

public class TimeoutHandler
{
    public static object InvokeWithTimeout(Action action, int timeout)
    {
        var task = Task.Factory.StartNew(action);
        if (task.Wait(timeout))
            return task.Result;
        else
            throw new TimeoutException();
    }
}
"@ -ReferencedAssemblies "System.Runtime"

$form = New-Object System.Windows.Forms.Form
$form.Text = "Certificate Checker GUI"
$form.Size = New-Object System.Drawing.Size(600, 420)
$form.StartPosition = "CenterScreen"

function DisableSSLValidation {
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {
        param($sender, $certificate, $chain, $sslPolicyErrors)
        return $true
    }
}

function CheckSSLCertificate {
    param (
        [string]$FQDN
    )
    
    DisableSSLValidation

    $details = [pscustomobject]@{
        FQDN      = $FQDN
        Issuer    = "NA"
        IssueDate = "NA"
        ExpiryDate= "NA"
        Thumbprint= "NA"
        Status    = "Error"
    }

    try {
        # Extract hostname and port if specified
        $uri = New-Object System.UriBuilder -ArgumentList "https://$FQDN"
        $hostname = $uri.Host
        $port = if ($uri.Port -ne 443) { $uri.Port } else { 443 }
        
        $tcpClient = New-Object System.Net.Sockets.TcpClient($hostname, $port)
        $sslStream = New-Object System.Net.Security.SslStream($tcpClient.GetStream(), $false, {
            param (
                $sender,
                $certificate,
                $chain,
                $sslPolicyErrors
            )
            $true
        })
        
        $sslStream.AuthenticateAsClient($hostname)
        
        $certificate = $sslStream.RemoteCertificate

        if ($certificate) {
            $details.Issuer = $certificate.Issuer
            $details.IssueDate = [datetime]::Parse($certificate.GetEffectiveDateString()).ToString("yyyy-MM-dd")
            $details.ExpiryDate = [datetime]::Parse($certificate.GetExpirationDateString()).ToString("yyyy-MM-dd")
            $details.Thumbprint = $certificate.GetCertHashString()
            $validDate = [datetime]::Parse($certificate.GetExpirationDateString())

            $details.Status = if ($validDate -gt (Get-Date)) { "Valid" } else { "Expired" }
        }

        $sslStream.Close()
        $tcpClient.Close()
    } catch {
        $details.Status = "Error: $($_.Exception.Message)"
    }

    return $details
}

# Function to add timeout capability to an action
function Invoke-CommandWithTimeout {
    param (
        [ScriptBlock]$Command,
        [int]$TimeoutMilliseconds = 10000
    )

    $asyncResult = $Command.BeginInvoke()
    $completed = $asyncResult.AsyncWaitHandle.WaitOne($TimeoutMilliseconds)

    if (!$completed) {
        # Close the handle to clean up
        $asyncResult.AsyncWaitHandle.Close()
        return [pscustomobject]@{
            FQDN       = $Command.ToString()
            Issuer     = 'NA'
            IssueDate  = 'NA'
            ExpiryDate = 'NA'
            Thumbprint = 'NA'
            Status     = "Timeout after $TimeoutMilliseconds milliseconds"
        }
    }

    $result = $Command.EndInvoke($asyncResult)
    return $result
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

$outputTextbox = New-Object System.Windows.Forms.TextBox
$outputTextbox.Location = New-Object System.Drawing.Point(170, 40)
$outputTextbox.Size = New-Object System.Drawing.Size(310, 23)
$form.Controls.Add($outputTextbox)

$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Location = New-Object System.Drawing.Point(10, 70)
$checkButton.Size = New-Object System.Drawing.Size(470, 23)
$checkButton.Text = "Check SSL Certificates and Save Results"
$checkButton.Add_Click({
    $domains = Import-Csv -Path $inputTextbox.Text
    $results = @()

    foreach ($domain in $domains) {
        $cleanDomain = $domain.FQDN -replace "^https://", "" -replace "^http://", ""
        $status = Invoke-CommandWithTimeout -Command { CheckSSLCertificate -FQDN $cleanDomain } -TimeoutMilliseconds 10000
        $results += $status
    }

    if ($verboseCheckbox.Checked) {
        $results | ForEach-Object {
            $verboseOutputTextbox.AppendText("FQDN: $($_.FQDN)`r`nIssuer: $($_.Issuer)`r`nIssue Date: $($_.IssueDate)`r`nExpiry Date: $($_.ExpiryDate)`r`nThumbprint: $($_.Thumbprint)`r`nStatus: $($_.Status)`r`n`r`n")
        }
    }
    $results | Export-Csv -Path $outputTextbox.Text -NoTypeInformation
    [System.Windows.Forms.MessageBox]::Show("SSL certificate check completed and results saved.")
})
$form.Controls.Add($checkButton)

# Author credit label
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