﻿#------------------------------------------------------------------------
# Source File Information (DO NOT MODIFY)
# Source ID: 8c368515-c779-4ebb-95db-cf067966c8e3
# Source File: ..\LazyTS.psproj
#------------------------------------------------------------------------
#region Project Recovery Data (DO NOT MODIFY)
<#RecoveryData:
rg0AAB+LCAAAAAAABAC1V9tu4jAQfa/Uf0C8h8t2uUlppAILqtTtrkjE00orJxmKd904sp0W9ut3
QgwiV1rJvDGZ43OOZwZbtlcQ8DcQ+zlRxLm9abXsn4L/gUAdAgzXICTlkfOl07O7x0DnFpTB49wZ
B3fD8aA/sILRaGJ9Bd+3JoPQt4JNbziaDIfBGO7srgbrpVrF28fgIPF5eCTnLEQ5HZ4+OJskChSa
kEiZfdELuvkVB3fnyzFsTRPKwvt2r+24igiVxJ1YysxbHbK1gg0IiAJYaOX7trvl79Z3QqMFF6+/
Y7lpO8cICTcXCN0tEYA/PZFAJf1j9Mb/grVk3CdMIn+/7egA6ftN9IicESkp6YSMfdDHAnkbjTzE
MUShtaLBVsFO+XyH1VOJdnZqyK863CXTn/EyY0CElY7rUtBwTeHdBQYHTNFOA9SoIx7hP0h5/CDl
EZ9ByUkZYtLBnMqU0pomSpXLkM8a1g14FGFJLc91QcqKJlRhTHr4FjVsPZc0qbqgOOfnk7UmLCm1
vRpl0scS0rLOEoFpVdOBKox5D9n55ILAQX8mr6VaVKPM+8CbJMA9VsvrpHnVbEc1qjp5DdWGhl+h
00+c5Oe5qFwCmFR/BrxzcVPkBaZ8V5TOZ03qrkBiQc93ld7zRBUd1OHMe/Hwcq0oQS5pUtVNb3XP
1eUtyuazZnVV47gV89fURh4F4pKDDGXaR02/3Wt1+yPneTXKqA/F4/oDPZ+9gm7N4ZrPlnSzIHt2
nJ41WfRD0BcaEZYC0lo5T+Tf3nORIkaY3S3lb2/sbu599h92D9S5rg0AAA==#>
#endregion
<#'
#>


[CmdletBinding()]
PARAM()
#region Source: Startup.pss
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
TQQAAB+LCAAAAAAABAC91FtLwzAUAOB3wf8Q9lx6sat10BZkshdBh5Xpa5qdjrBcxkmykX9vN8cU
fRAHC3nJScj5yMmlegGmt4D+gVpKho7hWtWjmzgbNddXhFTPyFdcUTHjAp6ohKa1FK3bxBtjquTX
7GHNvTEgO8HBHOKvEd9IwzQK3kVkcbTGcbpvEZk6YR1CrcBZpCIiczfkYI/gX/UaVN2VJS1YcZtN
8jGkd5MqOWX9qbTeWJAhjPiNq6XemXimUZog4v6kwkBId1ytzrHSvC/6ss+yZZHSnP5tvUsRZE9T
jRCmeByBWY2+BdxyBmfdjX+X8YjNUQ/iBchT+Pm2q+T799F8ALPhELBNBAAA#>
#endregion
#----------------------------------------------
#region Import Assemblies
#----------------------------------------------
[void][Reflection.Assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
[void][Reflection.Assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][Reflection.Assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][Reflection.Assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
#endregion Import Assemblies


#Define a Param block to use custom parameters in the project
#Param ($CustomParameter)
$scriptname = (Get-Variable -Name MyInvocation -Scope 0 -ValueOnly).Mycommand
Write-Verbose -Message "[$ScriptName] Startup..."

function Main {
    Param ([String]$Commandline)
    if((Show-MainForm_psf) -eq "OK")
    {    }

    $global:ExitCode = 0 #Set the exit code for the Packager
}
#endregion Source: Startup.pss

#region Source: MainForm.psf
function Show-MainForm_psf
{
#region File Recovery Data (DO NOT MODIFY)
<#RecoveryData:
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
O2vc4tzh//f/DxOSH3+qrwAA#>
#endregion
    #----------------------------------------------
    #region Import the Assemblies
    #----------------------------------------------
    [void][reflection.assembly]::Load('System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')
    [void][reflection.assembly]::Load('System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.ServiceProcess, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    [void][reflection.assembly]::Load('System.Design, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
    #endregion Import Assemblies

    #----------------------------------------------
    #region Generated Form Objects
    #----------------------------------------------
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $MainForm = New-Object -TypeName 'System.Windows.Forms.Form'
    $datagridviewOutput = New-Object -TypeName 'System.Windows.Forms.DataGridView'
    $groupbox1 = New-Object -TypeName 'System.Windows.Forms.GroupBox'
    $buttonReset = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonFilter = New-Object -TypeName 'System.Windows.Forms.Button'
    $textboxHighlight = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $buttonHighlight = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonMessage = New-Object -TypeName 'System.Windows.Forms.Button'
    $textboxComputerName = New-Object -TypeName 'System.Windows.Forms.TextBox'
    $labelComputerName = New-Object -TypeName 'System.Windows.Forms.Label'
    $buttonGetTsSession = New-Object -TypeName 'System.Windows.Forms.Button'
    $buttonProcess = New-Object -TypeName 'System.Windows.Forms.Button'
    $richtextboxStatus = New-Object -TypeName 'System.Windows.Forms.RichTextBox'
    $statusstrip1 = New-Object -TypeName 'System.Windows.Forms.StatusStrip'
    $tooltip1 = New-Object -TypeName 'System.Windows.Forms.ToolTip'
    $helpprovider1 = New-Object -TypeName 'System.Windows.Forms.HelpProvider'
    $contextmenustripTSSession = New-Object -TypeName 'System.Windows.Forms.ContextMenuStrip'
    $disconnectTSSessionToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $sendTSSessionToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $contextmenustripTSProcess = New-Object -TypeName 'System.Windows.Forms.ContextMenuStrip'
    $sendTSMessageToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $stopTSProcessToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $stopTSSessionToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $toolstripstatuslabel1 = New-Object -TypeName 'System.Windows.Forms.ToolStripStatusLabel'
    $toolstripstatuslabel2 = New-Object -TypeName 'System.Windows.Forms.ToolStripStatusLabel'
    $toolstripseparator1 = New-Object -TypeName 'System.Windows.Forms.ToolStripSeparator'
    $remoteDesktopToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $powerShellRemotingToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $toolstripseparator2 = New-Object -TypeName 'System.Windows.Forms.ToolStripSeparator'
    $remoteDesktopToolStripMenuItem1 = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $powerShellRemotingToolStripMenuItem1 = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $remoteDesktopShadowIDToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $remoteDesktopShadowToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $remoteDesktopShadowControlToolStripMenuItem = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $remoteDesktopShadowControlToolStripMenuItem1 = New-Object -TypeName 'System.Windows.Forms.ToolStripMenuItem'
    $toolstripstatuslabel3 = New-Object -TypeName 'System.Windows.Forms.ToolStripStatusLabel'
    $InitialFormWindowState = New-Object -TypeName 'System.Windows.Forms.FormWindowState'
    #endregion Generated Form Objects

    #----------------------------------------------
    # User Generated Script
    #----------------------------------------------
    Write-Verbose -Message "[$ScriptName] MainForm processing..."

    $OnLoadFormEvent={
        #Stuff to do when the Form is loading
        Write-Verbose -Message "[$ScriptName] Loading Form..."
        Set-DataGridView -DataGridView $datagridviewOutput -AlternativeRowColor -ForeColor 'black' -BackColor 'AliceBlue'
    }
    $buttonGetTsSession_Click={
        TRY{
            # Set the ContextMenuStrip for TsSession
            $datagridviewOutput.ContextMenuStrip = $contextmenustripTSSession
            # Show the progression in the Richtextbox
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSSession" -Message "Loading"

            # Get the Data from Get-TSSession and convert to datatable
            $output = Get-TSSession -ComputerName $textboxComputerName.Text -ErrorAction 'Stop' -ErrorVariable ErrorGetTSSession | ForEach-Object {
                [pscustomobject][ordered]@{
                    #Server = $item.server
                    ComputerName = $textboxComputerName.Text
                    SessionID = $_.SessionID
                    State = $_.state
                    UserAccount = $_.useraccount
                    IPAddress    = $_.IPAddress
                    ClientName = $_.ClientName
                    WindowStationName = $_.WindowStationName
                    #Username = $_.username
                    ConnectionState = $_.ConnectionState
                    ClientBuildNumber = $_.ClientBuildNumber
                    ClientDisplay = $_.ClientDisplay
                    ClientIPAddress = $_.ClientIPAddress
                    ConnectTime = $_.ConnectTime
                    CurrentTime = $_.CurrentTime
                    DisconnectTime = $_.DisconnectTime
                    #DomainName = $_.DomainName
                    LastInputTime = $_.LastInputTime
                    LoginTime = $_.LoginTime
                    }#pscustomobject properties
            } #FOREACH

            $global:outputDT = ConvertTo-DataTable -InputObject $output
            Load-DataGridView -DataGridView $datagridviewOutput -Item $outputDT -ErrorAction 'Stop' -ErrorVariable ErrorLoadDGW
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSSession" -Message "Loaded"
        }
        CATCH{
            Write-Warning -Message "$textboxComputerName.Text - Get-TSSession Error"
            Write-Warning -Message $_.Exception.Message
            IF ($ErrorGetTSSession) { Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSSession" -Message "Loading Error: $($_.Exception.Message) " -MessageColor 'Red' }
            IF ($ErrorLoadDGW) { Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Load-DataGridView" -Message "Loading Error: $($_.Exception.Message) " -MessageColor 'Red' }
            ELSE { Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Unknown Error" -Message "Last Error: $($_.Exception.Message) " -MessageColor 'Red' }
        }#CATCH
    }

    $buttonProcess_Click={
        TRY{
            # Set the ContextMenuStrip for TsProcess
            $datagridviewOutput.ContextMenuStrip = $contextmenustripTSProcess

            # Show the progression in the Richtextbox
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSProcess" -Message "Loading"

            # Get the TSProcess and output a PSObject
            $output = Get-TSProcess -ComputerName $textboxComputerName.Text | ForEach-Object {
                [pscustomobject][ordered]@{
                    #Server = $Item.Server
                    ComputerName = $textboxComputerName.Text
                    SessionId = $_.SessionId
                    ProcessName = $_.ProcessName
                    ProcessId = $_.ProcessId
                    SecurityIdentifier = $_.SecurityIdentifier
                    UnderlyingProcess = $_.UnderlyingProcess
                }#pscustomobject properties
            }#FOREACH-OBJECT

            $global:outputDT = ConvertTo-DataTable -InputObject $output
            Load-DataGridView -DataGridView $datagridviewOutput -Item $outputDT
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSProcess" -Message "Loaded"
        }#TRY
        CATCH{
            Write-Warning -Message "$textboxComputerName.Text - Get-TSProcess Error"
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Get-TSProcess" -Message "Loading Error" -MessageColor 'Red'
        }
    }

    $sendTSSessionToolStripMenuItem_Click={
        TRY{
            # ONE ROW SELECTED
            IF($datagridviewOutput.SelectedRows.Count -eq 1){
                # Input Message box to send
                [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                $Message = [Microsoft.VisualBasic.Interaction]::InputBox("Message", "Send-TSMessage", "IMPORTANT: The Server is going down for maintenance in 10 minutes. Please save your work and logoff.")
                IF ($Message)
                {
                    $Message = "$Message - Sent by $env:userdomain\$env:username"

                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sending Message to session $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)"
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "MESSAGE: $message"

                    # Sending the message
                    Send-TSMessage -ComputerName $textboxComputerName.Text -Text $message -Id $datagridviewOutput.currentrow.Cells[1].value -caption "Administrator Message" -ErrorAction 'Stop'
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Message Sent!"
                }
            }#IF($datagridviewOutput.SelectedRows.Count -eq 1)

            # MULTIPLE ROWS SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                # Get the Values for the rows Selected
                $values = @()
                FOREACH ($SelectedRow in $datagridviewOutput.SelectedRows)
                {
                    # Store the session ID in $values
                    $values += $SelectedRow.Cells[1].Value
                }#FOREACH

                IF([System.Windows.Forms.MessageBox]::Show("You selected multiple sessions ($(($values|select-object -unique).count)), do you want to continue ?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
                {
                    # Input Message box to send
                    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
                    $Message = [Microsoft.VisualBasic.Interaction]::InputBox("Message", "Send-TSMessage", "IMPORTANT: The Server is going down for maintenance in 10 minutes. Please save your work and logoff.")
                    IF ($Message)
                    {
                        $Message = "$Message - Sent by $env:userdomain\$env:username"
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "MESSAGE: $message"
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sessions to receive the Message: $(($values|Select-Object -Unique).Count) sessions"

                        # Send Message to each session
                        # if the same session was selected twice, only one command will be sent
                        FOREACH ($item in ($values | Select-Object -Unique))
                        {
                            IF ($item -eq 0) { Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sending Message to session 0 might failed..." }
                            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sending Message to session $item"
                            Send-TSMessage -ComputerName $textboxComputerName.Text -Text $Message -Id $item -caption "Administrator Message"
                            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sent Message to session $item"
                        }#FOREACH
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Message Sent!"
                    }#IF $MESSAGE
                }#IF Message Box "You selected multiple rows"
            }#IF($datagridviewOutput.SelectedRows.Count -gt 1)
        }#TRY
        CATCH{
            Write-Warning -Message "$textboxComputerName.Text - Send-TSMessage Error"
            Write-Warning -Message $_.Exception.Message
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Sending Error" -MessageColor 'Red'
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Send-TSMessage" -Message "Last Error ($($_.Exception.Message))" -MessageColor 'Red'
        }
    }

    $disconnectTSSessionToolStripMenuItem_Click={
        TRY{
            # ONE ROW SELECTED
            IF($datagridviewOutput.SelectedRows.Count -eq 1){
                IF([System.Windows.Forms.MessageBox]::Show("Do you want to Disconnect the SessionID $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Disconnecting session $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)"
                    Disconnect-TSSession -ComputerName $textboxComputerName.Text -Id $datagridviewOutput.currentrow.Cells[1].value -Force
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Session $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text) Disconnected!"
                }# Message Box
            }# IF($datagridviewOutput.SelectedRows.Count -eq 1)

            # MULTIPLE ROWS SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                # Get the Values for the Rows Selected
                $values = @()
                FOREACH ($SelectedRow in $datagridviewOutput.SelectedRows)
                {
                    # Store the session ID in $values
                    $values += $SelectedRow.Cells[1].Value
                }#FOREACH

                #SingleValues
                $Singlevalues = $values | Select-Object -Unique

                IF([System.Windows.Forms.MessageBox]::Show("You selected multiple sessions ($($Singlevalues.count)) to disconnect, do you want to continue ?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Sessions to Disconnect: $(($values | Select-Object -Unique).count)"

                    # Disconnect each sessions selected
                    # if the same session was selected twice, only one command will be sent
                    FOREACH ($item in ($values | Select-Object -Unique)){
                        IF ($item -eq 0){Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Disconnecting the session 0 might failed..."}
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Disconnecting session $item"
                        Disconnect-TSSession -ComputerName $textboxComputerName.Text -Id $item
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Session $item on $($textboxComputerName.text) should be Disconnected!"
                    }#FOREACH
                }#IF Message Box "You selected multiple rows"
            }#IF($datagridviewOutput.SelectedRows.Count -gt 1)
        }#TRY
        CATCH{
            Write-Warning -Message "$textboxComputerName.Text - Disconnect-TSSession Error"
            Write-Warning -Message $_.Exception.Message
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Disconnecting Session Failed!" -MessageColor 'red'
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Disconnect-TSSession" -Message "Last Error ($($_.Exception.Message))" -MessageColor 'Red'
        }#CATCH
    }

    $stopTSProcessToolStripMenuItem_Click={
        TRY{
            # ONE ROW SELECTED
            IF($datagridviewOutput.SelectedRows.Count -eq 1){
                IF([System.Windows.Forms.MessageBox]::Show("Do you want to stop the Process $($datagridviewOutput.currentrow.Cells[2].value) ID $($datagridviewOutput.currentrow.Cells[3].value) on $($textboxComputerName.text)?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Stopping - Process $($datagridviewOutput.currentrow.Cells[2].value) ID $($datagridviewOutput.currentrow.Cells[3].value) on $($textboxComputerName.text)"
                    Stop-TSProcess -ComputerName $textboxComputerName.Text -Id $($datagridviewOutput.currentrow.Cells[3].value) -Force
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Stopped - Process $($datagridviewOutput.currentrow.Cells[2].value) ID $($datagridviewOutput.currentrow.Cells[3].value) on $($textboxComputerName.text)"
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Reloading Process List..."

                    # Refresh the list of process
                    $buttonProcess_Click.invoke()
                }# Message box "Do you want to stop the Process"
            }#IF($datagridviewOutput.SelectedRows.Count -eq 1)

            # MULTIPLE ROWS SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                IF([System.Windows.Forms.MessageBox]::Show("You selected multiple Processes ($($datagridviewOutput.SelectedRows.Count)) to Stop, do you want to continue ?", "Question",[System.Windows.Forms.MessageBoxButtons]::YesNo) -eq "Yes")
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Processes to Stop: $($datagridviewOutput.SelectedRows.Count)"

                    # Get the Values for the Rows Selected and kill each process
                    FOREACH ($SelectedRow in $datagridviewOutput.SelectedRows) {
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Stopping - Process $($SelectedRow.Cells[2].Value) ID $($SelectedRow.Cells[3].Value) on $($textboxComputerName.text)"
                        Stop-TSProcess -ComputerName $textboxComputerName.Text -Id $($SelectedRow.Cells[3].Value) -Force
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Stopped - Process $($SelectedRow.Cells[2].Value) ID $($SelectedRow.Cells[3].Value)  on $($textboxComputerName.text)"
                    }#FOREACH
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Done. Reloading Process List..."

                    # Refresh the list of process
                    $buttonProcess_Click.invoke()
                }#IF Message Box "You selected multiple rows"
            }#IF($datagridviewOutput.SelectedRows.Count -gt 1)
        }#TRY
        CATCH{
            Write-Warning -Message "$textboxComputerName.Text - Disconnect-TSSession Error"
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSProcess" -Message "Stopping Process $($datagridviewOutput.currentrow.Cells[2].value) ID $($datagridviewOutput.currentrow.Cells[3].value) on $($textboxComputerName.text) Failed!" -MessageColor 'red'
        }#CATCH
    }

    $buttonExit_Click = {
        # Close the GUI
        $MainForm.Close()
    }

    $datagridviewOutput_ColumnHeaderMouseClick=[System.Windows.Forms.DataGridViewCellMouseEventHandler]{
        IF($datagridviewOutput.DataSource -is [System.Data.DataTable])
        {
            #$column = $datagridviewOutput.Columns[$_.Columns]
            $column = $datagridviewOutput.Columns[$_.ColumnIndex]
            $direction = [System.ComponentModel.ListSortDirection]::Ascending

            IF($column.HeaderCell.SortGlyphDirection -eq 'Descending'){
                $direction = [System.ComponentModel.ListSortDirection]::Descending
            }#IF($column.HeaderCell.SortGlyphDirection -eq 'Descending')
            $datagridviewOutput.Sort($datagridviewOutput.Columns[$_.ColumnIndex], $direction)

        }#IF($datagridviewOutput.DataSource -is [System.Data.DataTable])
    }

    $toolstripstatuslabel1_Click = {
        # Open LazyWinAdmin Blog
        Start-Process -FilePath $($Configuration.author.website.uri)
    }

    $buttonMessage_Click={
        $sendTSSessionToolStripMenuItem_Click.Invoke()
    }

    $stopTSSessionToolStripMenuItem_Click={
        TRY
        {
            # ONE ROW SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -eq 1)
            {
                IF ([System.Windows.Forms.MessageBox]::Show("Do you want to Close the SessionID $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Closing session $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)"
                    Stop-TSSession -ComputerName $textboxComputerName.Text -Id $datagridviewOutput.currentrow.Cells[1].value -Force
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Session $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text) Closed!"
                }# Message Box
            }# IF($datagridviewOutput.SelectedRows.Count -eq 1)

            # MULTIPLE ROWS SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                # Get the Values for the Rows Selected
                $values = @()
                FOREACH ($SelectedRow in $datagridviewOutput.SelectedRows)
                {
                    # Store the session ID in $values
                    $values += $SelectedRow.Cells[1].Value
                }#FOREACH

                $SingleValues = $values | Select-Object -Unique

                IF ([System.Windows.Forms.MessageBox]::Show("You selected multiple sessions ($($SingleValues.count)) to Close, do you want to continue ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Sessions to Close: $(($values | Select-Object -Unique).count)"

                    # Disconnect each sessions selected
                    # if the same session was selected twice, only one command will be sent
                    FOREACH ($item in ($values | Select-Object -Unique))
                    {
                        IF ($item -eq 0) { Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Closing the session 0 might failed..." }
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Closing session $item"
                        Stop-TSSession -ComputerName $textboxComputerName.Text -Id $item -Force
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Session $item on $($textboxComputerName.text) should be Closed!"
                    }#FOREACH
                }#IF Message Box "You selected multiple rows"
            }#IF($datagridviewOutput.SelectedRows.Count -gt 1)

            # Refresh the list of Session(s)
            $buttonGetTsSession_Click.invoke()
        }#TRY
        CATCH
        {
            Write-Warning -Message "$textboxComputerName.Text - Stop-TSSession Error"
            Write-Warning -Message $_.Exception.Message
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Closing Session Failed!" -MessageColor 'red'
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Stop-TSSession" -Message "Last Error ($($_.Exception.Message))" -MessageColor 'Red'
        }#CATCH
    }

    $remoteDesktopToolStripMenuItem_Click={
        Start-Process -FilePath 'mstsc' -ArgumentList "/v:$($textboxComputerName.Text):3389 /admin"
    }

    $powerShellRemotingToolStripMenuItem_Click={
        Start-Process powershell.exe -ArgumentList "-noexit -command Enter-PSSession -ComputerName $($textboxComputerName.Text)"
    }

    $buttonHighlight_Click = {
        if ($textboxHighlight.Text)
        {
            Clear-DataGridViewSelection -DataGridView $datagridviewOutput
            Reset-DataGridViewFormat -DataGridView $datagridviewOutput
            Set-DataGridView -DataGridView $datagridviewOutput -AlternativeRowColor -BackColor 'AliceBlue' -ForeColor 'Black'
            Find-DataGridViewValue -DataGridView $datagridviewOutput -RowBackColor 'Yellow' -Value $textboxHighlight.Text
        }
    }

    $buttonReset_Click={
        Reset-DataGridViewFormat -DataGridView $datagridviewOutput
        Reset-TextBox -TextBox $textboxHighlight
    }

    $buttonFilter_Click={
        Set-DataGridViewFilter -DataGridView $datagridviewOutput -AllColumns -DataTable $outputDT -Filter $textboxHighlight.Text
    }

    $remoteDesktopShadowIDToolStripMenuItem_Click = {
        TRY
        {
            # ONE ROW SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -eq 1)
            {
                IF ([System.Windows.Forms.MessageBox]::Show("Do you want to connect the SessionID $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    $SessionID = $($datagridviewOutput.currentrow.Cells[1].value)
                    $Machine = $($datagridviewOutput.currentrow.Cells[0].value)
                    IF ([System.Windows.Forms.MessageBox]::Show("Do you want to prompt the user for consent ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes') {
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Opening Shadow (view) session $SessionID on $($textboxComputerName.text) ..."
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Prompting user for consent ..."
                        $Arguments = "/v:$($Machine) /shadow:$SessionID"
                    } ELSE {
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Opening Shadow (view) session $SessionID on $($textboxComputerName.text) ..."
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "No consent prompt"
                        $Arguments = "/v:$($Machine) /shadow:$SessionID /noConsentPrompt"
                    }
                    Start-Process -FilePath 'mstsc' -ArgumentList $Arguments #Default port is 3389
                }# Message Box
            }#IF
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Please select only one row" -MessageColor 'red'
            }
        }
        CATCH
        {
            Write-Warning -Message "$textboxComputerName.Text - Remote Desktop (Shadow/View) Error"
            Write-Warning -Message $_.Exception.Message
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Error while opening session" -MessageColor 'red'
        }
    }
    $remoteDesktopShadowControlToolStripMenuItem_Click={
        TRY
        {
            # ONE ROW SELECTED
            IF ($datagridviewOutput.SelectedRows.Count -eq 1)
            {
                IF ([System.Windows.Forms.MessageBox]::Show("Do you want to connect the SessionID $($datagridviewOutput.currentrow.Cells[1].value) on $($textboxComputerName.text)?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes')
                {
                    $SessionID = $($datagridviewOutput.currentrow.Cells[1].value)
                    $Machine = $($datagridviewOutput.currentrow.Cells[0].value)
                    IF ([System.Windows.Forms.MessageBox]::Show("Do you want to prompt the user for consent ?", "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes') {
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Opening Shadow (control) session $SessionID on $($textboxComputerName.text) ..."
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Prompting user for consent ..."
                        $Arguments = "/v:$($Machine) /shadow:$SessionID /control"
                    } ELSE {
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Opening Shadow (control) session $SessionID on $($textboxComputerName.text) ..."
                        Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "No consent prompt"
                        $Arguments = "/v:$($Machine) /shadow:$SessionID /noConsentPrompt /control"
                    }
                    Start-Process -FilePath 'mstsc' -ArgumentList $Arguments #Default port is 3389
                }# Message Box
            }#IF
            IF ($datagridviewOutput.SelectedRows.Count -gt 1)
            {
                Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Please select only one row" -MessageColor 'red'
            }
            IF ($datagridviewOutput.SelectedRows.Count -eq 0)
            {
                New-MessageBox -Buttons 'OK' -Icon 'Information' -Message "Please select a row" -Title "No session specified"
            }
        }
        CATCH
        {
            Write-Warning -Message "$textboxComputerName.Text - Remote Desktop (Shadow/Control) Error"
            Write-Warning -Message $_.Exception.Message
            Append-RichtextboxStatus -ComputerName $textboxComputerName.Text -Source "Remote Desktop" -Message "Error while opening session" -MessageColor 'red'
        }
    }


    foreach ($i in $datagridviewOutput.Columns)
    {
        $i.GetType

    }
    $textboxHighlight_TextChanged = {
        if ($textboxHighlight.Text -ne "")
        {
            Enable-Button -Button $buttonHighlight,$buttonReset
            Set-TextBox -TextBox $textboxHighlight -BackColor 'Yellow'
        }
        else
        {
            Disable-Button -Button $buttonHighlight, $buttonReset
            Set-TextBox -TextBox $textboxHighlight -BackColor 'White'
        }
    }

    $toolstripstatuslabel3_Click={
        Start-Process $($Configuration.ProjectURI)
    }
        # --End User Generated Script--
    #----------------------------------------------
    #region Generated Events
    #----------------------------------------------

    $Form_StateCorrection_Load=
    {
        #Correct the initial state of the form to prevent the .Net maximized form issue
        $MainForm.WindowState = $InitialFormWindowState
    }

    $Form_StoreValues_Closing=
    {
        #Store the control values
        $script:MainForm_datagridviewOutput = $datagridviewOutput.SelectedCells
        $script:MainForm_textboxHighlight = $textboxHighlight.Text
        $script:MainForm_textboxComputerName = $textboxComputerName.Text
        $script:MainForm_richtextboxStatus = $richtextboxStatus.Text
    }


    $Form_Cleanup_FormClosed=
    {
        #Remove all event handlers from the controls
        try
        {
            $buttonReset.remove_Click($buttonReset_Click)
            $buttonFilter.remove_Click($buttonFilter_Click)
            $textboxHighlight.remove_TextChanged($textboxHighlight_TextChanged)
            $buttonHighlight.remove_Click($buttonHighlight_Click)
            $buttonMessage.remove_Click($buttonMessage_Click)
            $buttonGetTsSession.remove_Click($buttonGetTsSession_Click)
            $buttonProcess.remove_Click($buttonProcess_Click)
            $MainForm.remove_Load($OnLoadFormEvent)
            $disconnectTSSessionToolStripMenuItem.remove_Click($disconnectTSSessionToolStripMenuItem_Click)
            $sendTSSessionToolStripMenuItem.remove_Click($sendTSSessionToolStripMenuItem_Click)
            $sendTSMessageToolStripMenuItem.remove_Click($sendTSSessionToolStripMenuItem_Click)
            $stopTSProcessToolStripMenuItem.remove_Click($stopTSProcessToolStripMenuItem_Click)
            $stopTSSessionToolStripMenuItem.remove_Click($stopTSSessionToolStripMenuItem_Click)
            $toolstripstatuslabel1.remove_Click($toolstripstatuslabel1_Click)
            $remoteDesktopToolStripMenuItem.remove_Click($remoteDesktopToolStripMenuItem_Click)
            $powerShellRemotingToolStripMenuItem.remove_Click($powerShellRemotingToolStripMenuItem_Click)
            $remoteDesktopToolStripMenuItem1.remove_Click($remoteDesktopToolStripMenuItem_Click)
            $powerShellRemotingToolStripMenuItem1.remove_Click($powerShellRemotingToolStripMenuItem_Click)
            $remoteDesktopShadowIDToolStripMenuItem.remove_Click($remoteDesktopShadowIDToolStripMenuItem_Click)
            $remoteDesktopShadowToolStripMenuItem.remove_Click($remoteDesktopShadowIDToolStripMenuItem_Click)
            $remoteDesktopShadowControlToolStripMenuItem.remove_Click($remoteDesktopShadowControlToolStripMenuItem_Click)
            $remoteDesktopShadowControlToolStripMenuItem1.remove_Click($remoteDesktopShadowControlToolStripMenuItem_Click)
            $toolstripstatuslabel3.remove_Click($toolstripstatuslabel3_Click)
            $MainForm.remove_Load($Form_StateCorrection_Load)
            $MainForm.remove_Closing($Form_StoreValues_Closing)
            $MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
        }
        catch { Out-Null <# Prevent PSScriptAnalyzer warning #> }
    }
    #endregion Generated Events

    #----------------------------------------------
    #region Generated Form Code
    #----------------------------------------------
    $MainForm.SuspendLayout()
    $groupbox1.SuspendLayout()
    $statusstrip1.SuspendLayout()
    $contextmenustripTSSession.SuspendLayout()
    $contextmenustripTSProcess.SuspendLayout()
    #
    # MainForm
    #
    $MainForm.Controls.Add($datagridviewOutput)
    $MainForm.Controls.Add($groupbox1)
    $MainForm.Controls.Add($richtextboxStatus)
    $MainForm.Controls.Add($statusstrip1)
    $MainForm.AutoScaleDimensions = '6, 13'
    $MainForm.AutoScaleMode = 'Font'
    $MainForm.ClientSize = '723, 340'
    #region Binary Data
    $MainForm.Icon = [System.Convert]::FromBase64String('
AAABAAQAGBgAAAEAIACICQAARgAAACAgAAABACAAqBAAAM4JAAAQEAAAAQAgAGgEAAB2GgAAFhYA
AAEAIAAQCAAA3h4AACgAAAAYAAAAMAAAAAEAIAAAAAAAAAkAANcNAADXDQAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAXAAAAKwAAACsAAAArAAAAFwAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAVVVVAVVVVQFVVVUGVVVVIFVVVW1VVVW4VVVV3VVVVfNVVVX9
VVVV/8DAwP+cnJz/hoaG/1VVVf9VVVX/VVVV9VVVVdVVVVWWVVVVVVVVVSxVVVUcVVVVHAAAAAAA
AAAAyMjIAsjIyATIyMgHyMjIHcjIyGDIyMiuyMjI28jIyPOgoKD9gICA/+Li4v/IyMj/oKCg/2Zm
Zv+AgID/oKCg9cjIyNfIyMiXyMjIVcjIyCzIyMgdyMjIHQAAAAAAAAAAVVVVC1VVVQ9VVVUZVVVV
M1VVVW9VVVWzVVVV4FVVVfRVVVX9VVVV/+Li4v/i4uL/x8fH/1VVVf9VVVX/VVVV+lVVVeFVVVWx
VVVVcFVVVUBVVVUrVVVVKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAzAAAAXFNWVf+goKD/U1ZV/wAAAFwAAAAzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEpOTP/IyMj/Sk5M
/wAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAg4CAZoOAgP+DgID/
g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/2ZnZv+lo6P/Zmdm/4OAgP+DgID/g4CA/4OAgP+D
gID/g4CA/4OAgP+DgID/g4CAZgAAAAAAAAAAg4CA/8vJyf/6+Pj/+vj4//r4+P/6+Pj/+vj4//r4
+P/6+Pj/+vj4/+Tj4//6+Pj/5OPj//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4//r4+P/Lycn/g4CA
/wAAAAAAAAAAg4CA//n4+P/s6un/7Ono/+zp6P/s6ej/7Onp/+zp6f/s6en/7Onp/+zp6f/s6en/
7Onp/+zp6f/s6en/7Onp/+zp6f/s6en/6+jm/+zq6f/5+Pj/g4CA/wAAAAAAAAAAg4CA//n4+P+H
SiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dK
IP+HSiD/h0og/4dKIP/5+Pj/g4CA/wAAAAAAAAAAg4CA//r5+f+HSiD/qHFG/6x3S/+yfFD/toFW
/7uHW//AjGD/xZNm/8mYa//NnXD/zJtu/8eWaf/DkWT/vope/7mFWf+0gFT/sHpP/4dKIP/5+Pj/
g4CA/wAAAAAAAAAAg4CA//r5+f+JTiX/uYxp/7yQbv+/lHH/7vDw/+zu7v/s7u7/7O7u/+zu7v/J
mWz/yZdq/8WUZ//CjmP/vYpe/7mEWP+0f1P/r3pO/4dKIP/5+Pj/g4CA/wAAAAAAAAAAg4CA//r5
+f+JTiX/uY1r/7ySb//AlnP/8PHx//Dx8f/v8PD/7e/v/+zu7v/Fk2b/xZJm/8KQZP++i1//u4db
/7aDV/+yfVL/rnhM/4dKIP/5+Pj/g4CA/wAAAAAAAAAAg4CA//r5+f+JTiX/uY9u/72Tcf/AlnX/
8PHx//Dx8f/w8fH/8PHx/+/x8f/Fl27/wI1g/+zu7v/s7u7/7O7u/7WAU/+we0//rXZK/4dKIP/6
+fn/g4CA/wAAAAAAAAAAg4CA//r5+f+JTiX/uo9v/72Tcv+/l3b/pHZX/6R2V/+kdlf/pHZX/6R2
V//LpYT/y6KB/+7v7//s7u7/7O7u/7J8Uf+teE3/qnNI/4dKIP/6+fn/g4CA/wAAAAAAAAAAg4CA
//r5+f+JTiX/upFx/72UdP+/lnf/wpp7/8Scff/Gn3//yKGB/8ijg//Jo4P/yaSD//Hy8v/w8vL/
7vDw/7J9U/+teU7/qHJH/4dKIP/6+fn/g4CA/wAAAAAAAAAAg4CA//r5+f+JTiX/uZFz/7yUdf+/
lnj/wZl6/8Ocff/Fn3//xp+A/8eggv/IooL/yKKC/6l9X/+pfV//qX1f/8KbfP/AmXr/vZR1/4dK
IP/6+fn/g4CA/wAAAAAAAAAAg4CA//r5+f+JTiX/uZFz/7uUdv++l3j/wJl7/8Kbff/DnX//xZ6A
/8Wggf/FoIL/xaCB/8Wfgf/Enn//w5x+/8CafP+/mHr/vJN2/4dKIP/6+fn/g4CA/wAAAAAAAAAA
g4CA//r5+f+JTST/8vPz//Lz8//y8/P/8vPz//Lz8//y8/P/8vPz//Lz8//y8/P/8vPz//Lz8//y
8/P/8vPz//Lz8//y8/P/8fPz/4dKIP/6+fn/g4CA/wAAAAAAAAAAg4CA//v5+f+HSiD/h0og/4dK
IP+HSiD/h0og/4dLIf+HSyH/h0sh/4hMIv+ITCL/iEwi/4hMI/+ITCP/iEwj/4lNJP+JTST/iU0k
/4dKIP/7+fn/g4CA/wAAAAAAAAAAg4CA//X09Pv7+fn/+/n5//v5+f/6+Pj/+vj4//r4+P/6+Pj/
+vj4//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4//r5+f/29PT8g4CA/wAAAAAA
AAAAg4CAo4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/4OA
gP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CAwQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAD///8A/8H/AIAAAQCAAAEAgAABAP+A/wD/4/8AgAABAIAAAQCA
AAEAgAABAIAAAQCAAAEAgAABAIAAAQCAAAEAgAABAIAAAQCAAAEAgAABAIAAAQCAAAEAgAABAP//
/wAoAAAAIAAAAEAAAAABACAAAAAAAAAQAAASCwAAEgsAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEBACEBA
QBBLS0sRS0tLEUtLSxFLS0sRRkZGRUdHR+5FRUX3RUVF90VFRfdFRUX3RUVF90hISMBLS0sRS0tL
EUtLSxFLS0sRS0tLEUlJSQ5VVVUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAI6Ojgl0dHRsdHR0rXV1dchycnLfcnJy9nJycvZYWFj7V1dX/2hoaP9oaGj/aGho/2hoaP9o
aGj/R0dH/2xsbPdycnL2cnJy9nJycuF1dXXIfHx8pYGBgV0AAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAv7+/DJaWlnmcnJy6mpqa0JaWluCSkpLykpKS8mdnZ/mFhYX/tra2
/8PDw//Dw8P/w8PD/729vf9NTU3+iYmJ85KSkvKSkpLylpaW4pqamtClpaW1oqKibv///wEAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQEBACEBAQBBLS0sRS0tLEUtLSxFL
S0sRRUVFSkRERPZDQ0P/dHR0/3p6ev9RUVH5Q0ND/0ZGRstHR0cSS0tLEUtLSxFLS0sRS0tLEUlJ
SQ5VVVUGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEVFRZxwcHD/dXV1/1NTU+9ERERaAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACLkpCKhouJ+YWKiP+F
ioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WK
iP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/houJ+4qRj5IAAAAAiIiI
D4iNi/f29/f/////////////////////////////////////////////////////////////////
///////////////////////////////////////////////////////////////////////////2
9vb/iI2L+Y6OjgmIiIgejJGP8//////VzMP/1cvC/9XKwf/VysH/1crB/9TKwf/UysH/1MrB/9TJ
wP/UycD/08nA/9PJwP/TycD/08nA/9PJwP/SycD/0snA/9LJwP/SycD/0si//9HIv//RyL//0ci/
/9HIv//RyL//0si///////+IjYv5iIiID4iIiB6MkY/z/////4dKIP+TWjP/k1oz/5NbNP+UWzT/
lFw1/5RcNf+VXTb/lV02/5VeN/+WXjf/ll43/5ZeN/+WXjf/ll43/5ZeN/+VXTb/lV02/5VdNf+U
XDX/lFw0/5RbNP+TWzP/k1oz/5NaMv+HSiD//////4iNi/mIiIgPiIiIHoyRj/P/////h0og/690
RP+yeEj/tXxM/7d/UP+6g1T/vYZY/7+KXP/CjmD/xZFj/8eVZ//KmGv/zJtt/8ybbv/KmGv/x5Vn
/8WSZP/CjmD/wIpc/72HWP+6g1T/t39Q/7V8TP+yeEn/r3VF/45ULP//////iI2L+YiIiA+IiIge
jJGP8/////+HSiD/r3RE/7J4SP+1fEz/t39Q/7qDVP+9h1j/wIpc/8KOYP/FkWP/x5Vn/8qYa//M
m27/zJxu/8qZa//IlWj/xZJk/8KOYP/Ai1z/vYdY/7qDVP+4gFH/tXxN/7J4Sf+vdUX/jlQs////
//+IjYv5iIiID4iIiB6MkY/z/////4dKIP+wdUX/snhJ/7R8TP+3f1D/uYJT/7yFV/++iVr/wYxe
/+zu7v/s7u7/7O7u/+zu7v/s7u7/7O7u/8WSZP/Dj2H/wYxe/7+JWv+8hlf/uoJT/7d/T/+0e0v/
snhI/690RP+OVCz//////4iNi/mIiIgPiIiIHoyRj/P/////h0og/7N7Tv+1flL/uIJV/7qFWf+8
iFz/vold/76KXf+/i13/7O7u/+zu7v/s7u7/7O7u/+zu7v/s7u7/wo5f/8CLXf++iVr/vIZX/7qD
VP+4gFH/tX1N/7N5Sv+wdkb/rnJC/45ULP//////iI2L+YiIiA+IiIgejJGP8/////+HSiD/toFW
/7iFW/+7h13/vYph/76MY//Aj2b/wpJp/8STa//u8PD/7e/v/+zu7v/s7u7/7O7u/+zu7v++iFr/
vYdY/7uEVv+6glP/t39Q/7V9Tf+zekr/sXZH/65zQ/+scED/jlQr//////+IjYv5iIiID4iIiB6M
kY/z/////4dKIP/v8fH/7/Hx/+/x8f/v8fH/wJBq/8KTbP/DlG7/xJZw/+/x8f/v8fH/7/Hx/+7w
8P/t7+//7O7u/7qDVP+5glP/7O7u/+zu7v/s7u7/7O7u/+zu7v+uc0P/rHBA/6ptPf+OVCv/////
/4iNi/mIiIgPiIiIHoyRj/P/////h0og//Dy8v/a4N3/2uDd//Dy8v/ClG//w5Zy/8SXc//FmXX/
8PLy//Dy8v/w8vL/8PLy//Dy8v/v8fH/u4VZ/7V8Tf/s7u7/7O7u/+zu7v/s7u7/7O7u/6xvP/+q
bTz/p2o5/45TK///////iI2L+YiIiA+IiIgejJGP8/////+HSiD/8fPz//Hz8//x8/P/8fPz/8SZ
dv/Emnj/xpt5/8edev/x8/P/8fPz//Hz8//x8/P/8fPz//Hz8//Hnnz/xJdz/+7w8P/s7u7/7O7u
/+zu7v/s7u7/qGs7/6dpOP+lZjX/jVMr//////+IjYv5iIiID4iIiB6MkY/z/////4dKIP/y9PT/
4OXi/+Dl4v/z9PT/xp5+/8eff//IoIH/yaGB/7CIbf+wiG3/sIht/7CIbf+wiG3/sIht/8mjg//J
ooP/8/T0//L09P/w8vL/7/Dw/+3v7/+lZzb/pGU0/6RlNP+NUyv//////4iNi/mIiIgPiIiIHoyR
j/P/////h0og//P09P/09fX/9PX1//T19f/Io4X/yqSG/8qliP/Lpoj/y6aJ/8unif/Lp4r/y6eK
/8univ/Lp4r/y6eJ/8umif+3knn/t5J5/7eSef+3knn/t5J5/8aef//BlXT/u4xo/41TK///////
iI2L+YiIiA+IiIgejJGP8/////+HSiD/9PX1/+Xp5//l6ef/9fb2/82qkP/NqpD/zaqQ/82qkP/N
qpD/zquQ/86rkP/OrJH/zqyR/86rkP/Oq5D/zaqQ/82qkP/NqpD/zaqQ/82qkP/NqpD/zaqQ/82q
kP/NqpD/jVMr//////+IjYv5iIiID4iIiB6MkY/z/////4dKIP/19vb/9vf3//b39//29/f/0rOb
/9Kzm//Ss5v/0rOb/9Kzm//Ss5v/0rOb/9Kzm//Ss5v/0rOb/9Kzm//Ss5v/0rOb/9Kzm//Ss5v/
0rOb/9Kzm//Ss5v/0rOb/9Kzm/+NUyv//////4iNi/mIiIgPiIiIHoyRj/P/////h0og//b39//q
7ez/6u3s//f4+P/Xu6b/17um/9e7pv/Xu6b/17um/9e7pv/Xu6b/17um/9e7pv/Xu6b/17um/9e7
pv/Xu6b/17um/9e7pv/Xu6b/17um/9e7pv/Xu6b/17um/45TK///////iI2L+YiIiA+IiIgejJGP
8/////+HSiD/9fTz//b19P/29fT/9vX0//b19P/29fT/9vX0//b19P/29fT/9vX0//b19P/29fT/
9vX0//b19P/29fT/9vX0//b19P/29fT/9vX0//b19P/29fT/9vX0//b19P/08vD/jlQs//////+I
jYv5iIiID4CSgA6Jjoz4/Pz8/6d6W/+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dK
IP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og
/4dKIP+meFn/+fj4/4iNi/qAgIAGAAAAAIuSj8DHysn8/Pz8////////////////////////////
////////////////////////////////////////////////////////////////////////////
//////////////////////////z8/P/JzMz8jZKQvAAAAAAAAAAAgIuLFouQjr2HjIr6hYqI/4WK
iP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI
/4WKiP+Fioj/hYqI/4WKiP+Fioj/hYqI/4WKiP+Fioj/iI2L+IuPjsCFhYUXAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAP////////////////wAAB/4AAAf+AAA
D/wAAB///B//gAAAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAGAAAAB////////////////KAAAABAAAAAg
AAAAAQAgAAAAAAAABAAA1w0AANcNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFwAAACsAAAArAAAAKwAAABcAAAAAAAAAAAAAAAAAAAAAAAAAABVVVV2VVVVnFVVVclVVVXrVVVV
/FVVVf9VVVX/wMDA/5ycnP+Ghob/VVVV/1VVVf9VVVX+VVVV91VVVd5VVVWtyMjIdMjIyJvIyMjG
yMjI6sjIyPugoKD/gICA/+Li4v/IyMj/oKCg/2ZmZv+AgID/oKCg/sjIyPbIyMjYyMjIoVVVVZBV
VVWxVVVV1VVVVfBVVVX9VVVV/1VVVf/i4uL+4uLi/sfHx/5VVVX/VVVV/1VVVf5VVVX4VVVV5FVV
VbmDgID/g4CA/4OAgP+DgID/g4CA/2lmZv9VU1P/WFhY/52dnf9YWFj/VlRU/2lmZv+DgID/g4CA
/4OAgP+DgID/g4CA/+zu7v/s7u7/7O7u/+zu7v/s7u7/7O7u/52env/X2Nj/oaKi/+zu7v/s7u7/
7O7u/+zu7v/s7u7/g4CA/4OAgP/s7u7/0tPT/9LT0//S09P/0tPT/9LT0//S09P/0tPT/9LT0//S
09P/0tPT/9LT0//S09P/7O7u/4OAgP+DgID/7O7u/4VkRP+FZET/hWRE/4VkRP+FZET/hWRE/4Vk
RP+FZET/hWRE/4VkRP+FZET/hWRE/+zu7v+DgID/g4CA/+zu7v+FZET/s41n/7CJYv+wiWH/sIlh
/7CJYf+wiWH/sotk/7WPav+1j2r/tY9q/4VkRP/s7u7/g4CA/4OAgP/s7u7/hWRE/7eTbv+zjWf/
7O7u/+zu7v/s7u7/roZd/+zu7v/s7u7/roZd/66GXf+FZET/7O7u/4OAgP+DgID/7O7u/4VkRP/B
ooT/waKE/+zu7v/s7u7/7O7u/7WRbP+kZTT/pGU0/6qAVv+ofVL/hWRE/+zu7v+DgID/g4CA/+zu
7v+FZET/yrCW/8qwlv+kZTT/pGU0/6RlNP/KsJb/yrCW/8WpjP+5lnT/sIli/4VkRP/s7u7/g4CA
/4OAgP/s7u7/hWRE/9S+qf/Uvqn/1L6p/9S+qf/Uvqn/1L6p/9S+qf/Uvqn/1L6p/9S+qf+FZET/
7O7u/4OAgP+DgID/7O7u/4VkRP+FZET/hWRE/4VkRP+FZET/hWRE/4VkRP+FZET/hWRE/4VkRP+F
ZET/hWRE/+zu7v+DgID/g4CA/+zu7v/s7u7/7O7u/+zu7v/s7u7/7O7u/+zu7v/s7u7/7O7u/+zu
7v/s7u7/7O7u/+zu7v/s7u7/g4CA/4+MjMODgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA
/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/5CNjcb8HwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAKAAAABYAAAAsAAAAAQAgAAAAAACQ
BwAA1w0AANcNAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AFwAAACsAAAArAAAAKwAAABcAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABVVVUBVVVV
AVVVVQZVVVUgVVVVbVVVVbhVVVXdVVVV81VVVf1VVVX/wMDA/5ycnP+Ghob/VVVV/1VVVf9VVVX1
VVVV1VVVVZZVVVVVVVVVLFVVVRxVVVUcyMjIAsjIyATIyMgHyMjIHcjIyGDIyMiuyMjI28jIyPOg
oKD9gICA/+Li4v/IyMj/oKCg/2ZmZv+AgID/oKCg9cjIyNfIyMiXyMjIVcjIyCzIyMgdyMjIHVVV
VQtVVVUPVVVVGVVVVTNVVVVvVVVVs1VVVeBVVVX0VVVV/VVVVf/i4uL/4uLi/8fHx/9VVVX/VVVV
/1VVVfpVVVXhVVVVsVVVVXBVVVVAVVVVK1VVVSgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAADMAAABcU1ZV/6CgoP9TVlX/AAAAXAAAADMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEpOTP/IyMj/Sk5M/wAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIOAgGaDgID/g4CA/4OAgP+DgID/g4CA
/4OAgP+DgID/g4CA/4OAgP9mZ2b/paOj/2ZnZv+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/
g4CA/4OAgGaDgID/y8nJ//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4//r4+P/6+Pj/5OPj//r4+P/k
4+P/+vj4//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4/8vJyf+DgID/g4CA//n4+P/s6un/7Ono/+zp
6P/s6ej/7Onp/+zp6f/s6en/7Onp/+zp6f/s6en/7Onp/+zp6f/s6en/7Onp/+zp6f/s6en/6+jm
/+zq6f/5+Pj/g4CA/4OAgP/5+Pj/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/
h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/h0og/4dKIP+HSiD/+fj4/4OAgP+DgID/+vn5/4dKIP+o
cUb/rHdL/7J8UP+2gVb/u4db/8CMYP/Fk2b/yZhr/82dcP/Mm27/x5Zp/8ORZP++il7/uYVZ/7SA
VP+wek//h0og//n4+P+DgID/g4CA//r5+f+JTiX/uYxp/7yQbv+/lHH/7vDw/+zu7v/s7u7/7O7u
/+zu7v/JmWz/yZdq/8WUZ//CjmP/vYpe/7mEWP+0f1P/r3pO/4dKIP/5+Pj/g4CA/4OAgP/6+fn/
iU4l/7mNa/+8km//wJZz//Dx8f/w8fH/7/Dw/+3v7//s7u7/xZNm/8WSZv/CkGT/votf/7uHW/+2
g1f/sn1S/654TP+HSiD/+fj4/4OAgP+DgID/+vn5/4lOJf+5j27/vZNx/8CWdf/w8fH/8PHx//Dx
8f/w8fH/7/Hx/8WXbv/AjWD/7O7u/+zu7v/s7u7/tYBT/7B7T/+tdkr/h0og//r5+f+DgID/g4CA
//r5+f+JTiX/uo9v/72Tcv+/l3b/pHZX/6R2V/+kdlf/pHZX/6R2V//LpYT/y6KB/+7v7//s7u7/
7O7u/7J8Uf+teE3/qnNI/4dKIP/6+fn/g4CA/4OAgP/6+fn/iU4l/7qRcf+9lHT/v5Z3/8Kae//E
nH3/xp9//8ihgf/Io4P/yaOD/8mkg//x8vL/8PLy/+7w8P+yfVP/rXlO/6hyR/+HSiD/+vn5/4OA
gP+DgID/+vn5/4lOJf+5kXP/vJR1/7+WeP/BmXr/w5x9/8Wff//Gn4D/x6CC/8iigv/IooL/qX1f
/6l9X/+pfV//wpt8/8CZev+9lHX/h0og//r5+f+DgID/g4CA//r5+f+JTiX/uZFz/7uUdv++l3j/
wJl7/8Kbff/DnX//xZ6A/8Wggf/FoIL/xaCB/8Wfgf/Enn//w5x+/8CafP+/mHr/vJN2/4dKIP/6
+fn/g4CA/4OAgP/6+fn/iU0k//Lz8//y8/P/8vPz//Lz8//y8/P/8vPz//Lz8//y8/P/8vPz//Lz
8//y8/P/8vPz//Lz8//y8/P/8vPz//Hz8/+HSiD/+vn5/4OAgP+DgID/+/n5/4dKIP+HSiD/h0og
/4dKIP+HSiD/h0sh/4dLIf+HSyH/iEwi/4hMIv+ITCL/iEwj/4hMI/+ITCP/iU0k/4lNJP+JTST/
h0og//v5+f+DgID/g4CA//X09Pv7+fn/+/n5//v5+f/6+Pj/+vj4//r4+P/6+Pj/+vj4//r4+P/6
+Pj/+vj4//r4+P/6+Pj/+vj4//r4+P/6+Pj/+vj4//r5+f/29PT8g4CA/4OAgKODgID/g4CA/4OA
gP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA/4OAgP+DgID/g4CA
/4OAgP+DgID/g4CA/4OAgMH/g/wAAAAAAAAAAAAAAAAA/wH8AP/H/AAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA')
    #endregion
    $MainForm.MinimumSize = '700, 379'
    $MainForm.Name = 'MainForm'
    $MainForm.StartPosition = 'CenterScreen'
    $MainForm.Text = "$($Configuration.Name) $($Configuration.Version)"
    $MainForm.add_Load($OnLoadFormEvent)
    #
    # datagridviewOutput
    #
    $datagridviewOutput.AllowUserToAddRows = $False
    $datagridviewOutput.AllowUserToDeleteRows = $False
    $datagridviewOutput.ContextMenuStrip = $contextmenustripTSSession
    $datagridviewOutput.Dock = 'Fill'
    $datagridviewOutput.Location = '0, 57'
    $datagridviewOutput.Name = 'datagridviewOutput'
    $datagridviewOutput.ReadOnly = $True
    $datagridviewOutput.RowHeadersVisible = $False
    $datagridviewOutput.SelectionMode = 'FullRowSelect'
    $datagridviewOutput.Size = '723, 190'
    $datagridviewOutput.TabIndex = 5
    #
    # groupbox1
    #
    $groupbox1.Controls.Add($buttonReset)
    $groupbox1.Controls.Add($buttonFilter)
    $groupbox1.Controls.Add($textboxHighlight)
    $groupbox1.Controls.Add($buttonHighlight)
    $groupbox1.Controls.Add($buttonMessage)
    $groupbox1.Controls.Add($textboxComputerName)
    $groupbox1.Controls.Add($labelComputerName)
    $groupbox1.Controls.Add($buttonGetTsSession)
    $groupbox1.Controls.Add($buttonProcess)
    $groupbox1.Dock = 'Top'
    $groupbox1.Location = '0, 0'
    $groupbox1.Name = 'groupbox1'
    $groupbox1.Size = '723, 57'
    $groupbox1.TabIndex = 9
    $groupbox1.TabStop = $False
    #
    # buttonReset
    #
    $buttonReset.Anchor = 'Top, Right'
    $buttonReset.Enabled = $False
    $buttonReset.Location = '680, 32'
    $buttonReset.Name = 'buttonReset'
    $buttonReset.Size = '43, 22'
    $buttonReset.TabIndex = 13
    $buttonReset.Text = 'Reset'
    $buttonReset.UseVisualStyleBackColor = $True
    $buttonReset.add_Click($buttonReset_Click)
    #
    # buttonFilter
    #
    $buttonFilter.Anchor = 'Top, Right'
    $buttonFilter.Enabled = $False
    $buttonFilter.Location = '619, 12'
    $buttonFilter.Name = 'buttonFilter'
    $buttonFilter.Size = '62, 22'
    $buttonFilter.TabIndex = 12
    $buttonFilter.Text = 'Filter'
    $buttonFilter.UseVisualStyleBackColor = $True
    $buttonFilter.Visible = $False
    $buttonFilter.add_Click($buttonFilter_Click)
    #
    # textboxHighlight
    #
    $textboxHighlight.Anchor = 'Top, Right'
    $textboxHighlight.Location = '532, 33'
    $textboxHighlight.Name = 'textboxHighlight'
    $textboxHighlight.Size = '87, 20'
    $textboxHighlight.TabIndex = 10
    $textboxHighlight.add_TextChanged($textboxHighlight_TextChanged)
    #
    # buttonHighlight
    #
    $buttonHighlight.Anchor = 'Top, Right'
    $buttonHighlight.Enabled = $False
    $buttonHighlight.Location = '619, 32'
    $buttonHighlight.Name = 'buttonHighlight'
    $buttonHighlight.Size = '62, 22'
    $buttonHighlight.TabIndex = 9
    $buttonHighlight.Text = 'Highlight'
    $buttonHighlight.UseVisualStyleBackColor = $True
    $buttonHighlight.add_Click($buttonHighlight_Click)
    #
    # buttonMessage
    #
    $buttonMessage.Font = 'Microsoft Sans Serif, 8.25pt'
    #region Binary Data
    $buttonMessage.Image = [System.Convert]::FromBase64String('
iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABGdBTUEAALGPC/xhBQAABdNJREFU
SEvNVWtQlFUYPip52RIHSRQFpNLR+tEYPzT91Tgy1nShmxCWQKGlNAmWDM4UUpZpjU1qIl2mLNht
AWWBlpuCgMt9v4VVdFl2ueyysCCw7LKLK3h5v6ez6zYxTk1j04+emXfe73znPc9z3ss3H/tfQCKR
sPn+/r7Vf4yoqCgGIJTbyk7VV6xgv4R98KJv879Ap6GLOZ1Ond3hgr1PxZ8O652Va9dovl/GZs/w
Bf1bZLzhx6Ymhh93Ohy44bJhwumEw3ULtkEdxnpK0oHbLEcqZaGhob4T94Cqw7PZeMWqCFtvuavq
8igiv7PgfOc4plxj5HaOw5uRwzE1OTl5gJdP8tprsb6T/4BHFjM2dDqIjZxbnzrlUuGa8zo2flZG
LEFBC/c20lMnzaJUsJPD7hCvj9vE69cmUFhYhOjol1lYoI/k75D0DGPXzoXPGG3aWnnDLcDUrMDQ
sI0mnWPIVApYlvQrZsTlk3/yBTz+pZEO19gwOnETu5Pfxxcpa35yli9niZE+srtxJt2P2ctXrB43
Hxmzm37DhV0hsFiHYezuobySCgwOWCC67ShuvEJPfqgAi5XSA+9WYtWnl5GR9j6lf/wRprrTuuwV
q5cUfeLnY/UhbctM5lJtSHS7lDAUp+JsQoCo+TmN+odGqfx8rdiu66CSylqyDlppdHhQ5A2ndkOv
GPNVObGYXyhsX53YoKqnruojcI/lYbx2Q8zxXff52Dl2R81mGPq2t1OxC6UJwTi7PZA69R24pOuA
2WSic7X1kBeVoVFow6DVSgMDA7g6NEhuxwj6rVZUt3XRuG0Yha9KYFbGgCyZlw5tn+9j59j9ooSJ
liwDqhhqUxdT5cHXRX2PhTRtF0lWWCoqK6vJYjZTvVqgqroG0dLXRxaLRezv7yfrwADZHG5RlbmX
DCcCCI0MN3qyhEOJ077896Ik7Kb5pAE1/KNtuQ/9dUmkaR9CdkEpVA1NpOOZ5CvP4rSyAh16PVn6
zDCZesls5r6PZzFgJWeZP6ANAOpmYarrpPD59mkC73oEejINqOUCDQu4yFyaUK9HTV0TqlStJFOU
oKyqBvoOPQorqiinQIlmoY2MPQ5cNcqA1kWE1iX83CKg3g9uY6Zw8K3pAi/MZZPdJ+4I1C8AtfAD
Gh7cFoSSwgzKLxVQ16JFbZMWeoOJDN19UFaryXVlG9C+BGJbOI8P4QIPegUm9CeEA29OE9j5/Dzm
Nn7jFaAGfxJbHhRJs5REbRihI1h06V6iHw7tI2PRVtLIk0R12TGCKUKEbiVR+yoe9xCPX0bUEkie
Ejl1x4UDCQt87BzvcIEJw/E7PaibD2rmgQJP+WIYbqmDSfvjZsDEb2tcCXQ9RiMtO2GQRRA6+bp9
BahtOY8PBpoDAdUs2K8cEzLip03R28/NY86OYwZUewTuBzXyaWhdCHvJQugKdxC6NvAG8hJoHwZp
HyZcWoqb+leoXRqFW02LeSmDCeogoIk3+QKDrf2osD9umsCOZ+cx+5WjBpxnEFVzCM1zRLNsJQ3W
JxK0i0RRHUQkBHtNFIJFr1cH8hI9Sl2liTRWvFiExp885fUIjFz8WkjfNq0H8RsZG7t8xOjNgIsY
pBvpettGPlF+oPr5hAY+gj7zkvzh6+/nt56N0aZoMskj+P4seMo8oj3SmhYz8w75B6lp7OesDHb5
bIpMrHgApuIXADXvgacfnqmqZeTzd6//fF/D6LZmNUwFT/PnAFwqT/5O9n06S9v3IWNyuXx5jjQ3
7dSpH7Ors2NFsyIQPbmB6M4LuifryQ3glwtD1altdIojm3N6uFlRUZFEoVA8kX/6zObc08otsvzy
HTny0mRpXnkGt/2//FqSkS0v/YZblsf4Oisnt8zjPe+8MTze45NleWU75PlFW86cKdhcUFDwBOeV
sJCQELZ27Vq2adMm7w8+OjqaJSQksPj4eBYXF8f27NnDUlJS/tI8e3/EenxsbKyXIzIykq1bt46F
h4ez3wFCsmtsCMAhlwAAAABJRU5ErkJggg==')
    #endregion
    $buttonMessage.ImageAlign = 'TopCenter'
    $buttonMessage.Location = '368, 9'
    $buttonMessage.Name = 'buttonMessage'
    $buttonMessage.Size = '64, 45'
    $buttonMessage.TabIndex = 8
    $buttonMessage.Text = 'Message'
    $buttonMessage.TextAlign = 'BottomCenter'
    $buttonMessage.UseVisualStyleBackColor = $True
    $buttonMessage.add_Click($buttonMessage_Click)
    #
    # textboxComputerName
    #
    $textboxComputerName.Font = 'Microsoft Sans Serif, 12pt'
    $textboxComputerName.Location = '16, 27'
    $textboxComputerName.Name = 'textboxComputerName'
    $textboxComputerName.Size = '206, 26'
    $textboxComputerName.TabIndex = 0
    $textboxComputerName.Text = "$($Configuration.settings.computer)"
    #
    # labelComputerName
    #
    $labelComputerName.Font = 'Microsoft Sans Serif, 11.25pt, style=Bold'
    $labelComputerName.Location = '13, 9'
    $labelComputerName.Name = 'labelComputerName'
    $labelComputerName.Size = '177, 23'
    $labelComputerName.TabIndex = 5
    $labelComputerName.Text = 'Computer Name:'
    #
    # buttonGetTsSession
    #
    $buttonGetTsSession.ContextMenuStrip = $contextmenustripTSSession
    $buttonGetTsSession.Font = 'Microsoft Sans Serif, 8.25pt'
    #region Binary Data
    $buttonGetTsSession.Image = [System.Convert]::FromBase64String('
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
rkJggg==')
    #endregion
    $buttonGetTsSession.ImageAlign = 'TopCenter'
    $buttonGetTsSession.Location = '228, 9'
    $buttonGetTsSession.Name = 'buttonGetTsSession'
    $buttonGetTsSession.Size = '64, 45'
    $buttonGetTsSession.TabIndex = 1
    $buttonGetTsSession.Text = 'Sessions'
    $buttonGetTsSession.TextAlign = 'BottomCenter'
    $tooltip1.SetToolTip($buttonGetTsSession, 'Use Get-TSSession to get a list of sessions from a local or remote computers')
    $buttonGetTsSession.UseVisualStyleBackColor = $True
    $buttonGetTsSession.add_Click($buttonGetTsSession_Click)
    #
    # buttonProcess
    #
    $buttonProcess.Font = 'Microsoft Sans Serif, 8.25pt'
    #region Binary Data
    $buttonProcess.Image = [System.Convert]::FromBase64String('
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
RK5CYII=')
    #endregion
    $buttonProcess.ImageAlign = 'TopCenter'
    $buttonProcess.Location = '298, 9'
    $buttonProcess.Name = 'buttonProcess'
    $buttonProcess.Size = '64, 45'
    $buttonProcess.TabIndex = 4
    $buttonProcess.Text = 'Process'
    $buttonProcess.TextAlign = 'BottomCenter'
    $tooltip1.SetToolTip($buttonProcess, 'Use Get-TSProcess to get a list of session processes from a local or remote computers.')
    $buttonProcess.UseVisualStyleBackColor = $True
    $buttonProcess.add_Click($buttonProcess_Click)
    #
    # richtextboxStatus
    #
    $richtextboxStatus.Dock = 'Bottom'
    $richtextboxStatus.Font = 'Consolas, 8.25pt'
    $richtextboxStatus.Location = '0, 247'
    $richtextboxStatus.Name = 'richtextboxStatus'
    $richtextboxStatus.Size = '723, 71'
    $richtextboxStatus.TabIndex = 8
    $richtextboxStatus.Text = ''
    #
    # statusstrip1
    #
    [void]$statusstrip1.Items.Add($toolstripstatuslabel1)
    [void]$statusstrip1.Items.Add($toolstripstatuslabel2)
    [void]$statusstrip1.Items.Add($toolstripstatuslabel3)
    $statusstrip1.Location = '0, 318'
    $statusstrip1.Name = 'statusstrip1'
    $statusstrip1.RenderMode = 'Professional'
    $statusstrip1.Size = '723, 22'
    $statusstrip1.TabIndex = 12
    $statusstrip1.Text = 'statusstrip1'
    #
    # tooltip1
    #
    #
    # helpprovider1
    #
    #
    # contextmenustripTSSession
    #
    [void]$contextmenustripTSSession.Items.Add($sendTSSessionToolStripMenuItem)
    [void]$contextmenustripTSSession.Items.Add($disconnectTSSessionToolStripMenuItem)
    [void]$contextmenustripTSSession.Items.Add($stopTSSessionToolStripMenuItem)
    [void]$contextmenustripTSSession.Items.Add($toolstripseparator1)
    [void]$contextmenustripTSSession.Items.Add($remoteDesktopToolStripMenuItem)
    [void]$contextmenustripTSSession.Items.Add($remoteDesktopShadowIDToolStripMenuItem)
    [void]$contextmenustripTSSession.Items.Add($remoteDesktopShadowControlToolStripMenuItem1)
    [void]$contextmenustripTSSession.Items.Add($powerShellRemotingToolStripMenuItem)
    $contextmenustripTSSession.Name = 'contextmenustrip1'
    $contextmenustripTSSession.RenderMode = 'System'
    $contextmenustripTSSession.ShowImageMargin = $False
    $contextmenustripTSSession.Size = '233, 142'
    $contextmenustripTSSession.Text = 'TSSession'
    #
    # disconnectTSSessionToolStripMenuItem
    #
    $disconnectTSSessionToolStripMenuItem.Name = 'disconnectTSSessionToolStripMenuItem'
    $disconnectTSSessionToolStripMenuItem.Size = '232, 22'
    $disconnectTSSessionToolStripMenuItem.Text = 'Disconnect Session'
    $disconnectTSSessionToolStripMenuItem.add_Click($disconnectTSSessionToolStripMenuItem_Click)
    #
    # sendTSSessionToolStripMenuItem
    #
    $sendTSSessionToolStripMenuItem.Name = 'sendTSSessionToolStripMenuItem'
    $sendTSSessionToolStripMenuItem.Size = '232, 22'
    $sendTSSessionToolStripMenuItem.Text = 'Send Message'
    $sendTSSessionToolStripMenuItem.ToolTipText = 'Displays a message box to the selected session ID'
    $sendTSSessionToolStripMenuItem.add_Click($sendTSSessionToolStripMenuItem_Click)
    #
    # contextmenustripTSProcess
    #
    [void]$contextmenustripTSProcess.Items.Add($sendTSMessageToolStripMenuItem)
    [void]$contextmenustripTSProcess.Items.Add($stopTSProcessToolStripMenuItem)
    [void]$contextmenustripTSProcess.Items.Add($toolstripseparator2)
    [void]$contextmenustripTSProcess.Items.Add($remoteDesktopToolStripMenuItem1)
    [void]$contextmenustripTSProcess.Items.Add($remoteDesktopShadowToolStripMenuItem)
    [void]$contextmenustripTSProcess.Items.Add($remoteDesktopShadowControlToolStripMenuItem)
    [void]$contextmenustripTSProcess.Items.Add($powerShellRemotingToolStripMenuItem1)
    $contextmenustripTSProcess.Name = 'contextmenustripTSProcess'
    $contextmenustripTSProcess.RenderMode = 'System'
    $contextmenustripTSProcess.ShowImageMargin = $False
    $contextmenustripTSProcess.Size = '233, 120'
    #
    # sendTSMessageToolStripMenuItem
    #
    $sendTSMessageToolStripMenuItem.Name = 'sendTSMessageToolStripMenuItem'
    $sendTSMessageToolStripMenuItem.Size = '232, 22'
    $sendTSMessageToolStripMenuItem.Text = 'Send Message'
    $sendTSMessageToolStripMenuItem.ToolTipText = 'Displays a message box to the selected session ID'
    $sendTSMessageToolStripMenuItem.add_Click($sendTSSessionToolStripMenuItem_Click)
    #
    # stopTSProcessToolStripMenuItem
    #
    $stopTSProcessToolStripMenuItem.Name = 'stopTSProcessToolStripMenuItem'
    $stopTSProcessToolStripMenuItem.Size = '232, 22'
    $stopTSProcessToolStripMenuItem.Text = 'Stop Process'
    $stopTSProcessToolStripMenuItem.ToolTipText = 'Use Stop-TSProcess to terminate one or more processes'
    $stopTSProcessToolStripMenuItem.add_Click($stopTSProcessToolStripMenuItem_Click)
    #
    # stopTSSessionToolStripMenuItem
    #
    $stopTSSessionToolStripMenuItem.Name = 'stopTSSessionToolStripMenuItem'
    $stopTSSessionToolStripMenuItem.Size = '232, 22'
    $stopTSSessionToolStripMenuItem.Text = 'Stop Session'
    $stopTSSessionToolStripMenuItem.add_Click($stopTSSessionToolStripMenuItem_Click)
    #
    # toolstripstatuslabel1
    #
    $toolstripstatuslabel1.Font = 'Segoe UI, 9pt, style=Bold'
    $toolstripstatuslabel1.IsLink = $True
    $toolstripstatuslabel1.Name = 'toolstripstatuslabel1'
    $toolstripstatuslabel1.Size = '225, 17'
    $toolstripstatuslabel1.Text = "$($Configuration.author.website.name)"
    $toolstripstatuslabel1.add_Click($toolstripstatuslabel1_Click)
    #
    # toolstripstatuslabel2
    #
    $toolstripstatuslabel2.Enabled = $False
    $toolstripstatuslabel2.Font = 'Segoe UI, 9pt, style=Bold'
    $toolstripstatuslabel2.Name = 'toolstripstatuslabel2'
    $toolstripstatuslabel2.Size = '178, 17'
    $toolstripstatuslabel2.Text = "$($Configuration.author.name)"
    #
    # toolstripseparator1
    #
    $toolstripseparator1.Name = 'toolstripseparator1'
    $toolstripseparator1.Size = '229, 6'
    #
    # remoteDesktopToolStripMenuItem
    #
    $remoteDesktopToolStripMenuItem.Name = 'remoteDesktopToolStripMenuItem'
    $remoteDesktopToolStripMenuItem.Size = '232, 22'
    $remoteDesktopToolStripMenuItem.Text = 'Remote Desktop'
    $remoteDesktopToolStripMenuItem.add_Click($remoteDesktopToolStripMenuItem_Click)
    #
    # powerShellRemotingToolStripMenuItem
    #
    $powerShellRemotingToolStripMenuItem.Name = 'powerShellRemotingToolStripMenuItem'
    $powerShellRemotingToolStripMenuItem.Size = '232, 22'
    $powerShellRemotingToolStripMenuItem.Text = 'PowerShell Remoting'
    $powerShellRemotingToolStripMenuItem.add_Click($powerShellRemotingToolStripMenuItem_Click)
    #
    # toolstripseparator2
    #
    $toolstripseparator2.Name = 'toolstripseparator2'
    $toolstripseparator2.Size = '229, 6'
    #
    # remoteDesktopToolStripMenuItem1
    #
    $remoteDesktopToolStripMenuItem1.Name = 'remoteDesktopToolStripMenuItem1'
    $remoteDesktopToolStripMenuItem1.Size = '232, 22'
    $remoteDesktopToolStripMenuItem1.Text = 'Remote Desktop'
    $remoteDesktopToolStripMenuItem1.add_Click($remoteDesktopToolStripMenuItem_Click)
    #
    # powerShellRemotingToolStripMenuItem1
    #
    $powerShellRemotingToolStripMenuItem1.Name = 'powerShellRemotingToolStripMenuItem1'
    $powerShellRemotingToolStripMenuItem1.Size = '232, 22'
    $powerShellRemotingToolStripMenuItem1.Text = 'PowerShell Remoting'
    $powerShellRemotingToolStripMenuItem1.add_Click($powerShellRemotingToolStripMenuItem_Click)
    #
    # remoteDesktopShadowIDToolStripMenuItem
    #
    $remoteDesktopShadowIDToolStripMenuItem.Name = 'remoteDesktopShadowIDToolStripMenuItem'
    $remoteDesktopShadowIDToolStripMenuItem.Size = '232, 22'
    $remoteDesktopShadowIDToolStripMenuItem.Text = 'Remote Desktop (Shadow View)'
    $remoteDesktopShadowIDToolStripMenuItem.add_Click($remoteDesktopShadowIDToolStripMenuItem_Click)
    #
    # remoteDesktopShadowToolStripMenuItem
    #
    $remoteDesktopShadowToolStripMenuItem.Name = 'remoteDesktopShadowToolStripMenuItem'
    $remoteDesktopShadowToolStripMenuItem.Size = '232, 22'
    $remoteDesktopShadowToolStripMenuItem.Text = 'Remote Desktop (Shadow View)'
    $remoteDesktopShadowToolStripMenuItem.add_Click($remoteDesktopShadowIDToolStripMenuItem_Click)
    #
    # remoteDesktopShadowControlToolStripMenuItem
    #
    $remoteDesktopShadowControlToolStripMenuItem.Name = 'remoteDesktopShadowControlToolStripMenuItem'
    $remoteDesktopShadowControlToolStripMenuItem.Size = '232, 22'
    $remoteDesktopShadowControlToolStripMenuItem.Text = 'Remote Desktop (Shadow Control)'
    $remoteDesktopShadowControlToolStripMenuItem.add_Click($remoteDesktopShadowControlToolStripMenuItem_Click)
    #
    # remoteDesktopShadowControlToolStripMenuItem1
    #
    $remoteDesktopShadowControlToolStripMenuItem1.Name = 'remoteDesktopShadowControlToolStripMenuItem1'
    $remoteDesktopShadowControlToolStripMenuItem1.Size = '232, 22'
    $remoteDesktopShadowControlToolStripMenuItem1.Text = 'Remote Desktop (Shadow Control)'
    $remoteDesktopShadowControlToolStripMenuItem1.add_Click($remoteDesktopShadowControlToolStripMenuItem_Click)
    #
    # toolstripstatuslabel3
    #
    $toolstripstatuslabel3.IsLink = $True
    $toolstripstatuslabel3.Name = 'toolstripstatuslabel3'
    $toolstripstatuslabel3.Size = '64, 17'
    $toolstripstatuslabel3.Text = 'Contribute'
    $toolstripstatuslabel3.add_Click($toolstripstatuslabel3_Click)
    $contextmenustripTSProcess.ResumeLayout()
    $contextmenustripTSSession.ResumeLayout()
    $statusstrip1.ResumeLayout()
    $groupbox1.ResumeLayout()
    $MainForm.ResumeLayout()
    #endregion Generated Form Code

    #----------------------------------------------

    #Save the initial state of the form
    $InitialFormWindowState = $MainForm.WindowState
    #Init the OnLoad event to correct the initial state of the form
    $MainForm.add_Load($Form_StateCorrection_Load)
    #Clean up the control events
    $MainForm.add_FormClosed($Form_Cleanup_FormClosed)
    #Store the control values when form is closing
    $MainForm.add_Closing($Form_StoreValues_Closing)
    #Show the Form
    return $MainForm.ShowDialog()

}
#endregion Source: MainForm.psf

#region Source: Globals.ps1
    Write-Verbose -Message "[$ScriptName] Globals.ps1 processing..."
    #--------------------------------------------
    # Declare Global Variables and Functions here
    #--------------------------------------------

    #Location of the script
    function Get-ScriptDirectory
    {
        if($null -ne $hostinvocation)
        {
            Split-Path -path $hostinvocation.MyCommand.path
        }
        else
        {
            Split-Path -path $script:MyInvocation.MyCommand.Path
        }
    }

    [System.string]$ScriptDirectory = Get-ScriptDirectory


    # Load Configuration
    $configurationpath = Join-Path -Path $ScriptDirectory -ChildPath 'config.psd1'
    if(-not(Test-Path $configurationpath)){Write-Warning -message "can't retrieve the configuration file 'config.psd1'"}
    $Configuration= Import-LocalizedData -BaseDirectory (Get-ScriptDirectory) -FileName "config.psd1"
    if (-not($Configuration.settings.computer)){ $Configuration.settings.computer = $env:computername}


    # PSTerminalServices Module Requirements
    # Get the path of the DLL file: Cassia.dll (.net library)
    $CassiaPath = Join-Path -Path (Get-ScriptDirectory) -ChildPath "Cassia.dll"
    # Load the DLL
    if (-not(Test-Path $CassiaPath)){Write-Warning -Message "The file Cassia.dll is missing. The Script can't continue without this file";exit}
    [Reflection.Assembly]::LoadFile($CassiaPath) | Out-Null

    # Import all the helper functions
    $functionpath = Join-Path -Path $ScriptDirectory -child "functions"
    $script:Server = 'localhost'
    foreach ($file in (Get-ChildItem -path $functionpath))
    {
        . $file.fullname
    }



    #
    ## From WinFormPS
    #function Append-RichtextboxStatus{
    #    PARAM(
    #    [Parameter(Mandatory=$true)]
    #    [string]$Message,
    #    [string]$MessageColor = "DarkGreen",
    #    [string]$DateTimeColor="Black",
    #    [string]$Source,
    #    [string]$SourceColor="Gray",
    #    [string]$ComputerName,
    #    [String]$ComputerNameColor= "Blue")
    #
    #    $SortableTime = get-date -Format "yyyy-MM-dd HH:mm:ss"
    #    $richtextboxStatus.SelectionColor = $DateTimeColor
    #    $richtextboxStatus.AppendText("[$SortableTime] ")
    #
    #    IF ($PSBoundParameters['ComputerName']){
    #        $richtextboxStatus.SelectionColor = $ComputerNameColor
    #        $richtextboxStatus.AppendText(("$ComputerName ").ToUpper())
    #    }
    #
    #    IF ($PSBoundParameters['Source']){
    #        $richtextboxStatus.SelectionColor = $SourceColor
    #        $richtextboxStatus.AppendText("$Source ")
    #    }
    #
    #    $richtextboxStatus.SelectionColor = $MessageColor
    #    $richtextboxStatus.AppendText("$Message`r")
    #    $richtextboxStatus.Refresh()
    #    $richtextboxStatus.ScrollToCaret()
    #
    #    Write-Verbose -Message "$SortableTime $Message"
    #}
    #
    #function Set-DataGridView
    #{
    #    <#
    #        .SYNOPSIS
    #            This function helps you edit the datagridview control
    #
    #        .DESCRIPTION
    #            This function helps you edit the datagridview control
    #
    #        .EXAMPLE
    #            Set-DataGridView -DataGridView $datagridview1 -ProperFormat -FontFamily $listboxFontFamily.Text -FontSize $listboxFontSize.Text
    #
    #        .EXAMPLE
    #            Set-DataGridView -DataGridView $datagridview1 -AlternativeRowColor -BackColor 'AliceBlue' -ForeColor 'Black'
    #
    #        .EXAMPLE
    #            Set-DataGridViewRowHeader -DataGridView $datagridview1 -HideRowHeader
    #
    #        .EXAMPLE
    #            Set-DataGridViewRowHeader -DataGridView $datagridview1 -ShowRowHeader
    #
    #    #>
    #
    #    [CmdletBinding()]
    #    PARAM (
    #        [ValidateNotNull()]
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView,
    #
    #        [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
    #        [Switch]$AlternativeRowColor,
    #
    #        [Parameter(ParameterSetName = "DefaultRowColor")]
    #        [Switch]$DefaultRowColor,
    #
    #        [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
    #        [Parameter(ParameterSetName = "DefaultRowColor")]
    #        [System.Drawing.Color]$ForeColor,
    #
    #        [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
    #        [Parameter(ParameterSetName = "DefaultRowColor")]
    #        [System.Drawing.Color]$BackColor,
    #
    #        [Parameter(Mandatory = $true, ParameterSetName = "Proper")]
    #        [Switch]$ProperFormat,
    #
    #        [Parameter(ParameterSetName = "Proper")]
    #        [String]$FontFamily = "Consolas",
    #
    #        [Parameter(ParameterSetName = "Proper")]
    #        [Int]$FontSize = 10,
    #
    #        [Parameter(ParameterSetName = "HideRowHeader")]
    #        [Switch]$HideRowHeader,
    #        [Parameter(ParameterSetName = "ShowRowHeader")]
    #        [Switch]$ShowRowHeader
    #    )
    #    PROCESS
    #    {
    #        if ($psboundparameters['AlternativeRowColor'])
    #        {
    #            $DataGridView.AlternatingRowsDefaultCellStyle.ForeColor = $ForeColor
    #            $DataGridView.AlternatingRowsDefaultCellStyle.BackColor = $BackColor
    #        }
    #
    #        if ($psboundparameters['DefaultRowColor'])
    #        {
    #            $DataGridView.RowsDefaultCellStyle.ForeColor = $ForeColor
    #            $DataGridView.RowsDefaultCellStyle.BackColor = $BackColor
    #        }
    #
    #
    #        if ($psboundparameters['ProperFormat'])
    #        {
    #            #$Font = New-Object -TypeName System.Drawing.Font -ArgumentList "Segoi UI", 10
    #            $Font = New-Object -TypeName System.Drawing.Font -ArgumentList $FontFamily, $FontSize
    #
    #            #[System.Drawing.FontStyle]::Bold
    #
    #            $DataGridView.ColumnHeadersBorderStyle = 'Raised'
    #            $DataGridView.BorderStyle = 'Fixed3D'
    #            $DataGridView.SelectionMode = 'FullRowSelect'
    #            $DataGridView.AllowUserToResizeRows = $false
    #            $datagridview.DefaultCellStyle.font = $Font
    #        }
    #
    #        if ($psboundparameters['HideRowHeader'])
    #        {
    #            $DataGridView.RowHeadersVisible = $false
    #        }
    #        if ($psboundparameters['ShowRowHeader'])
    #        {
    #            $DataGridView.RowHeadersVisible = $true
    #        }
    #    }
    #
    #}#Set-DataGridView
    #
    #function Reset-DataGridViewFormat
    #{
    #<#
    #    .SYNOPSIS
    #        The Reset-DataGridViewFormat function will reset the format of a datagridview control
    #
    #    .DESCRIPTION
    #        The Reset-DataGridViewFormat function will reset the format of a datagridview control
    #
    #    .PARAMETER DataGridView
    #        Specifies the DataGridView Control.
    #
    #    .EXAMPLE
    #        PS C:\> Reset-DataGridViewFormat -DataGridView $DataGridViewObj
    ##>
    #    [CmdletBinding()]
    #    PARAM (
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView)
    #    PROCESS
    #    {
    #        $DataSource = $DataGridView.DataSource
    #        $DataGridView.DataSource = $null
    #        $DataGridView.DataSource = $DataSource
    #
    #        #$DataGridView.RowsDefaultCellStyle.BackColor = 'White'
    #        #$DataGridView.RowsDefaultCellStyle.ForeColor = 'Black'
    #        $DataGridView.RowsDefaultCellStyle = $null
    #        $DataGridView.AlternatingRowsDefaultCellStyle = $null
    #    }
    #}#Reset-DataGridViewFormat
    #
    #function Find-DataGridViewValue
    #{
    #<#
    #    .SYNOPSIS
    #        The Find-DataGridViewValue function helps you to find a specific value and select the cell, row or to set a fore and back color.
    #
    #    .DESCRIPTION
    #        The Find-DataGridViewValue function helps you to find a specific value and select the cell, row or to set a fore and back color.
    #
    #    .PARAMETER DataGridView
    #        Specifies the DataGridView Control to use
    #
    #    .PARAMETER RowBackColor
    #        Specifies the back color of the row to use
    #
    #    .PARAMETER RowForeColor
    #        Specifies the fore color of the row to use
    #
    #    .PARAMETER SelectCell
    #        Specifies to select only the cell when the value is found
    #
    #    .PARAMETER SelectRow
    #        Specifies to select the entire row when the value is found
    #
    #    .PARAMETER Value
    #        Specifies the value to search
    #
    #    .EXAMPLE
    #        PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text
    #
    #        This will find the value and select the cell(s)
    #
    #    .EXAMPLE
    #        PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text -RowForeColor 'Red' -RowBackColor 'Black'
    #
    #        This will find the value and color the fore and back of the row
    #    .EXAMPLE
    #        PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text -SelectRow
    #
    #        This will find the value and select the entire row
    #
    ##>
    #    [CmdletBinding(DefaultParameterSetName = "Cell")]
    #    PARAM (
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView,
    #
    #    [ValidateNotNull()]
    #    [Parameter(Mandatory = $true)]
    #        [String]$Value,
    #        [Parameter(ParameterSetName = "Cell")]
    #        [Switch]$SelectCell,
    #
    #        [Parameter(ParameterSetName = "Row")]
    #        [Switch]$SelectRow,
    #
    #        #[Parameter(ParameterSetName = "Column")]
    #        #[Switch]$SelectColumn,
    #        [Parameter(ParameterSetName = "RowColor")]
    #        [system.Drawing.Color]$RowForeColor,
    #        [Parameter(ParameterSetName = "RowColor")]
    #        [system.Drawing.Color]$RowBackColor
    #    )
    #
    #    PROCESS
    #    {
    #        $DataGridView.ClearSelection()
    #
    #        FOR ([int]$i = 0; $i -lt $DataGridView.RowCount; $i++)
    #        {
    #            FOR ([int] $j = 0; $j -lt $DataGridView.ColumnCount; $j++)
    #            {
    #                $CurrentCell = $dataGridView.Rows[$i].Cells[$j]
    #
    #                #if ((-not $CurrentCell.Value.Equals([DBNull]::Value)) -and ($CurrentCell.Value.ToString() -like "*$Value*"))
    #                if ($CurrentCell.Value.ToString() -match $Value)
    #                {
    #
    #                    # Row Selection
    #                    IF ($PSBoundParameters['SelectRow'])
    #                    {
    #                        $dataGridView.Rows[$i].Selected = $true
    #                    }
    #
    #                    <#
    #                    # Column Selection
    #                    IF ($PSBoundParameters['SelectColumn'])
    #                    {
    #                        #$DataGridView.Columns[$($CurrentCell.ColumnIndex)].Selected = $true
    #                        #$DataGridView.Columns[$j].Selected = $true
    #                        #$CurrentCell.DataGridView.Columns[$j].Selected = $true
    #                    }
    #                    #>
    #
    #                    # Row Fore Color
    #                    IF ($PSBoundParameters['RowForeColor'])
    #                    {
    #                        $dataGridView.Rows[$i].DefaultCellStyle.ForeColor = $RowForeColor
    #                    }
    #                    # Row Back Color
    #                    IF ($PSBoundParameters['RowBackColor'])
    #                    {
    #                        $dataGridView.Rows[$i].DefaultCellStyle.BackColor = $RowBackColor
    #                    }
    #
    #                    # Cell Selection
    #                    ELSEIF (-not ($PSBoundParameters['SelectRow']) -and -not ($PSBoundParameters['SelectColumn']))
    #                    {
    #                        $CurrentCell.Selected = $true
    #                    }
    #                }#IF not empty and contains value
    #            }#For Each column
    #        }#For Each Row
    #    }#PROCESS
    #}#Find-DataGridViewValue
    #
    #function Set-DataGridViewFilter
    #{
    #<#
    #    .SYNOPSIS
    #        The function Set-DataGridViewFilter helps to only show specific entries with a specific value
    #
    #    .DESCRIPTION
    #        The function Set-DataGridViewFilter helps to only show specific entries with a specific value.
    #        The data needs to be in a DataTable Object. You can use ConvertTo-DataTable to convert your
    #        PowerShell object into a DataTable object.
    #
    #    .PARAMETER AllColumns
    #        Specifies to search all the column
    #
    #    .PARAMETER ColumnName
    #        Specifies to search in a specific column name
    #
    #    .PARAMETER DataGridView
    #        Specifies the DataGridView control where the data will be filtered
    #
    #    .PARAMETER DataTable
    #        Specifies the DataTable object that is most likely the original source of the DataGridView
    #
    #    .PARAMETER Filter
    #        Specifies the string to search
    #
    #    .EXAMPLE
    #        PS C:\> Set-DataGridViewFilter -DataGridView $datagridview1 -DataTable $ProcessesDT -AllColumns -Filter $textbox1.Text
    #
    #    .EXAMPLE
    #        PS C:\> Set-DataGridViewFilter -DataGridView $datagridview1 -DataTable $ProcessesDT -ColumnName "Name" -Filter $textbox1.Text
    #
    ##>
    #    PARAM (
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView,
    #        [Parameter(Mandatory = $true)]
    #        [System.Data.DataTable]$DataTable,
    #        [Parameter(Mandatory = $true)]
    #        [String]$Filter,
    #
    #        [Parameter(Mandatory = $true, ParameterSetName = "OneColumn")]
    #        [String]$ColumnName,
    #        [Parameter(Mandatory = $true, ParameterSetName = "AllColumns")]
    #        [Switch]$AllColumns
    #    )
    #    PROCESS
    #    {
    #        $Filter = $Filter.ToString()
    #
    #        IF ($PSBoundParameters['AllColumns'])
    #        {
    #            FOREACH ($Column in $DataTable.Columns)
    #            {
    #                #$RowFilter += "Convert("+$($Column.ColumnName)+",'system.string') Like '%"{1}%' OR " -f $Column.ColumnName, $Filter
    #                $RowFilter += "Convert($($Column.ColumnName),'system.string') Like '%$Filter%' OR "
    #            }
    #
    #            # Remove the last 'OR'
    #            $RowFilter = $RowFilter -replace " OR $", ''
    #
    #            #Append-RichtextboxStatus -Message $RowFilter
    #        }
    #        IF ($PSBoundParameters['ColumnName'])
    #        {
    #            $RowFilter = "$ColumnName LIKE '%$Filter%'"
    #        }
    #
    #        $DataTable.defaultview.rowfilter = $RowFilter
    #        Load-DataGridView -DataGridView $DataGridView -Item $DataTable
    #    }
    #    END { Remove-Variable -Name $RowFilter -ErrorAction 'SilentlyContinue' | Out-Null }
    #}#Set-DataGridViewFilter
    #
    #function ConvertTo-DataTable
    #{
    #    <#
    #        .SYNOPSIS
    #            Converts objects into a DataTable.
    #
    #        .DESCRIPTION
    #            Converts objects into a DataTable, which are used for DataBinding.
    #
    #        .PARAMETER  InputObject
    #            The input to convert into a DataTable.
    #
    #        .PARAMETER  Table
    #            The DataTable you wish to load the input into.
    #
    #        .PARAMETER RetainColumns
    #            This switch tells the function to keep the DataTable's existing columns.
    #
    #        .PARAMETER FilterWMIProperties
    #            This switch removes WMI properties that start with an underline.
    #
    #        .EXAMPLE
    #            $DataTable = ConvertTo-DataTable -InputObject (Get-Process)
    #
    #        .NOTES
    #            SAPIEN Technologies, Inc.
    #            http://www.sapien.com/
    #
    #            VERSION HISTORY
    #            1.0 ????/??/?? From Sapien.com Version
    #    #>
    #    [CmdletBinding()]
    #    [OutputType([System.Data.DataTable])]
    #    param (
    #        [ValidateNotNull()]
    #        $InputObject,
    #        [ValidateNotNull()]
    #        [System.Data.DataTable]$Table,
    #        [switch]$RetainColumns,
    #        [switch]$FilterWMIProperties
    #    )
    #
    #    if ($Table -eq $null)
    #    {
    #        $Table = New-Object System.Data.DataTable
    #    }
    #
    #    if ($InputObject -is [System.Data.DataTable])
    #    {
    #        $Table = $InputObject
    #    }
    #    else
    #    {
    #        if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
    #        {
    #            #Clear out the Table Contents
    #            $Table.Clear()
    #
    #            if ($InputObject -eq $null) { return } #Empty Data
    #
    #            $object = $null
    #
    #            #find the first non null value
    #            foreach ($item in $InputObject)
    #            {
    #                if ($item -ne $null)
    #                {
    #                    $object = $item
    #                    break
    #                }
    #            }
    #
    #            if ($object -eq $null) { return } #All null then empty
    #
    #            #COLUMN
    #            #Get all the properties in order to create the columns
    #            foreach ($prop in $object.PSObject.Get_Properties())
    #            {
    #                if (-not $FilterWMIProperties -or -not $prop.Name.StartsWith('__'))#filter out WMI properties
    #                {
    #                    #Get the type from the Definition string
    #                    $type = $null
    #
    #                    if ($prop.Value -ne $null)
    #                    {
    #                        try { $type = $prop.Value.GetType() }
    #                        catch { Write-Verbose -Message "Can't find type of $prop" }
    #                    }
    #
    #                    if ($type -ne $null) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
    #                    {
    #                        Write-Verbose -Message "Creating Column: $($Prop.name) (Type: $type)"
    #                        [void]$table.Columns.Add($prop.Name, $type)
    #                    }
    #                    else #Type info not found
    #                    {
    #                        #if ($prop.name -eq "" -or $prop.name -eq $null) { [void]$table.Columns.Add([DBNull]::Value) }
    #                        [void]$table.Columns.Add($prop.Name)
    #                    }
    #                }
    #            }
    #
    #            if ($object -is [System.Data.DataRow])
    #            {
    #                foreach ($item in $InputObject)
    #                {
    #                    $Table.Rows.Add($item)
    #                }
    #                return @(, $Table)
    #            }
    #        }
    #        else
    #        {
    #            $Table.Rows.Clear()
    #        }
    #
    #        #Rows Work
    #        foreach ($item in $InputObject)
    #        {
    #            # Create a new row object
    #            $row = $table.NewRow()
    #
    #            if ($item)
    #            {
    #                foreach ($prop in $item.PSObject.Get_Properties())
    #                {
    #                    #Find the appropriate column to put the value
    #                    if ($table.Columns.Contains($prop.Name))
    #                    {
    #                        if ($prop.value -eq $null) { $prop.value = [DBNull]::Value }
    #                        $row.Item($prop.Name) = $prop.Value
    #                    }
    #                }
    #            }
    #            [void]$table.Rows.Add($row)
    #        }
    #    }
    #
    #    return @(, $Table)
    #}
    #
    #function Load-DataGridView
    #{
    #    <#
    #    .SYNOPSIS
    #        This functions helps you load items into a DataGridView.
    #
    #    .DESCRIPTION
    #        Use this function to dynamically load items into the DataGridView control.
    #
    #    .PARAMETER  DataGridView
    #        The ComboBox control you want to add items to.
    #
    #    .PARAMETER  Item
    #        The object or objects you wish to load into the ComboBox's items collection.
    #
    #    .PARAMETER  DataMember
    #        Sets the name of the list or table in the data source for which the DataGridView is displaying data.
    #    #>
    #
    #    [CmdletBinding()]
    #    Param (
    #        [ValidateNotNull()]
    #        [Parameter(Mandatory=$true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView,
    #        [ValidateNotNull()]
    #        [Parameter(Mandatory=$true)]
    #        $Item,
    #        [Parameter(Mandatory=$false)]
    #        [string]$DataMember
    #    )
    #    $DataGridView.SuspendLayout()
    #    $DataGridView.DataMember = $DataMember
    #
    #    if ($Item -is [System.ComponentModel.IListSource]`
    #    -or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView] )
    #    {
    #        $DataGridView.DataSource = $Item
    #    }
    #    else
    #    {
    #        $array = New-Object System.Collections.ArrayList
    #
    #        if ($Item -is [System.Collections.IList])
    #        {
    #            $array.AddRange($Item)
    #        }
    #        else
    #        {
    #            $array.Add($Item)
    #        }
    #        $DataGridView.DataSource = $array
    #    }
    #
    #    $DataGridView.ResumeLayout()
    #}
    #
    #function Set-TextBox
    #{
    #    [CmdletBinding()]
    #    PARAM (
    #        [System.Windows.Forms.TextBox]$TextBox,
    #        [System.Drawing.Color]$BackColor
    #    )
    #    BEGIN { }
    #    PROCESS
    #    {
    #        TRY
    #        {
    #            $TextBox.BackColor = $BackColor
    #        }
    #        CATCH { }
    #    }
    #}
    #
    #function Disable-Button
    #{
    #<#
    #.SYNOPSIS
    #    This function will disable a button control
    #.EXAMPLE
    #    Disable-Button -Button $Button
    ##>
    #    [CmdletBinding()]
    #    PARAM (
    #        [ValidateNotNull()]
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.Button[]]$Button
    #    )
    #    BEGIN
    #    {
    #        Add-Type -AssemblyName System.Windows.Forms
    #    }
    #    PROCESS
    #    {
    #        foreach ($ButtonObject in $Button)
    #        {
    #            $ButtonObject.Enabled = $false
    #        }
    #
    #    }
    #}#Disable-Button
    #
    #function Reset-TextBox
    #{
    #    [CmdletBinding()]
    #    PARAM (
    #        [System.Windows.Forms.TextBox]$TextBox,
    #        [System.Drawing.Color]$BackColor = "White",
    #        [System.Drawing.Color]$ForeColor = "Black"
    #    )
    #    BEGIN { }
    #    PROCESS
    #    {
    #        TRY
    #        {
    #            $TextBox.Text = ""
    #            $TextBox.BackColor = $BackColor
    #            $TextBox.ForeColor = $ForeColor
    #        }
    #        CATCH { }
    #    }
    #}
    #
    #function Enable-Button
    #{
    #<#
    #.SYNOPSIS
    #    This function will enable a button control
    #.EXAMPLE
    #    Enable-Button -Button $Button
    ##>
    #    [CmdletBinding()]
    #    PARAM (
    #        [ValidateNotNull()]
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.Button[]]$Button
    #    )
    #    BEGIN
    #    {
    #        Add-Type -AssemblyName System.Windows.Forms
    #    }
    #    PROCESS
    #    {
    #        foreach ($ButtonObject in $Button)
    #        {
    #            $ButtonObject.Enabled = $true
    #        }
    #    }
    #}#Enable-Button
    #
    #function Clear-DataGridViewSelection
    #{
    #    PARAM (
    #        [Parameter(Mandatory = $true)]
    #        [System.Windows.Forms.DataGridView]$DataGridView
    #    )
    #    $DataGridView.ClearSelection()
    #}
    #
    #function New-MessageBox
    #{
    #<#
    #    .SYNOPSIS
    #        The New-MessageBox functio will show a message box to the user
    #
    #    .DESCRIPTION
    #        The New-MessageBox functio will show a message box to the user
    #
    #    .PARAMETER Message
    #        Specifies the message to show
    #
    #    .PARAMETER Title
    #        Specifies the title of the message box
    #
    #    .PARAMETER Buttons
    #        Specifies which button to add. Just press tab to see the choices
    #
    #    .PARAMETER Icon
    #        Specifies the icon to show. Just press tab to see the choices
    #
    #    .EXAMPLE
    #        PS C:\> New-MessageBox -Message "Hello World" -Title "First Message" -Buttons "RetryCancel" -Icon "Asterix"
    #
    ##>
    #    [CmdletBinding()]
    #    PARAM (
    #
    #        [String]$Message,
    #        [String]$Title,
    #        [System.Windows.Forms.MessageBoxButtons]$Buttons = "OK",
    #        [System.Windows.Forms.MessageBoxIcon]$Icon = "None"
    #    )
    #    BEGIN
    #    {
    #        Add-Type -AssemblyName System.Windows.Forms
    #    }
    #    PROCESS
    #    {
    #        [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
    #    }
    #}#New-MessageBox
    #
    ## PSTerminalServices Module by Shay Levy
    #$script:Server='localhost'
    #function Get-TSSession
    #{
    #    <#
    #    .SYNOPSIS
    #        Lists the sessions on a given terminal server.
    #
    #    .DESCRIPTION
    #        Use Get-TSSession to get a list of sessions from a local or remote computers.
    #        Note that Get-TSSession is using Aliased properties to display the output on the console (IPAddress and State), these attributes
    #        are not the same as the original attributes (ClientIPAddress and ConnectionState).
    #        This is important when you want to use the -Filter parameter which requires the latter.
    #        To see all aliassed properties and their corresponding properties (Definition column), pipe the result to Get-Member:
    #
    #        PS > Get-TSSession | Get-Member -MemberType AliasProperty
    #
    #           TypeName: Cassia.Impl.TerminalServicesSession
    #
    #        Name      MemberType    Definition
    #        ----      ----------    ----------
    #        (...)
    #        IPAddress AliasProperty IPAddress = ClientIPAddress
    #        State     AliasProperty State = ConnectionState
    #
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Id
    #        Specifies the session Id number.
    #
    #    .PARAMETER InputObject
    #           Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.
    #
    #    .PARAMETER Filter
    #           Specifies a filter based on the session properties. The syntax of the filter, including the use of
    #           wildcards and depends on the properties of the session. Internally, The Filter parameter uses client side
    #           filtering using the Where-Object cmdlet, objects are filtered after they are retrieved.
    #
    #    .PARAMETER State
    #        The connection state of the session. Use this parameter to get sessions of a specific state. Valid values are:
    #
    #        Value         Description
    #        -----         -----------
    #        Active         A user is logged on to the session.
    #        ConnectQuery The session is in the process of connecting to a client.
    #        Connected     A client is connected to the session).
    #        Disconnected The session is active, but the client has disconnected from it.
    #        Down         The session is down due to an error.
    #        Idle         The session is waiting for a client to connect.
    #        Initializing The session is initializing.
    #        Listening      The session is listening for connections.
    #        Reset         The session is being reset.
    #        Shadowing     This session is shadowing another session.
    #
    #    .PARAMETER ClientName
    #        The name of the machine last connected to a session.
    #        Use this parameter to get sessions made from a specific computer name. Wildcrads are permitted.
    #
    #    .PARAMETER UserName
    #        Use this parameter to get sessions made by a specific user name. Wildcrads are permitted.
    #
    #    .EXAMPLE
    #        Get-TSSession
    #
    #        Description
    #        -----------
    #        Gets all the sessions from the local computer.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -State Disconnected
    #
    #        Description
    #        -----------
    #        Gets all the disconnected sessions from the remote computer 'comp1'.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -Filter {$_.ClientIPAddress -like '10*' -AND $_.ConnectionState -eq 'Active'}
    #
    #        Description
    #        -----------
    #        Gets all Active sessions from remote computer 'comp1', made from ip addresses that starts with '10'.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -UserName a*
    #
    #        Description
    #        -----------
    #        Gets all sessions from remote computer 'comp1' made by users with name starts with the letter 'a'.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -ClientName s*
    #
    #        Description
    #        -----------
    #        Gets all sessions from remote computer 'comp1' made from a computers names that starts with the letter 's'.
    #
    #    .OUTPUTS
    #        Cassia.Impl.TerminalServicesSession
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Stop-TSSession
    #        Disconnect-TSSession
    #        Send-TSMessage
    #    #>
    #
    #
    #    [OutputType('Cassia.Impl.TerminalServicesSession')]
    #    [CmdletBinding(DefaultParameterSetName='Session')]
    #
    #    Param(
    #
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName,
    #
    #        [Parameter(
    #            Position=0,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Session'
    #        )]
    #        [Alias('SessionID')]
    #        [ValidateRange(0,65536)]
    #        [System.Int32]$Id=-1,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesSession]$InputObject,
    #
    #        [Parameter(
    #            Mandatory=$true,
    #            ParameterSetName='Filter'
    #        )]
    #        [ScriptBlock]$Filter,
    #
    #        [Parameter()]
    #        [ValidateSet('Active','Connected','ConnectQuery','Shadowing','Disconnected','Idle','Listening','Reset','Down','Initializing')]
    #        [Alias('ConnectionState')]
    #        [System.String]$State='*',
    #
    #        [Parameter()]
    #        [System.String]$ClientName='*',
    #
    #        [Parameter()]
    #        [System.String]$UserName='*'
    #    )
    #
    #
    #    begin
    #    {
    #        try
    #        {
    #            $FuncName = $MyInvocation.MyCommand
    #            Write-Verbose "[$funcName] Entering Begin block."
    #
    #            if(!$ComputerName)
    #            {
    #                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #                $ComputerName = Get-TSGlobalServerName
    #            }
    #            else
    #            {
    #                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #            }
    #
    #
    #            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #            $TSManager = New-Object Cassia.TerminalServicesManager
    #            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #            $TSRemoteServer.Open()
    #
    #            if(!$TSRemoteServer.IsOpen)
    #            {
    #                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #            }
    #
    #            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #            $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    Process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #            if($PSCmdlet.ParameterSetName -eq 'Session')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Id -lt 0)
    #                {
    #                    $session = $TSRemoteServer.GetSessions()
    #                }
    #                else
    #                {
    #                    $session = $TSRemoteServer.GetSession($Id)
    #                }
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session = $InputObject
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Filter')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #
    #                $TSRemoteServer.GetSessions() | Where-Object $Filter
    #            }
    #
    #            if($session)
    #            {
    #                $session | Where-Object {$_.ConnectionState -like $State -AND $_.UserName -like $UserName -AND $_.ClientName -like $ClientName } | `
    #                Add-Member -MemberType AliasProperty -Name IPAddress -Value ClientIPAddress -PassThru | `
    #                Add-Member -MemberType AliasProperty -Name State -Value ConnectionState -PassThru
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Disconnect-TSSession
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Disconnects any connected user from the session.
    #
    #    .DESCRIPTION
    #        Disconnect-TSSession disconnects any connected user from a session on local or remote computers.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Id
    #        Specifies the session Id number.
    #
    #    .PARAMETER InputObject
    #           Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.
    #
    #    .PARAMETER Synchronous
    #           When the Synchronous parameter is present the command waits until the session is fully disconnected otherwise it returns
    #           immediately, even though the session may not be completely disconnected yet.
    #
    #    .PARAMETER Force
    #           Overrides any confirmations made by the command.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 | Disconnect-TSSession
    #
    #        Description
    #        -----------
    #        Disconnects all connected users from Active sessions on remote computer 'comp1'. The caller is prompted to
    #        By default, the caller is prompted to confirm each action.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -State Active | Disconnect-TSSession -Force
    #
    #        Description
    #        -----------
    #        Disconnects any connected user from Active sessions on remote computer 'comp1'.
    #        By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -State Active -Synchronous | Disconnect-TSSession -Force
    #
    #        Description
    #        -----------
    #        Disconnects any connected user from Active sessions on remote computer 'comp1'. The Synchronous parameter tells the command to
    #        wait until the session is fully disconnected and only tghen it proceeds to the next session object.
    #        By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.
    #
    #    .OUTPUTS
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #        Stop-TSSession
    #        Send-TSMessage
    #    #>
    #
    #    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParameterSetName='Id')]
    #
    #    Param(
    #
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ParameterSetName='Id',
    #            ValueFromPipelineByPropertyName=$true
    #        )]
    #        [Alias('SessionId')]
    #        [System.Int32]$Id,
    #
    #        [Parameter(
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesSession]$InputObject,
    #
    #        [switch]$Synchronous,
    #
    #        [switch]$Force
    #    )
    #
    #    begin
    #    {
    #        try
    #        {
    #            $FuncName = $MyInvocation.MyCommand
    #            Write-Verbose "[$funcName] Entering Begin block."
    #
    #            if(!$ComputerName)
    #            {
    #                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #                $ComputerName = Get-TSGlobalServerName
    #            }
    #            else
    #            {
    #                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #            }
    #
    #            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #            $TSManager = New-Object Cassia.TerminalServicesManager
    #            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #            $TSRemoteServer.Open()
    #
    #            if(!$TSRemoteServer.IsOpen)
    #            {
    #                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #            }
    #
    #            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #            $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    Process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #            if($PSCmdlet.ParameterSetName -eq 'Id')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session = $TSRemoteServer.GetSession($Id)
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session  = $InputObject
    #            }
    #
    #
    #            if($session -ne $null)
    #            {
    #                if($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName,"Disconnecting session id '$($session.sessionId)'"))
    #                {
    #                    if($session.ConnectionState -ne 'Disconnected')
    #                    {
    #                        $session.Disconnect($Synchronous)
    #                    }
    #                    else
    #                    {
    #                        Write-Verbose 'Session is already in Disconnected mode.'
    #                    }
    #                }
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Stop-TSSession
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Logs the session off, disconnecting any user that might be connected.
    #
    #    .DESCRIPTION
    #        Use Stop-TSSession to logoff the session and disconnect any user that might be connected.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Id
    #        Specifies the session Id number.
    #
    #    .PARAMETER InputObject
    #           Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.
    #
    #    .PARAMETER Synchronous
    #           When the Synchronous parameter is present the command waits until the session is fully disconnected otherwise it returns
    #           immediately, even though the session may not be completely disconnected yet.
    #
    #    .PARAMETER Force
    #           Overrides any confirmations made by the command.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 | Stop-TSSession
    #
    #        Description
    #        -----------
    #        logs off all connected users from Active sessions on remote computer 'comp1'. The caller is prompted to
    #        By default, the caller is prompted to confirm each action.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -State Active | Stop-TSSession -Force
    #
    #        Description
    #        -----------
    #        logs off any connected user from Active sessions on remote computer 'comp1'.
    #        By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.
    #
    #    .EXAMPLE
    #        Get-TSSession -ComputerName comp1 -State Active -Synchronous | Stop-TSSession -Force
    #
    #        Description
    #        -----------
    #        logs off any connected user from Active sessions on remote computer 'comp1'. The Synchronous parameter tells the command to
    #        wait until the session is fully disconnected and only tghen it proceeds to the next session object.
    #        By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.
    #
    #    .OUTPUTS
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #        Disconnect-TSSession
    #        Send-TSMessage
    #    #>
    #
    #    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParameterSetName='Id')]
    #
    #    Param(
    #
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ParameterSetName='Id',
    #            ValueFromPipelineByPropertyName=$true
    #        )]
    #        [Alias('SessionId')]
    #        [System.Int32]$Id,
    #
    #        [Parameter(
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesSession]$InputObject,
    #
    #        [switch]$Synchronous,
    #
    #        [switch]$Force
    #    )
    #
    #    begin
    #    {
    #        try
    #        {
    #            $FuncName = $MyInvocation.MyCommand
    #            Write-Verbose "[$funcName] Entering Begin block."
    #
    #            if(!$ComputerName)
    #            {
    #                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #                $ComputerName = Get-TSGlobalServerName
    #            }
    #            else
    #            {
    #                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #            }
    #
    #            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #            $TSManager = New-Object Cassia.TerminalServicesManager
    #            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #            $TSRemoteServer.Open()
    #
    #            if(!$TSRemoteServer.IsOpen)
    #            {
    #                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #            }
    #
    #            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #            $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #
    #    Process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Id')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session = $TSRemoteServer.GetSession($Id)
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session  = $InputObject
    #            }
    #
    #            if($session -ne $null)
    #            {
    #                if($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName,"Logging off session id '$($session.sessionId)'"))
    #                {
    #                    Write-Verbose "[$FuncName] Logging off session '$($session.SessionId)'"
    #                    $session.Logoff($Synchronous)
    #                }
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Get-TSProcess
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Gets a list of processes running in a specific session or in all sessions.
    #
    #    .DESCRIPTION
    #        Use Get-TSProcess to get a list of session processes from a local or remote computers.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Id
    #        Specifies the process Id number.
    #
    #    .PARAMETER InputObject
    #           Specifies a process object. Enter a variable that contains the object, or type a command or expression that gets the sessions.
    #
    #    .PARAMETER Name
    #           Specifies the process name. Wildcards are permitted.
    #
    #    .PARAMETER Session
    #        Specifies the session Id number.
    #
    #    .EXAMPLE
    #        Get-TSProcess
    #
    #        Description
    #        -----------
    #        Gets all the sessions processes from the local computer.
    #
    #    .EXAMPLE
    #        Get-TSSession -Id 0 -ComputerName comp1 | Get-TSProcess
    #
    #        Description
    #        -----------
    #        Gets all processes connected to session id 0 from remote computer 'comp1'.
    #
    #    .EXAMPLE
    #        Get-TSProcess -Name s* -ComputerName comp1
    #
    #        Description
    #        -----------
    #        Gets all the processes with name starts with the letter 's' from remote computer 'comp1'.
    #
    #    .OUTPUTS
    #        Cassia.Impl.TerminalServicesProcess
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #        Stop-TSProcess
    #    #>
    #
    #
    #    [OutputType('Cassia.Impl.TerminalServicesProcess')]
    #    [CmdletBinding(DefaultParameterSetName='Name')]
    #
    #    Param(
    #
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server,
    #
    #        [Parameter(
    #            Position=0,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Name'
    #        )]
    #        [Alias('ProcessName')]
    #        [System.String]$Name='*',
    #
    #        [Parameter(
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Id'
    #        )]
    #        [Alias('ProcessID')]
    #        [ValidateRange(0,65536)]
    #        [System.Int32]$Id=-1,
    #
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesProcess]$InputObject,
    #
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='Session'
    #        )]
    #        [Alias('SessionId')]
    #        [Cassia.Impl.TerminalServicesSession]$Session
    #    )
    #
    #
    #
    #    begin
    #    {
    #        $FuncName = $MyInvocation.MyCommand
    #        Write-Verbose "[$funcName] Entering Begin block."
    #
    #        if(!$ComputerName)
    #        {
    #            Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #            $ComputerName = Get-TSGlobalServerName
    #        }
    #        else
    #        {
    #            $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #
    #        Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #        $TSManager = New-Object Cassia.TerminalServicesManager
    #        $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #        $TSRemoteServer.Open()
    #
    #        if(!$TSRemoteServer.IsOpen)
    #        {
    #            Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #        }
    #
    #        Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #        Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #        $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #    }
    #
    #
    #
    #    Process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Name')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Name -eq '*')
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses()
    #                }
    #                else
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses() | Where-Object {$_.ProcessName -like $Name}
    #                }
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Id')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Id -lt 0)
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses()
    #                }
    #                else
    #                {
    #                    $proc = $TSRemoteServer.GetProcess($Id)
    #                }
    #            }
    #
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Session')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Session)
    #                {
    #                    $proc = $Session.GetProcesses()
    #                }
    #            }
    #
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $proc = $InputObject
    #            }
    #
    #
    #            if($proc)
    #            {
    #                $proc
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Stop-TSProcess
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Terminates the process running in a specific session or in all sessions.
    #
    #    .DESCRIPTION
    #        Use Stop-TSProcess to terminate one or more processes from a local or remote computers.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Id
    #        Specifies the process Id number.
    #
    #    .PARAMETER InputObject
    #        Specifies a process object. Enter a variable that contains the object, or type a command or expression that gets the sessions.
    #
    #    .PARAMETER Name
    #        Specifies the process name.
    #
    #    .PARAMETER Session
    #        Specifies the session Id number.
    #
    #    .PARAMETER Force
    #           Overrides any confirmations made by the command.
    #
    #    .EXAMPLE
    #         Get-TSProcess -Id 6552 | Stop-TSProcess
    #
    #        Description
    #        -----------
    #        Gets process Id 6552 from the local computer and stop it. Confirmations needed.
    #
    #    .EXAMPLE
    #        Get-TSSession -Id 3 -ComputerName comp1 | Stop-TSProcess -Force
    #
    #        Description
    #        -----------
    #        Terminats all processes connected to session id 3 from remote computer 'comp1', suppress confirmations.
    #
    #    .OUTPUTS
    #        Cassia.Impl.TerminalServicesProcess
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSProcess
    #        Get-TSSession
    #    #>
    #
    #    [CmdletBinding(SupportsShouldProcess=$true,ConfirmImpact='High',DefaultParameterSetName='Name')]
    #
    #    Param(
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server,
    #
    #        [Parameter(
    #            Position=0,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Name'
    #        )]
    #        [Alias("ProcessName")]
    #        [System.String]$Name='*',
    #
    #        [Parameter(
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Id'
    #        )]
    #        [Alias('ProcessID')]
    #        [ValidateRange(0,65536)]
    #        [System.Int32]$Id=-1,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesProcess]$InputObject,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='Session'
    #        )]
    #        [Alias('SessionId')]
    #        [Cassia.Impl.TerminalServicesSession]$Session,
    #
    #        [switch]$Force
    #    )
    #
    #
    #    begin
    #    {
    #        try
    #        {
    #            $FuncName = $MyInvocation.MyCommand
    #            Write-Verbose "[$funcName] Entering Begin block."
    #
    #            if(!$ComputerName)
    #            {
    #                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #                $ComputerName = Get-TSGlobalServerName
    #            }
    #            else
    #            {
    #                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #            }
    #
    #            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #            $TSManager = New-Object Cassia.TerminalServicesManager
    #            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #            $TSRemoteServer.Open()
    #
    #            if(!$TSRemoteServer.IsOpen)
    #            {
    #                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #            }
    #
    #            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #            $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    Process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Name')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Name -eq '*')
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses()
    #                }
    #                else
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses() | Where-Object {$_.ProcessName -like $Name}
    #                }
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Id')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Id -lt 0)
    #                {
    #                    $proc = $TSRemoteServer.GetProcesses()
    #                }
    #                else
    #                {
    #                    $proc = $TSRemoteServer.GetProcess($Id)
    #                }
    #            }
    #
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Session')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Session)
    #                {
    #                    $proc = $Session.GetProcesses()
    #                }
    #            }
    #
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $proc = $InputObject
    #            }
    #
    #
    #            if($proc)
    #            {
    #                foreach($p in $proc)
    #                {
    #                    if($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName,"Stop Process '$($p.ProcessName) ($($p.ProcessID))"))
    #                    {
    #                        Write-Verbose "[$FuncName] Killing process '$($p.ProcessName)' ($($p.ProcessId))"
    #                        $p.Kill()
    #                    }
    #                }
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Send-TSMessage
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Displays a message box in the specified session Id.
    #
    #    .DESCRIPTION
    #        Use Send-TSMessage display a message box in the specified session Id.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER Text
    #        The text to display in the message box.
    #
    #    .PARAMETER SessionID
    #        The number of the session Id.
    #
    #    .PARAMETER Caption
    #           The caption of the message box. The default caption is 'Alert'.
    #
    #    .EXAMPLE
    #        $Message = "Importnat message`n, the server is going down for maintanace in 10 minutes. Please save your work and logoff."
    #        Get-TSSession -State Active -ComputerName comp1 | Send-TSMessage -Message $Message
    #
    #        Description
    #        -----------
    #        Displays a message box inside all active sessions of computer name 'comp1'.
    #
    #    .OUTPUTS
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #    #>
    #
    #
    #    [CmdletBinding(DefaultParameterSetName='Session')]
    #
    #    Param(
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            HelpMessage='The text to display in the message box.'
    #        )]
    #        [System.String]$Text,
    #
    #        [Parameter(
    #            HelpMessage='The caption of the message box.'
    #        )]
    #        [ValidateNotNullOrEmpty()]
    #        [System.String]$Caption='Alert',
    #
    #        [Parameter(
    #            Position=0,
    #            ValueFromPipelineByPropertyName=$true,
    #            ParameterSetName='Session'
    #        )]
    #        [Alias('SessionID')]
    #        [ValidateRange(0,65536)]
    #        [System.Int32]$Id=-1,
    #
    #        [Parameter(
    #            Position=0,
    #            Mandatory=$true,
    #            ValueFromPipeline=$true,
    #            ParameterSetName='InputObject'
    #        )]
    #        [Cassia.Impl.TerminalServicesSession]$InputObject
    #    )
    #
    #    begin
    #    {
    #        try
    #        {
    #            $FuncName = $MyInvocation.MyCommand
    #            Write-Verbose "[$funcName] Entering Begin block."
    #
    #            if(!$ComputerName)
    #            {
    #                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
    #                $ComputerName = Get-TSGlobalServerName
    #            }
    #            else
    #            {
    #                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #            }
    #
    #            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
    #            $TSManager = New-Object Cassia.TerminalServicesManager
    #            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #            $TSRemoteServer.Open()
    #
    #            if(!$TSRemoteServer.IsOpen)
    #            {
    #                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #            }
    #
    #            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
    #            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
    #            $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    process
    #    {
    #
    #        Write-Verbose "[$funcName] Entering Process block."
    #
    #        try
    #        {
    #
    #            if($PSCmdlet.ParameterSetName -eq 'Session')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                if($Id -ge 0)
    #                {
    #                    $session = $TSRemoteServer.GetSession($Id)
    #                }
    #            }
    #
    #            if($PSCmdlet.ParameterSetName -eq 'InputObject')
    #            {
    #                Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
    #                $session = $InputObject
    #            }
    #
    #            if($session)
    #            {
    #                Write-Verbose "[$FuncName] Sending alert message to session id: '$($session.SessionId)' on '$ComputerName'"
    #                $session.MessageBox($Text,$Caption)
    #            }
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #
    #
    #    end
    #    {
    #        try
    #        {
    #            Write-Verbose "[$funcName] Entering End block."
    #            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #            $TSRemoteServer.Close()
    #            $TSRemoteServer.Dispose()
    #        }
    #        catch
    #        {
    #            Throw
    #        }
    #    }
    #}
    #
    #function Get-TSServers
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Enumerates all terminal servers in a given domain.
    #
    #    .DESCRIPTION
    #        Enumerates all terminal servers in a given domain.
    #
    #    .PARAMETER ComputerName
    #            The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).
    #
    #    .PARAMETER DomainName
    #        The name of the domain. The default is the caller domain name ($env:USERDOMAIN).
    #
    #    .EXAMPLE
    #        Get-TSDomainServers
    #
    #        Description
    #        -----------
    #        Get a list of all terminal servers of the caller default domain.
    #
    #    .OUTPUTS
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #    #>
    #
    #
    #    [OutputType('System.Management.Automation.PSCustomObject')]
    #    [CmdletBinding()]
    #
    #    Param(
    #        [Parameter(
    #            Position=0,
    #            ParameterSetName='Name'
    #        )]
    #        [System.String]$DomainName=$env:USERDOMAIN
    #    )
    #
    #
    #    try
    #    {
    #        $FuncName = $MyInvocation.MyCommand
    #        if(!$ComputerName)
    #        {
    #            Write-Verbose "[$funcName] ComputerName is not defined, loading global value '$script:Server'."
    #            $ComputerName = Get-TSGlobalServerName
    #        }
    #        else
    #        {
    #            $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #
    #        Write-Verbose "[$funcName] Enumerating terminal servers for '$DomainName' domain."
    #        Write-Warning 'Depending on your environment the command may take a while to complete.'
    #        $TSManager = New-Object Cassia.TerminalServicesManager
    #        $TSManager.GetServers($DomainName)
    #    }
    #    catch
    #    {
    #        Throw
    #    }
    #
    #}
    #
    #function Get-TSCurrentSession
    #{
    #
    #    <#
    #    .SYNOPSIS
    #        Provides information about the session in which the current process is running.
    #
    #    .DESCRIPTION
    #        Provides information about the session in which the current process is running.
    #
    #    .EXAMPLE
    #        Get-TSCurrentSession
    #
    #        Description
    #        -----------
    #        Displays the session in which the current process is running on the local computer.
    #
    #    .OUTPUTS
    #        Cassia.Impl.TerminalServicesSession
    #
    #    .COMPONENT
    #        TerminalServer
    #
    #    .NOTES
    #        Author: Shay Levy
    #        Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/
    #
    #    .LINK
    #        http://code.msdn.microsoft.com/PSTerminalServices
    #
    #    .LINK
    #        http://code.google.com/p/cassia/
    #
    #    .LINK
    #        Get-TSSession
    #    #>
    #
    #
    #    [OutputType('Cassia.Impl.TerminalServicesSession')]
    #    [CmdletBinding()]
    #
    #    param(
    #        [Parameter()]
    #        [Alias('CN','IPAddress')]
    #        [System.String]$ComputerName=$script:server
    #    )
    #
    #
    #    try
    #    {
    #        $FuncName = $MyInvocation.MyCommand
    #
    #        if(!$ComputerName)
    #        {
    #            Write-Verbose "[$funcName] ComputerName is not defined, loading global value '$script:Server'."
    #            $ComputerName = Get-TSGlobalServerName
    #        }
    #        else
    #        {
    #            $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
    #        }
    #
    #        Write-Verbose "[$funcName] Attempting remote connection to '$ComputerName'"
    #        $TSManager = New-Object Cassia.TerminalServicesManager
    #        $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
    #        $TSRemoteServer.Open()
    #
    #        if(!$TSRemoteServer.IsOpen)
    #        {
    #            Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
    #        }
    #
    #        Write-Verbose "[$funcName] Connection is open '$ComputerName'"
    #        Write-Verbose "[$funcName] Updating global Server name '$ComputerName'"
    #        $null = Set-TSGlobalServerName -ComputerName $ComputerName
    #
    #        Write-Verbose "[$funcName] Get CurrentSession from '$ComputerName'"
    #        $TSManager.CurrentSession
    #
    #        Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
    #        $TSRemoteServer.Close()
    #        $TSRemoteServer.Dispose()
    #    }
    #    catch
    #    {
    #        Throw
    #    }
    #}
    #
    #function Set-TSGlobalServerName
    #{
    #    [CmdletBinding()]
    #
    #    Param(
    #        [Parameter(Mandatory=$true)]
    #        [ValidateNotNullOrEmpty()]
    #        [System.String]$ComputerName
    #    )
    #
    #    if($ComputerName -eq "." -OR $ComputerName -eq $env:COMPUTERNAME)
    #    {
    #        $ComputerName='localhost'
    #    }
    #
    #    $script:Server=$ComputerName
    #    $script:Server
    #}
    #
    #function Get-TSGlobalServerName
    #{
    #    $script:Server
    #}
    #
    #
#endregion Source: Globals.ps1

#region Source: functions\Append-RichtextboxStatus.ps1
function Invoke-Append-RichtextboxStatus_ps1
{
    function Append-RichtextboxStatus
    {
        PARAM (
            [Parameter(Mandatory = $true)]
            [string]$Message,
            [string]$MessageColor = "DarkGreen",
            [string]$DateTimeColor = "Black",
            [string]$Source,
            [string]$SourceColor = "Gray",
            [string]$ComputerName,
            [String]$ComputerNameColor = "Blue")

        $SortableTime = get-date -Format "yyyy-MM-dd HH:mm:ss"
        $richtextboxStatus.SelectionColor = $DateTimeColor
        $richtextboxStatus.AppendText("[$SortableTime] ")

        IF ($PSBoundParameters['ComputerName'])
        {
            $richtextboxStatus.SelectionColor = $ComputerNameColor
            $richtextboxStatus.AppendText(("$ComputerName ").ToUpper())
        }

        IF ($PSBoundParameters['Source'])
        {
            $richtextboxStatus.SelectionColor = $SourceColor
            $richtextboxStatus.AppendText("$Source ")
        }

        $richtextboxStatus.SelectionColor = $MessageColor
        $richtextboxStatus.AppendText("$Message`r")
        $richtextboxStatus.Refresh()
        $richtextboxStatus.ScrollToCaret()

        Write-Verbose -Message "$SortableTime $Message"
    }
}
#endregion Source: Append-RichtextboxStatus.ps1

#region Source: functions\Clear-DataGridViewSelection.ps1
function Invoke-Clear-DataGridViewSelection_ps1
{
    function Clear-DataGridViewSelection
    {
        PARAM (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView
        )
        $DataGridView.ClearSelection()
    }
}
#endregion Source: Clear-DataGridViewSelection.ps1

#region Source: functions\ConvertTo-DataTable.ps1
function Invoke-ConvertTo-DataTable_ps1
{
    function ConvertTo-DataTable
    {
        <#
            .SYNOPSIS
                Converts objects into a DataTable.

            .DESCRIPTION
                Converts objects into a DataTable, which are used for DataBinding.

            .PARAMETER  InputObject
                The input to convert into a DataTable.

            .PARAMETER  Table
                The DataTable you wish to load the input into.

            .PARAMETER RetainColumns
                This switch tells the function to keep the DataTable's existing columns.

            .PARAMETER FilterWMIProperties
                This switch removes WMI properties that start with an underline.

            .EXAMPLE
                $DataTable = ConvertTo-DataTable -InputObject (Get-Process)

            .NOTES
                SAPIEN Technologies, Inc.
                http://www.sapien.com/

        #>
        [CmdletBinding()]
        [OutputType([System.Data.DataTable])]
        param (
            [ValidateNotNull()]
            $InputObject,
            [ValidateNotNull()]
            [System.Data.DataTable]$Table,
            [switch]$RetainColumns,
            [switch]$FilterWMIProperties
        )

        if ($null -eq $Table)
        {
            $Table = New-Object System.Data.DataTable
        }

        if ($InputObject -is [System.Data.DataTable])
        {
            $Table = $InputObject
        }
        else
        {
            if (-not $RetainColumns -or $Table.Columns.Count -eq 0)
            {
                #Clear out the Table Contents
                $Table.Clear()

                if ($null -eq $InputObject) { return } #Empty Data

                $object = $null

                #find the first non null value
                foreach ($item in $InputObject)
                {
                    if ($null -ne $item)
                    {
                        $object = $item
                        break
                    }
                }

                if ($null -eq $object) { return } #All null then empty

                #COLUMN
                #Get all the properties in order to create the columns
                foreach ($prop in $object.PSObject.Get_Properties())
                {
                    if (-not $FilterWMIProperties -or -not $prop.Name.StartsWith('__')) #filter out WMI properties
                    {
                        #Get the type from the Definition string
                        $type = $null

                        if ($null -ne $prop.Value)
                        {
                            try { $type = $prop.Value.GetType() }
                            catch { Write-Verbose -Message "Can't find type of $prop" }
                        }

                        if ($null -ne $type) # -and [System.Type]::GetTypeCode($type) -ne 'Object')
                        {
                            Write-Verbose -Message "Creating Column: $($Prop.name) (Type: $type)"
                            [void]$table.Columns.Add($prop.Name, $type)
                        }
                        else #Type info not found
                        {
                            #if ($prop.name -eq "" -or $prop.name -eq $null) { [void]$table.Columns.Add([DBNull]::Value) }
                            [void]$table.Columns.Add($prop.Name)
                        }
                    }
                }

                if ($object -is [System.Data.DataRow])
                {
                    foreach ($item in $InputObject)
                    {
                        $Table.Rows.Add($item)
                    }
                    return @( , $Table)
                }
            }
            else
            {
                $Table.Rows.Clear()
            }

            #Rows Work
            foreach ($item in $InputObject)
            {
                # Create a new row object
                $row = $table.NewRow()

                if ($item)
                {
                    foreach ($prop in $item.PSObject.Get_Properties())
                    {
                        #Find the appropriate column to put the value
                        if ($table.Columns.Contains($prop.Name))
                        {
                            if ($null -eq $prop.value) { $prop.value = [DBNull]::Value }
                            $row.Item($prop.Name) = $prop.Value
                        }
                    }
                }
                [void]$table.Rows.Add($row)
            }
        }

        return @( , $Table)
    }
}
#endregion Source: ConvertTo-DataTable.ps1

#region Source: functions\Disable-Button.ps1
function Invoke-Disable-Button_ps1
{
    function Disable-Button
    {
    <#
    .SYNOPSIS
        This function will disable a button control
    .EXAMPLE
        Disable-Button -Button $Button
    #>
        [CmdletBinding()]
        PARAM (
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Button[]]$Button
        )
        BEGIN
        {
            Add-Type -AssemblyName System.Windows.Forms
        }
        PROCESS
        {
            foreach ($ButtonObject in $Button)
            {
                $ButtonObject.Enabled = $false
            }

        }
    } #Disable-Button
}
#endregion Source: Disable-Button.ps1

#region Source: functions\Disconnect-TSSession.ps1
function Invoke-Disconnect-TSSession_ps1
{
    function Disconnect-TSSession
    {

        <#
        .SYNOPSIS
            Disconnects any connected user from the session.

        .DESCRIPTION
            Disconnect-TSSession disconnects any connected user from a session on local or remote computers.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Id
            Specifies the session Id number.

        .PARAMETER InputObject
               Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.

        .PARAMETER Synchronous
               When the Synchronous parameter is present the command waits until the session is fully disconnected otherwise it returns
               immediately, even though the session may not be completely disconnected yet.

        .PARAMETER Force
               Overrides any confirmations made by the command.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 | Disconnect-TSSession

            Description
            -----------
            Disconnects all connected users from Active sessions on remote computer 'comp1'. The caller is prompted to
            By default, the caller is prompted to confirm each action.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -State Active | Disconnect-TSSession -Force

            Description
            -----------
            Disconnects any connected user from Active sessions on remote computer 'comp1'.
            By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -State Active -Synchronous | Disconnect-TSSession -Force

            Description
            -----------
            Disconnects any connected user from Active sessions on remote computer 'comp1'. The Synchronous parameter tells the command to
            wait until the session is fully disconnected and only tghen it proceeds to the next session object.
            By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.

        .OUTPUTS

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
            Stop-TSSession
            Send-TSMessage
        #>

        [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Id')]
        Param (

            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ParameterSetName = 'Id',
                       ValueFromPipelineByPropertyName = $true
            )]
            [Alias('SessionId')]
            [System.Int32]$Id,
            [Parameter(
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesSession]$InputObject,
            [switch]$Synchronous,
            [switch]$Force
        )

        begin
        {
            try
            {
                $FuncName = $MyInvocation.MyCommand
                Write-Verbose "[$funcName] Entering Begin block."

                if (!$ComputerName)
                {
                    Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                    $ComputerName = Get-TSGlobalServerName
                }
                else
                {
                    $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
                }

                Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
                $TSManager = New-Object Cassia.TerminalServicesManager
                $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
                $TSRemoteServer.Open()

                if (!$TSRemoteServer.IsOpen)
                {
                    Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
                }

                Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
                Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
                $null = Set-TSGlobalServerName -ComputerName $ComputerName
            }
            catch
            {
                Throw
            }
        }


        Process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {
                if ($PSCmdlet.ParameterSetName -eq 'Id')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $TSRemoteServer.GetSession($Id)
                }

                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $InputObject
                }


                if ($null -ne $session)
                {
                    if ($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName, "Disconnecting session id '$($session.sessionId)'"))
                    {
                        if ($session.ConnectionState -ne 'Disconnected')
                        {
                            $session.Disconnect($Synchronous)
                        }
                        else
                        {
                            Write-Verbose 'Session is already in Disconnected mode.'
                        }
                    }
                }
            }
            catch
            {
                Throw
            }
        }

        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }
}
#endregion Source: Disconnect-TSSession.ps1

#region Source: functions\Enable-Button.ps1
function Invoke-Enable-Button_ps1
{
    function Enable-Button
    {
    <#
    .SYNOPSIS
        This function will enable a button control
    .EXAMPLE
        Enable-Button -Button $Button
    #>
        [CmdletBinding()]
        PARAM (
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.Button[]]$Button
        )
        BEGIN
        {
            Add-Type -AssemblyName System.Windows.Forms
        }
        PROCESS
        {
            foreach ($ButtonObject in $Button)
            {
                $ButtonObject.Enabled = $true
            }
        }
    } #Enable-Button

}
#endregion Source: Enable-Button.ps1

#region Source: functions\Find-DataGridViewValue.ps1
function Invoke-Find-DataGridViewValue_ps1
{
    function Find-DataGridViewValue
    {
    <#
        .SYNOPSIS
            The Find-DataGridViewValue function helps you to find a specific value and select the cell, row or to set a fore and back color.

        .DESCRIPTION
            The Find-DataGridViewValue function helps you to find a specific value and select the cell, row or to set a fore and back color.

        .PARAMETER DataGridView
            Specifies the DataGridView Control to use

        .PARAMETER RowBackColor
            Specifies the back color of the row to use

        .PARAMETER RowForeColor
            Specifies the fore color of the row to use

        .PARAMETER SelectCell
            Specifies to select only the cell when the value is found

        .PARAMETER SelectRow
            Specifies to select the entire row when the value is found

        .PARAMETER Value
            Specifies the value to search

        .EXAMPLE
            PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text

            This will find the value and select the cell(s)

        .EXAMPLE
            PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text -RowForeColor 'Red' -RowBackColor 'Black'

            This will find the value and color the fore and back of the row
        .EXAMPLE
            PS C:\> Find-DataGridViewValue -DataGridView $datagridview1 -Value $textbox1.Text -SelectRow

            This will find the value and select the entire row

    #>
        [CmdletBinding(DefaultParameterSetName = "Cell")]
        PARAM (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView,
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [String]$Value,
            [Parameter(ParameterSetName = "Cell")]
            [Switch]$SelectCell,
            [Parameter(ParameterSetName = "Row")]
            [Switch]$SelectRow,
            #[Parameter(ParameterSetName = "Column")]

            #[Switch]$SelectColumn,

            [Parameter(ParameterSetName = "RowColor")]
            [system.Drawing.Color]$RowForeColor,
            [Parameter(ParameterSetName = "RowColor")]
            [system.Drawing.Color]$RowBackColor
        )

        PROCESS
        {
            $DataGridView.ClearSelection()

            FOR ([int]$i = 0; $i -lt $DataGridView.RowCount; $i++)
            {
                FOR ([int]$j = 0; $j -lt $DataGridView.ColumnCount; $j++)
                {
                    $CurrentCell = $dataGridView.Rows[$i].Cells[$j]

                    #if ((-not $CurrentCell.Value.Equals([DBNull]::Value)) -and ($CurrentCell.Value.ToString() -like "*$Value*"))
                    if ($CurrentCell.Value.ToString() -match $Value)
                    {

                        # Row Selection
                        IF ($PSBoundParameters['SelectRow'])
                        {
                            $dataGridView.Rows[$i].Selected = $true
                        }

                        <#
                        # Column Selection
                        IF ($PSBoundParameters['SelectColumn'])
                        {
                            #$DataGridView.Columns[$($CurrentCell.ColumnIndex)].Selected = $true
                            #$DataGridView.Columns[$j].Selected = $true
                            #$CurrentCell.DataGridView.Columns[$j].Selected = $true
                        }
                        #>

                        # Row Fore Color
                        IF ($PSBoundParameters['RowForeColor'])
                        {
                            $dataGridView.Rows[$i].DefaultCellStyle.ForeColor = $RowForeColor
                        }
                        # Row Back Color
                        IF ($PSBoundParameters['RowBackColor'])
                        {
                            $dataGridView.Rows[$i].DefaultCellStyle.BackColor = $RowBackColor
                        }

                        # Cell Selection
                        ELSEIF (-not ($PSBoundParameters['SelectRow']) -and -not ($PSBoundParameters['SelectColumn']))
                        {
                            $CurrentCell.Selected = $true
                        }
                    } #IF not empty and contains value
                } #For Each column
            } #For Each Row
        } #PROCESS
    } #Find-DataGridViewValue
}
#endregion Source: Find-DataGridViewValue.ps1

#region Source: functions\Get-TSCurrentSession.ps1
function Invoke-Get-TSCurrentSession_ps1
{
    function Get-TSCurrentSession
    {

        <#
        .SYNOPSIS
            Provides information about the session in which the current process is running.

        .DESCRIPTION
            Provides information about the session in which the current process is running.

        .EXAMPLE
            Get-TSCurrentSession

            Description
            -----------
            Displays the session in which the current process is running on the local computer.

        .OUTPUTS
            Cassia.Impl.TerminalServicesSession

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
        #>


        [OutputType('Cassia.Impl.TerminalServicesSession')]
        [CmdletBinding()]
        param (
            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server
        )


        try
        {
            $FuncName = $MyInvocation.MyCommand

            if (!$ComputerName)
            {
                Write-Verbose "[$funcName] ComputerName is not defined, loading global value '$script:Server'."
                $ComputerName = Get-TSGlobalServerName
            }
            else
            {
                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
            }

            Write-Verbose "[$funcName] Attempting remote connection to '$ComputerName'"
            $TSManager = New-Object Cassia.TerminalServicesManager
            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
            $TSRemoteServer.Open()

            if (!$TSRemoteServer.IsOpen)
            {
                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
            }

            Write-Verbose "[$funcName] Connection is open '$ComputerName'"
            Write-Verbose "[$funcName] Updating global Server name '$ComputerName'"
            $null = Set-TSGlobalServerName -ComputerName $ComputerName

            Write-Verbose "[$funcName] Get CurrentSession from '$ComputerName'"
            $TSManager.CurrentSession

            Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
            $TSRemoteServer.Close()
            $TSRemoteServer.Dispose()
        }
        catch
        {
            Throw
        }
    }
}
#endregion Source: Get-TSCurrentSession.ps1

#region Source: functions\Get-TSGlobalServerName.ps1
function Invoke-Get-TSGlobalServerName_ps1
{
    function Get-TSGlobalServerName
    {
        $script:Server
    }
}
#endregion Source: Get-TSGlobalServerName.ps1

#region Source: functions\Get-TSProcess.ps1
function Invoke-Get-TSProcess_ps1
{
    function Get-TSProcess
    {

        <#
        .SYNOPSIS
            Gets a list of processes running in a specific session or in all sessions.

        .DESCRIPTION
            Use Get-TSProcess to get a list of session processes from a local or remote computers.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Id
            Specifies the process Id number.

        .PARAMETER InputObject
               Specifies a process object. Enter a variable that contains the object, or type a command or expression that gets the sessions.

        .PARAMETER Name
               Specifies the process name. Wildcards are permitted.

        .PARAMETER Session
            Specifies the session Id number.

        .EXAMPLE
            Get-TSProcess

            Description
            -----------
            Gets all the sessions processes from the local computer.

        .EXAMPLE
            Get-TSSession -Id 0 -ComputerName comp1 | Get-TSProcess

            Description
            -----------
            Gets all processes connected to session id 0 from remote computer 'comp1'.

        .EXAMPLE
            Get-TSProcess -Name s* -ComputerName comp1

            Description
            -----------
            Gets all the processes with name starts with the letter 's' from remote computer 'comp1'.

        .OUTPUTS
            Cassia.Impl.TerminalServicesProcess

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
            Stop-TSProcess
        #>


        [OutputType('Cassia.Impl.TerminalServicesProcess')]
        [CmdletBinding(DefaultParameterSetName = 'Name')]
        Param (

            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server,
            [Parameter(
                       Position = 0,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Name'
            )]
            [Alias('ProcessName')]
            [System.String]$Name = '*',
            [Parameter(
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Id'
            )]
            [Alias('ProcessID')]
            [ValidateRange(0, 65536)]
            [System.Int32]$Id = -1,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesProcess]$InputObject,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'Session'
            )]
            [Alias('SessionId')]
            [Cassia.Impl.TerminalServicesSession]$Session
        )



        begin
        {
            $FuncName = $MyInvocation.MyCommand
            Write-Verbose "[$funcName] Entering Begin block."

            if (!$ComputerName)
            {
                Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                $ComputerName = Get-TSGlobalServerName
            }
            else
            {
                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
            }

            Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
            $TSManager = New-Object Cassia.TerminalServicesManager
            $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
            $TSRemoteServer.Open()

            if (!$TSRemoteServer.IsOpen)
            {
                Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
            }

            Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
            Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
            $null = Set-TSGlobalServerName -ComputerName $ComputerName
        }



        Process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {

                if ($PSCmdlet.ParameterSetName -eq 'Name')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Name -eq '*')
                    {
                        $proc = $TSRemoteServer.GetProcesses()
                    }
                    else
                    {
                        $proc = $TSRemoteServer.GetProcesses() | Where-Object { $_.ProcessName -like $Name }
                    }
                }

                if ($PSCmdlet.ParameterSetName -eq 'Id')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Id -lt 0)
                    {
                        $proc = $TSRemoteServer.GetProcesses()
                    }
                    else
                    {
                        $proc = $TSRemoteServer.GetProcess($Id)
                    }
                }


                if ($PSCmdlet.ParameterSetName -eq 'Session')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Session)
                    {
                        $proc = $Session.GetProcesses()
                    }
                }


                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $proc = $InputObject
                }


                if ($proc)
                {
                    $proc
                }
            }
            catch
            {
                Throw
            }
        }


        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }

}
#endregion Source: Get-TSProcess.ps1

#region Source: functions\Get-TSServers.ps1
function Invoke-Get-TSServers_ps1
{
    function Get-TSServers
    {

        <#
        .SYNOPSIS
            Enumerates all terminal servers in a given domain.

        .DESCRIPTION
            Enumerates all terminal servers in a given domain.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER DomainName
            The name of the domain. The default is the caller domain name ($env:USERDOMAIN).

        .EXAMPLE
            Get-TSDomainServers

            Description
            -----------
            Get a list of all terminal servers of the caller default domain.

        .OUTPUTS

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
        #>


        [OutputType('System.Management.Automation.PSCustomObject')]
        [CmdletBinding()]
        Param (
            [Parameter(
                       Position = 0,
                       ParameterSetName = 'Name'
            )]
            [System.String]$DomainName = $env:USERDOMAIN
        )


        try
        {
            $FuncName = $MyInvocation.MyCommand
            if (!$ComputerName)
            {
                Write-Verbose "[$funcName] ComputerName is not defined, loading global value '$script:Server'."
                $ComputerName = Get-TSGlobalServerName
            }
            else
            {
                $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
            }

            Write-Verbose "[$funcName] Enumerating terminal servers for '$DomainName' domain."
            Write-Warning 'Depending on your environment the command may take a while to complete.'
            $TSManager = New-Object Cassia.TerminalServicesManager
            $TSManager.GetServers($DomainName)
        }
        catch
        {
            Throw
        }

    }
}
#endregion Source: Get-TSServers.ps1

#region Source: functions\Get-TSSession.ps1
function Invoke-Get-TSSession_ps1
{
    function Get-TSSession
    {
        <#
        .SYNOPSIS
            Lists the sessions on a given terminal server.

        .DESCRIPTION
            Use Get-TSSession to get a list of sessions from a local or remote computers.
            Note that Get-TSSession is using Aliased properties to display the output on the console (IPAddress and State), these attributes
            are not the same as the original attributes (ClientIPAddress and ConnectionState).
            This is important when you want to use the -Filter parameter which requires the latter.
            To see all aliassed properties and their corresponding properties (Definition column), pipe the result to Get-Member:

            PS > Get-TSSession | Get-Member -MemberType AliasProperty

               TypeName: Cassia.Impl.TerminalServicesSession

            Name      MemberType    Definition
            ----      ----------    ----------
            (...)
            IPAddress AliasProperty IPAddress = ClientIPAddress
            State     AliasProperty State = ConnectionState


        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Id
            Specifies the session Id number.

        .PARAMETER InputObject
               Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.

        .PARAMETER Filter
               Specifies a filter based on the session properties. The syntax of the filter, including the use of
               wildcards and depends on the properties of the session. Internally, The Filter parameter uses client side
               filtering using the Where-Object cmdlet, objects are filtered after they are retrieved.

        .PARAMETER State
            The connection state of the session. Use this parameter to get sessions of a specific state. Valid values are:

            Value         Description
            -----         -----------
            Active         A user is logged on to the session.
            ConnectQuery The session is in the process of connecting to a client.
            Connected     A client is connected to the session).
            Disconnected The session is active, but the client has disconnected from it.
            Down         The session is down due to an error.
            Idle         The session is waiting for a client to connect.
            Initializing The session is initializing.
            Listening      The session is listening for connections.
            Reset         The session is being reset.
            Shadowing     This session is shadowing another session.

        .PARAMETER ClientName
            The name of the machine last connected to a session.
            Use this parameter to get sessions made from a specific computer name. Wildcrads are permitted.

        .PARAMETER UserName
            Use this parameter to get sessions made by a specific user name. Wildcrads are permitted.

        .EXAMPLE
            Get-TSSession

            Description
            -----------
            Gets all the sessions from the local computer.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -State Disconnected

            Description
            -----------
            Gets all the disconnected sessions from the remote computer 'comp1'.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -Filter {$_.ClientIPAddress -like '10*' -AND $_.ConnectionState -eq 'Active'}

            Description
            -----------
            Gets all Active sessions from remote computer 'comp1', made from ip addresses that starts with '10'.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -UserName a*

            Description
            -----------
            Gets all sessions from remote computer 'comp1' made by users with name starts with the letter 'a'.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -ClientName s*

            Description
            -----------
            Gets all sessions from remote computer 'comp1' made from a computers names that starts with the letter 's'.

        .OUTPUTS
            Cassia.Impl.TerminalServicesSession

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Stop-TSSession
            Disconnect-TSSession
            Send-TSMessage
        #>


        [OutputType('Cassia.Impl.TerminalServicesSession')]
        [CmdletBinding(DefaultParameterSetName = 'Session')]
        Param (

            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName,
            [Parameter(
                       Position = 0,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Session'
            )]
            [Alias('SessionID')]
            [ValidateRange(0, 65536)]
            [System.Int32]$Id = -1,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesSession]$InputObject,
            [Parameter(
                       Mandatory = $true,
                       ParameterSetName = 'Filter'
            )]
            [ScriptBlock]$Filter,
            [Parameter()]
            [ValidateSet('Active', 'Connected', 'ConnectQuery', 'Shadowing', 'Disconnected', 'Idle', 'Listening', 'Reset', 'Down', 'Initializing')]
            [Alias('ConnectionState')]
            [System.String]$State = '*',
            [Parameter()]
            [System.String]$ClientName = '*',
            [Parameter()]
            [System.String]$UserName = '*'
        )


        begin
        {
            try
            {
                $FuncName = $MyInvocation.MyCommand
                Write-Verbose "[$funcName] Entering Begin block."

                if (!$ComputerName)
                {
                    Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                    $ComputerName = Get-TSGlobalServerName
                }
                else
                {
                    $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
                }


                Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
                $TSManager = New-Object Cassia.TerminalServicesManager
                $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
                $TSRemoteServer.Open()

                if (!$TSRemoteServer.IsOpen)
                {
                    Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
                }

                Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
                Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
                $null = Set-TSGlobalServerName -ComputerName $ComputerName
            }
            catch
            {
                Throw
            }
        }


        Process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {
                if ($PSCmdlet.ParameterSetName -eq 'Session')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Id -lt 0)
                    {
                        $session = $TSRemoteServer.GetSessions()
                    }
                    else
                    {
                        $session = $TSRemoteServer.GetSession($Id)
                    }
                }

                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $InputObject
                }

                if ($PSCmdlet.ParameterSetName -eq 'Filter')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"

                    $TSRemoteServer.GetSessions() | Where-Object $Filter
                }

                if ($session)
                {
                    $session | Where-Object { $_.ConnectionState -like $State -AND $_.UserName -like $UserName -AND $_.ClientName -like $ClientName } | `
                    Add-Member -MemberType AliasProperty -Name IPAddress -Value ClientIPAddress -PassThru | `
                    Add-Member -MemberType AliasProperty -Name State -Value ConnectionState -PassThru
                }
            }
            catch
            {
                Throw
            }
        }

        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }
}
#endregion Source: Get-TSSession.ps1

#region Source: functions\Load-DataGridView.ps1
function Invoke-Load-DataGridView_ps1
{
    function Load-DataGridView
    {
        <#
        .SYNOPSIS
            This functions helps you load items into a DataGridView.

        .DESCRIPTION
            Use this function to dynamically load items into the DataGridView control.

        .PARAMETER  DataGridView
            The ComboBox control you want to add items to.

        .PARAMETER  Item
            The object or objects you wish to load into the ComboBox's items collection.

        .PARAMETER  DataMember
            Sets the name of the list or table in the data source for which the DataGridView is displaying data.
        #>

        [CmdletBinding()]
        Param (
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView,
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            $Item,
            [Parameter(Mandatory = $false)]
            [string]$DataMember
        )
        $DataGridView.SuspendLayout()
        $DataGridView.DataMember = $DataMember

        if ($Item -is [System.ComponentModel.IListSource]`
            -or $Item -is [System.ComponentModel.IBindingList] -or $Item -is [System.ComponentModel.IBindingListView])
        {
            $DataGridView.DataSource = $Item
        }
        else
        {
            $array = New-Object System.Collections.ArrayList

            if ($Item -is [System.Collections.IList])
            {
                $array.AddRange($Item)
            }
            else
            {
                $array.Add($Item)
            }
            $DataGridView.DataSource = $array
        }

        $DataGridView.ResumeLayout()
    }
}
#endregion Source: Load-DataGridView.ps1

#region Source: functions\New-MessageBox.ps1
function Invoke-New-MessageBox_ps1
{
    function New-MessageBox
    {
    <#
        .SYNOPSIS
            The New-MessageBox functio will show a message box to the user

        .DESCRIPTION
            The New-MessageBox functio will show a message box to the user

        .PARAMETER Message
            Specifies the message to show

        .PARAMETER Title
            Specifies the title of the message box

        .PARAMETER Buttons
            Specifies which button to add. Just press tab to see the choices

        .PARAMETER Icon
            Specifies the icon to show. Just press tab to see the choices

        .EXAMPLE
            PS C:\> New-MessageBox -Message "Hello World" -Title "First Message" -Buttons "RetryCancel" -Icon "Asterix"


    #>
        [CmdletBinding()]
        PARAM (

            [String]$Message,
            [String]$Title,
            [System.Windows.Forms.MessageBoxButtons]$Buttons = "OK",
            [System.Windows.Forms.MessageBoxIcon]$Icon = "None"
        )
        BEGIN
        {
            Add-Type -AssemblyName System.Windows.Forms
        }
        PROCESS
        {
            [System.Windows.Forms.MessageBox]::Show($Message, $Title, $Buttons, $Icon)
        }
    } #New-MessageBox
}
#endregion Source: New-MessageBox.ps1

#region Source: functions\Reset-DataGridViewFormat.ps1
function Invoke-Reset-DataGridViewFormat_ps1
{
    function Reset-DataGridViewFormat
    {
    <#
        .SYNOPSIS
            The Reset-DataGridViewFormat function will reset the format of a datagridview control

        .DESCRIPTION
            The Reset-DataGridViewFormat function will reset the format of a datagridview control

        .PARAMETER DataGridView
            Specifies the DataGridView Control.

        .EXAMPLE
            PS C:\> Reset-DataGridViewFormat -DataGridView $DataGridViewObj

    #>
        [CmdletBinding()]
        PARAM (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView)
        PROCESS
        {
            $DataSource = $DataGridView.DataSource
            $DataGridView.DataSource = $null
            $DataGridView.DataSource = $DataSource

            #$DataGridView.RowsDefaultCellStyle.BackColor = 'White'
            #$DataGridView.RowsDefaultCellStyle.ForeColor = 'Black'
            $DataGridView.RowsDefaultCellStyle = $null
            $DataGridView.AlternatingRowsDefaultCellStyle = $null
        }
    } #Reset-DataGridViewFormat
}
#endregion Source: Reset-DataGridViewFormat.ps1

#region Source: functions\Reset-TextBox.ps1
function Invoke-Reset-TextBox_ps1
{
    function Reset-TextBox
    {
        [CmdletBinding()]
        PARAM (
            [System.Windows.Forms.TextBox]$TextBox,
            [System.Drawing.Color]$BackColor = "White",
            [System.Drawing.Color]$ForeColor = "Black"
        )
        BEGIN { }
        PROCESS
        {
            TRY
            {
                $TextBox.Text = ""
                $TextBox.BackColor = $BackColor
                $TextBox.ForeColor = $ForeColor
            }
            CATCH { }
        }
    }
}
#endregion Source: Reset-TextBox.ps1

#region Source: functions\Send-TSMessage.ps1
function Invoke-Send-TSMessage_ps1
{
    function Send-TSMessage
    {

        <#
        .SYNOPSIS
            Displays a message box in the specified session Id.

        .DESCRIPTION
            Use Send-TSMessage display a message box in the specified session Id.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Text
            The text to display in the message box.

        .PARAMETER SessionID
            The number of the session Id.

        .PARAMETER Caption
               The caption of the message box. The default caption is 'Alert'.

        .EXAMPLE
            $Message = "Importnat message`n, the server is going down for maintanace in 10 minutes. Please save your work and logoff."
            Get-TSSession -State Active -ComputerName comp1 | Send-TSMessage -Message $Message

            Description
            -----------
            Displays a message box inside all active sessions of computer name 'comp1'.

        .OUTPUTS

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
        #>


        [CmdletBinding(DefaultParameterSetName = 'Session')]
        Param (
            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       HelpMessage = 'The text to display in the message box.'
            )]
            [System.String]$Text,
            [Parameter(
                       HelpMessage = 'The caption of the message box.'
            )]
            [ValidateNotNullOrEmpty()]
            [System.String]$Caption = 'Alert',
            [Parameter(
                       Position = 0,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Session'
            )]
            [Alias('SessionID')]
            [ValidateRange(0, 65536)]
            [System.Int32]$Id = -1,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesSession]$InputObject
        )

        begin
        {
            try
            {
                $FuncName = $MyInvocation.MyCommand
                Write-Verbose "[$funcName] Entering Begin block."

                if (!$ComputerName)
                {
                    Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                    $ComputerName = Get-TSGlobalServerName
                }
                else
                {
                    $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
                }

                Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
                $TSManager = New-Object Cassia.TerminalServicesManager
                $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
                $TSRemoteServer.Open()

                if (!$TSRemoteServer.IsOpen)
                {
                    Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
                }

                Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
                Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
                $null = Set-TSGlobalServerName -ComputerName $ComputerName
            }
            catch
            {
                Throw
            }
        }


        process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {

                if ($PSCmdlet.ParameterSetName -eq 'Session')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Id -ge 0)
                    {
                        $session = $TSRemoteServer.GetSession($Id)
                    }
                }

                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $InputObject
                }

                if ($session)
                {
                    Write-Verbose "[$FuncName] Sending alert message to session id: '$($session.SessionId)' on '$ComputerName'"
                    $session.MessageBox($Text, $Caption)
                }
            }
            catch
            {
                Throw
            }
        }


        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }
}
#endregion Source: Send-TSMessage.ps1

#region Source: functions\Set-DataGridView.ps1
function Invoke-Set-DataGridView_ps1
{
    function Set-DataGridView
    {
        <#
            .SYNOPSIS
                This function helps you edit the datagridview control

            .DESCRIPTION
                This function helps you edit the datagridview control

            .EXAMPLE
                Set-DataGridView -DataGridView $datagridview1 -ProperFormat -FontFamily $listboxFontFamily.Text -FontSize $listboxFontSize.Text

            .EXAMPLE
                Set-DataGridView -DataGridView $datagridview1 -AlternativeRowColor -BackColor 'AliceBlue' -ForeColor 'Black'

            .EXAMPLE
                Set-DataGridViewRowHeader -DataGridView $datagridview1 -HideRowHeader

            .EXAMPLE
                Set-DataGridViewRowHeader -DataGridView $datagridview1 -ShowRowHeader

        #>

        [CmdletBinding()]
        PARAM (
            [ValidateNotNull()]
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView,
            [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
            [Switch]$AlternativeRowColor,
            [Parameter(ParameterSetName = "DefaultRowColor")]
            [Switch]$DefaultRowColor,
            [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
            [Parameter(ParameterSetName = "DefaultRowColor")]
            [System.Drawing.Color]$ForeColor,
            [Parameter(Mandatory = $true, ParameterSetName = "AlternativeRowColor")]
            [Parameter(ParameterSetName = "DefaultRowColor")]
            [System.Drawing.Color]$BackColor,
            [Parameter(Mandatory = $true, ParameterSetName = "Proper")]
            [Switch]$ProperFormat,
            [Parameter(ParameterSetName = "Proper")]
            [String]$FontFamily = "Consolas",
            [Parameter(ParameterSetName = "Proper")]
            [Int]$FontSize = 10,
            [Parameter(ParameterSetName = "HideRowHeader")]
            [Switch]$HideRowHeader,
            [Parameter(ParameterSetName = "ShowRowHeader")]
            [Switch]$ShowRowHeader
        )
        PROCESS
        {
            if ($psboundparameters['AlternativeRowColor'])
            {
                $DataGridView.AlternatingRowsDefaultCellStyle.ForeColor = $ForeColor
                $DataGridView.AlternatingRowsDefaultCellStyle.BackColor = $BackColor
            }

            if ($psboundparameters['DefaultRowColor'])
            {
                $DataGridView.RowsDefaultCellStyle.ForeColor = $ForeColor
                $DataGridView.RowsDefaultCellStyle.BackColor = $BackColor
            }


            if ($psboundparameters['ProperFormat'])
            {
                #$Font = New-Object -TypeName System.Drawing.Font -ArgumentList "Segoi UI", 10
                $Font = New-Object -TypeName System.Drawing.Font -ArgumentList $FontFamily, $FontSize

                #[System.Drawing.FontStyle]::Bold

                $DataGridView.ColumnHeadersBorderStyle = 'Raised'
                $DataGridView.BorderStyle = 'Fixed3D'
                $DataGridView.SelectionMode = 'FullRowSelect'
                $DataGridView.AllowUserToResizeRows = $false
                $datagridview.DefaultCellStyle.font = $Font
            }

            if ($psboundparameters['HideRowHeader'])
            {
                $DataGridView.RowHeadersVisible = $false
            }
            if ($psboundparameters['ShowRowHeader'])
            {
                $DataGridView.RowHeadersVisible = $true
            }
        }

    } #Set-DataGridView
}
#endregion Source: Set-DataGridView.ps1

#region Source: functions\Set-DataGridViewFilter.ps1
function Invoke-Set-DataGridViewFilter_ps1
{
    function Set-DataGridViewFilter
    {
    <#
        .SYNOPSIS
            The function Set-DataGridViewFilter helps to only show specific entries with a specific value

        .DESCRIPTION
            The function Set-DataGridViewFilter helps to only show specific entries with a specific value.
            The data needs to be in a DataTable Object. You can use ConvertTo-DataTable to convert your
            PowerShell object into a DataTable object.

        .PARAMETER AllColumns
            Specifies to search all the column

        .PARAMETER ColumnName
            Specifies to search in a specific column name

        .PARAMETER DataGridView
            Specifies the DataGridView control where the data will be filtered

        .PARAMETER DataTable
            Specifies the DataTable object that is most likely the original source of the DataGridView

        .PARAMETER Filter
            Specifies the string to search

        .EXAMPLE
            PS C:\> Set-DataGridViewFilter -DataGridView $datagridview1 -DataTable $ProcessesDT -AllColumns -Filter $textbox1.Text

        .EXAMPLE
            PS C:\> Set-DataGridViewFilter -DataGridView $datagridview1 -DataTable $ProcessesDT -ColumnName "Name" -Filter $textbox1.Text

    #>
        PARAM (
            [Parameter(Mandatory = $true)]
            [System.Windows.Forms.DataGridView]$DataGridView,
            [Parameter(Mandatory = $true)]
            [System.Data.DataTable]$DataTable,
            [Parameter(Mandatory = $true)]
            [String]$Filter,
            [Parameter(Mandatory = $true, ParameterSetName = "OneColumn")]
            [String]$ColumnName,
            [Parameter(Mandatory = $true, ParameterSetName = "AllColumns")]
            [Switch]$AllColumns
        )
        PROCESS
        {
            $Filter = $Filter.ToString()

            IF ($PSBoundParameters['AllColumns'])
            {
                FOREACH ($Column in $DataTable.Columns)
                {
                    #$RowFilter += "Convert("+$($Column.ColumnName)+",'system.string') Like '%"{1}%' OR " -f $Column.ColumnName, $Filter
                    $RowFilter += "Convert($($Column.ColumnName),'system.string') Like '%$Filter%' OR "
                }

                # Remove the last 'OR'
                $RowFilter = $RowFilter -replace " OR $", ''

                #Append-RichtextboxStatus -Message $RowFilter
            }
            IF ($PSBoundParameters['ColumnName'])
            {
                $RowFilter = "$ColumnName LIKE '%$Filter%'"
            }

            $DataTable.defaultview.rowfilter = $RowFilter
            Load-DataGridView -DataGridView $DataGridView -Item $DataTable
        }
        END { Remove-Variable -Name $RowFilter -ErrorAction 'SilentlyContinue' | Out-Null }
    } #Set-DataGridViewFilter
}
#endregion Source: Set-DataGridViewFilter.ps1

#region Source: functions\Set-TextBox.ps1
function Invoke-Set-TextBox_ps1
{
    function Set-TextBox
    {
        [CmdletBinding()]
        PARAM (
            [System.Windows.Forms.TextBox]$TextBox,
            [System.Drawing.Color]$BackColor
        )
        BEGIN { }
        PROCESS
        {
            TRY
            {
                $TextBox.BackColor = $BackColor
            }
            CATCH { }
        }
    }
}
#endregion Source: Set-TextBox.ps1

#region Source: functions\Set-TSGlobalServerName.ps1
function Invoke-Set-TSGlobalServerName_ps1
{
    function Set-TSGlobalServerName
    {
        [CmdletBinding()]
        Param (
            [Parameter(Mandatory = $true)]
            [ValidateNotNullOrEmpty()]
            [System.String]$ComputerName
        )

        if ($ComputerName -eq "." -OR $ComputerName -eq $env:COMPUTERNAME)
        {
            $ComputerName = 'localhost'
        }

        $script:Server = $ComputerName
        $script:Server
    }

}
#endregion Source: Set-TSGlobalServerName.ps1

#region Source: functions\Stop-TSProcess.ps1
function Invoke-Stop-TSProcess_ps1
{
    function Stop-TSProcess
    {

        <#
        .SYNOPSIS
            Terminates the process running in a specific session or in all sessions.

        .DESCRIPTION
            Use Stop-TSProcess to terminate one or more processes from a local or remote computers.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Id
            Specifies the process Id number.

        .PARAMETER InputObject
            Specifies a process object. Enter a variable that contains the object, or type a command or expression that gets the sessions.

        .PARAMETER Name
            Specifies the process name.

        .PARAMETER Session
            Specifies the session Id number.

        .PARAMETER Force
               Overrides any confirmations made by the command.

        .EXAMPLE
             Get-TSProcess -Id 6552 | Stop-TSProcess

            Description
            -----------
            Gets process Id 6552 from the local computer and stop it. Confirmations needed.

        .EXAMPLE
            Get-TSSession -Id 3 -ComputerName comp1 | Stop-TSProcess -Force

            Description
            -----------
            Terminats all processes connected to session id 3 from remote computer 'comp1', suppress confirmations.

        .OUTPUTS
            Cassia.Impl.TerminalServicesProcess

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSProcess
            Get-TSSession
        #>

        [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Name')]
        Param (
            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server,
            [Parameter(
                       Position = 0,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Name'
            )]
            [Alias("ProcessName")]
            [System.String]$Name = '*',
            [Parameter(
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ValueFromPipelineByPropertyName = $true,
                       ParameterSetName = 'Id'
            )]
            [Alias('ProcessID')]
            [ValidateRange(0, 65536)]
            [System.Int32]$Id = -1,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesProcess]$InputObject,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'Session'
            )]
            [Alias('SessionId')]
            [Cassia.Impl.TerminalServicesSession]$Session,
            [switch]$Force
        )


        begin
        {
            try
            {
                $FuncName = $MyInvocation.MyCommand
                Write-Verbose "[$funcName] Entering Begin block."

                if (!$ComputerName)
                {
                    Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                    $ComputerName = Get-TSGlobalServerName
                }
                else
                {
                    $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
                }

                Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
                $TSManager = New-Object Cassia.TerminalServicesManager
                $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
                $TSRemoteServer.Open()

                if (!$TSRemoteServer.IsOpen)
                {
                    Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
                }

                Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
                Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
                $null = Set-TSGlobalServerName -ComputerName $ComputerName
            }
            catch
            {
                Throw
            }
        }


        Process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {

                if ($PSCmdlet.ParameterSetName -eq 'Name')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Name -eq '*')
                    {
                        $proc = $TSRemoteServer.GetProcesses()
                    }
                    else
                    {
                        $proc = $TSRemoteServer.GetProcesses() | Where-Object { $_.ProcessName -like $Name }
                    }
                }

                if ($PSCmdlet.ParameterSetName -eq 'Id')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Id -lt 0)
                    {
                        $proc = $TSRemoteServer.GetProcesses()
                    }
                    else
                    {
                        $proc = $TSRemoteServer.GetProcess($Id)
                    }
                }


                if ($PSCmdlet.ParameterSetName -eq 'Session')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    if ($Session)
                    {
                        $proc = $Session.GetProcesses()
                    }
                }


                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $proc = $InputObject
                }


                if ($proc)
                {
                    foreach ($p in $proc)
                    {
                        if ($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName, "Stop Process '$($p.ProcessName) ($($p.ProcessID))"))
                        {
                            Write-Verbose "[$FuncName] Killing process '$($p.ProcessName)' ($($p.ProcessId))"
                            $p.Kill()
                        }
                    }
                }
            }
            catch
            {
                Throw
            }
        }


        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }
}
#endregion Source: Stop-TSProcess.ps1

#region Source: functions\Stop-TSSession.ps1
function Invoke-Stop-TSSession_ps1
{
    function Stop-TSSession
    {

        <#
        .SYNOPSIS
            Logs the session off, disconnecting any user that might be connected.

        .DESCRIPTION
            Use Stop-TSSession to logoff the session and disconnect any user that might be connected.

        .PARAMETER ComputerName
                The name of the terminal server computer. The default is the local computer. Default value is the local computer (localhost).

        .PARAMETER Id
            Specifies the session Id number.

        .PARAMETER InputObject
               Specifies a session object. Enter a variable that contains the object, or type a command or expression that gets the sessions.

        .PARAMETER Synchronous
               When the Synchronous parameter is present the command waits until the session is fully disconnected otherwise it returns
               immediately, even though the session may not be completely disconnected yet.

        .PARAMETER Force
               Overrides any confirmations made by the command.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 | Stop-TSSession

            Description
            -----------
            logs off all connected users from Active sessions on remote computer 'comp1'. The caller is prompted to
            By default, the caller is prompted to confirm each action.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -State Active | Stop-TSSession -Force

            Description
            -----------
            logs off any connected user from Active sessions on remote computer 'comp1'.
            By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.

        .EXAMPLE
            Get-TSSession -ComputerName comp1 -State Active -Synchronous | Stop-TSSession -Force

            Description
            -----------
            logs off any connected user from Active sessions on remote computer 'comp1'. The Synchronous parameter tells the command to
            wait until the session is fully disconnected and only tghen it proceeds to the next session object.
            By default, the caller is prompted to confirm each action. To override confirmations, the Force Switch parameter is specified.

        .OUTPUTS

        .COMPONENT
            TerminalServer

        .NOTES
            Author: Shay Levy
            Blog  : http://blogs.microsoft.co.il/blogs/ScriptFanatic/

        .LINK
            http://code.msdn.microsoft.com/PSTerminalServices

        .LINK
            http://code.google.com/p/cassia/

        .LINK
            Get-TSSession
            Disconnect-TSSession
            Send-TSMessage
        #>

        [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Id')]
        Param (

            [Parameter()]
            [Alias('CN', 'IPAddress')]
            [System.String]$ComputerName = $script:server,
            [Parameter(
                       Position = 0,
                       Mandatory = $true,
                       ParameterSetName = 'Id',
                       ValueFromPipelineByPropertyName = $true
            )]
            [Alias('SessionId')]
            [System.Int32]$Id,
            [Parameter(
                       Mandatory = $true,
                       ValueFromPipeline = $true,
                       ParameterSetName = 'InputObject'
            )]
            [Cassia.Impl.TerminalServicesSession]$InputObject,
            [switch]$Synchronous,
            [switch]$Force
        )

        begin
        {
            try
            {
                $FuncName = $MyInvocation.MyCommand
                Write-Verbose "[$funcName] Entering Begin block."

                if (!$ComputerName)
                {
                    Write-Verbose "[$funcName] $ComputerName is not defined, loading global value '$script:Server'."
                    $ComputerName = Get-TSGlobalServerName
                }
                else
                {
                    $ComputerName = Set-TSGlobalServerName -ComputerName $ComputerName
                }

                Write-Verbose "[$FuncName] Attempting remote connection to '$ComputerName'"
                $TSManager = New-Object Cassia.TerminalServicesManager
                $TSRemoteServer = $TSManager.GetRemoteServer($ComputerName)
                $TSRemoteServer.Open()

                if (!$TSRemoteServer.IsOpen)
                {
                    Throw 'Connection to remote server is not open. Use Connect-TSServer to connect first.'
                }

                Write-Verbose "[$FuncName] Connection is open '$ComputerName'"
                Write-Verbose "[$FuncName] Updating global Server name '$ComputerName'"
                $null = Set-TSGlobalServerName -ComputerName $ComputerName
            }
            catch
            {
                Throw
            }
        }



        Process
        {

            Write-Verbose "[$funcName] Entering Process block."

            try
            {

                if ($PSCmdlet.ParameterSetName -eq 'Id')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $TSRemoteServer.GetSession($Id)
                }

                if ($PSCmdlet.ParameterSetName -eq 'InputObject')
                {
                    Write-Verbose "[$FuncName] Binding to ParameterSetName '$($PSCmdlet.ParameterSetName)'"
                    $session = $InputObject
                }

                if ($null -ne $session)
                {
                    if ($Force -or $PSCmdlet.ShouldProcess($TSRemoteServer.ServerName, "Logging off session id '$($session.sessionId)'"))
                    {
                        Write-Verbose "[$FuncName] Logging off session '$($session.SessionId)'"
                        $session.Logoff($Synchronous)
                    }
                }
            }
            catch
            {
                Throw
            }
        }


        end
        {
            try
            {
                Write-Verbose "[$funcName] Entering End block."
                Write-Verbose "[$funcName] Disconnecting from remote server '$($TSRemoteServer.ServerName)'"
                $TSRemoteServer.Close()
                $TSRemoteServer.Dispose()
            }
            catch
            {
                Throw
            }
        }
    }
}
#endregion Source: Stop-TSSession.ps1

#Start the application
Main ($CommandLine)
