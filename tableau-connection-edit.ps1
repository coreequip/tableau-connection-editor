[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Tableau Connection Editor" Height="250" Width="390" MinHeight="250" MinWidth="350"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,0,0" >
        <Grid.RowDefinitions>
            <RowDefinition Height="1*"/>
            <RowDefinition Height="50"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="120"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="1" Margin="5,10,10.4,10" Grid.RowSpan="2">
            <TextBox Margin="0,5,0,0" TextWrapping="Wrap" Text="{Binding schema}"   VerticalAlignment="Top" Height="25" IsEnabled="{Binding enabled}" />
            <ComboBox Margin="0,5,0,0" IsEditable="True" VerticalAlignment="Top" Text="{Binding server}" Height="25" ItemsSource="{Binding configs}" IsEnabled="{Binding enabled}" Name="cbServer" />
            <TextBox Margin="0,5,0,0" TextWrapping="Wrap" Text="{Binding port}"     VerticalAlignment="Top" Height="25" IsEnabled="{Binding enabled}" />
            <TextBox Margin="0,5,0,0" TextWrapping="Wrap" Text="{Binding service}"  VerticalAlignment="Top" Height="25" IsEnabled="{Binding enabled}" />
            <TextBox Margin="0,5,0,0" TextWrapping="Wrap" Text="{Binding username}" VerticalAlignment="Top" Height="25" IsEnabled="{Binding enabled}" />
        </StackPanel>
        <StackPanel Grid.Column="0" Margin="10,10,5,10" Grid.RowSpan="2" >
            <Label Content="Schema"   Margin="0,5,0,0" VerticalAlignment="Top" Height="25" />
            <Label Content="Server"   Margin="0,5,0,0" VerticalAlignment="Top" Height="25" />
            <Label Content="Port"     Margin="0,5,0,0" VerticalAlignment="Top" Height="25" />
            <Label Content="Service"  Margin="0,5,0,0" VerticalAlignment="Top" Height="25" />
            <Label Content="Username" Margin="0,5,0,0" VerticalAlignment="Top" Height="25" />
        </StackPanel>
        <StackPanel Grid.Column="1" Margin="0,0,0,0" Grid.Row="1" HorizontalAlignment="Right" VerticalAlignment="Center" Orientation="Horizontal">
            <Button Content="Open Workbook" Width="110" Margin="0,0,10,0" Name="btnOpen" />
            <Button Content="Save" Width="50" Margin="0,0,10,0" FontWeight="Bold" IsDefault="True" Name="btnSave" IsEnabled="{Binding enabled}" />
        </StackPanel>
        <TextBlock Grid.Column="0" Grid.Row="1" VerticalAlignment="Center" HorizontalAlignment="Center"><Hyperlink Name="link" NavigateUri="https://github.com/coreequip/tableau-connection-editor/issues" FontWeight="Bold">Support</Hyperlink></TextBlock>
    </Grid>
</Window>
"@

$version = '0.3 beta'

Write-Host ("`n Program v{0}, PowerShell v{1}" -f $version,$Host.Version.ToString(2))

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms


$form = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml)) 
$form.Title += ' v' + $version 

$viewModel = New-Object PSObject -Property @{
	schema = ""
	server = ""
	port   = ""
	service = ""
	username = ""
    configs = @()
    enabled = $false
}
$viewModel2 = New-Object psobject -Property @{
    configs = @()
    server = ""
}

$scriptPath = Split-Path $MyInvocation.MyCommand.Path

# Load serverconfigs
$csv = $null
$csvFile = Join-Path -Path $scriptPath -ChildPath 'tableau-connections.csv'
if (Test-Path -Path $csvFile) {
    $csv = Import-Csv -Path $csvFile -Delimiter ";"
    $viewModel.configs = @($csv.Server)
}

function saveConfig() {
    if ($csv -ne $null -and $csv.server.IndexOf($viewModel.server) -ge 0) {
        return
    }
    
    # PSCustomObject
    $entry = New-Object PSObject |
        Add-Member -Name Server  -Value $viewModel.server  -MemberType NoteProperty -PassThru |
        Add-Member -Name Port    -Value $viewModel.port    -MemberType NoteProperty -PassThru |
        Add-Member -Name Service -Value $viewModel.service -MemberType NoteProperty -PassThru

    if ($csv -eq $null) {
        $global:csv = @()
    } 

    $global:csv += $entry

    $csv | Export-Csv -Path $csvFile -Delimiter ';' -NoTypeInformation -ErrorAction SilentlyContinue
}


$btnSave = $form.FindName('btnSave')
$btnSave = $form.FindName('btnSave')

$form.FindName('btnOpen').add_Click({

    $ofd = New-Object Windows.Forms.OpenFileDialog
    $ofd.Filter = ' TWB - Tablaeu Workbook|*.twb'
    $res = $ofd.ShowDialog()
    if ($res -ne [windows.forms.DialogResult]::OK) { return }

    $global:filename = $ofd.FileName
    $strXml = Get-Content -Path $filename
    $global:xml = [xml] $strXml

    Write-Host (" - Opened file: '{0}'" -f $filename)
    Write-Host ("   Read bytes: {0}, XML is {1}" -f $strXml.Length, $xml.OuterXml.Length)

    $form.DataContext = $null

    $viewModel.enabled = $true
    @('schema'; 'server'; 'service'; 'username'; 'port') | % {
        $viewModel.$_ = $xml.SelectSingleNode('.//workbook/datasources/datasource[@name]/connection/@' + $_).'#text'
    }
    
    saveConfig

    $form.DataContext = $viewModel
})

$form.FindName('cbServer').add_SelectionChanged({
    $e = [Windows.Controls.SelectionChangedEventArgs]$_

    $global:x = $e.AddedItems
    $i = $csv.Server.IndexOf($e.AddedItems[0])

    if ($i -lt 0) {
        #$e.Handled = $true
        return
    }
    $global:cfg = $csv[$i]
    #$e.Handled = $true

    $viewModel2.configs = $viewModel.configs
    $viewModel2.server = $e.AddedItems[0]
    $form.DataContext = $viewModel2
    $viewModel.server = $e.AddedItems[0]
    $viewModel.port = $cfg.Port
    $viewModel.service = $cfg.Service
    $form.DataContext = $viewModel

})

$btnSave.Add_Click({
    $btnSave.IsEnabled = $false
    $xml.SelectNodes('/workbook/datasources/datasource[@name]/connection') | % {

        $node = $_

        @('schema'; 'server'; 'service'; 'username'; 'port') | % {
            $node.SelectSingleNode('@' + $_).'#text' = $viewModel.$_
        }

        foreach ($el in $node.SelectNodes('//relation[@type=''table'']')) {
            $el.table = '[' + $viewModel.schema + '].[' + $el.name + ']'
        }
    }

    $ms = new-object IO.MemoryStream
    $settings = new-object System.Xml.XmlWriterSettings
    $settings.CloseOutput = $true
    $settings.Indent = $true
    $settings.IndentChars = "`t"
    $settings.NewLineChars = "`r`n"
    $settings.Encoding = New-Object Text.UTF8Encoding $false # UTF8 w/o BOM
    $writer = [Xml.XmlWriter]::Create($ms, $settings)
    $xml.Save($writer)
    $writer.Flush()
    $writer.Close()

    Rename-Item -Force -Path $filename -NewName ($filename.Replace('.twb','.backup'))
    Set-Content -Force -Path $filename -Value ([Text.Encoding]::UTF8.GetString($ms.ToArray())) -Encoding UTF8
    
    [Windows.MessageBox]::Show('Saving successful.', 'Success', [Windows.MessageBoxButton]::OK, [Windows.MessageBoxImage]::Information) | Out-Null
    $btnSave.IsEnabled = $true
})

$form.FindName('link').Add_RequestNavigate({
    $e= [Windows.Navigation.RequestNavigateEventArgs] $_
    [System.Diagnostics.Process]::Start((new-object System.Diagnostics.ProcessStartInfo $e.Uri.AbsoluteUri))
    $e.Handled = $true
})

$form.DataContext = $viewModel
$form.ShowDialog() | Out-Null

#EOF
