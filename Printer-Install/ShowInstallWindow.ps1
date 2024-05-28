param (
    [Parameter(Mandatory=$true)]
    [string]$PrinterName
)

Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Printer Installation" Height="200" Width="400" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Image Grid.Row="0" Source="C:\ProgramData\FDS\PrinterInstall\FDSLogo.png" Margin="10"/>
        <TextBlock Grid.Row="1" Text="Installing printer: $PrinterName" HorizontalAlignment="Center" VerticalAlignment="Center" Margin="10"/>
        <ProgressBar Grid.Row="2" IsIndeterminate="True" Margin="10"/>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$window.ShowDialog() | Out-Null
