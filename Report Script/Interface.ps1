Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic

#Layout of UI window displayed: 
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Email Scrapper" Height="500" Width="300" WindowStartupLocation="CenterScreen">
        <Grid Margin="10">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
              
                <Label Content = "Enter Start and End Date:" Grid.Row="0" HorizontalAlignment="Center"/>
                <DatePicker x:Name="startDate" Grid.Row="1" Padding="5" Margin="5"/>
                <DatePicker x:Name="endDate" Grid.Row="2" Padding="5" Margin="5"/>
                <Label x:Name="Progress" Grid.Row="3" Margin="5" HorizontalAlignment="Center"/>
                <Button x:Name="SubmitButton" Grid.Row="4" Margin="5" Height="30" Content="OK"/>

        </Grid>
</Window>
"@

#A Xml reader and loader for powershell so it can interprit the layout and render it:
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

#Alocate ID's for user input values 
$StartDate = $window.FindName("startDate")
$EndDate = $window.FindName("endDate")
$progressStats=$Window.FindName("Progress")
$SubmitButton = $window.FindName("SubmitButton")

#Specifies current path of application
$currentPath = Split-Path -Parent -Path $PSCommandPath

#Alocated the path where file will be saved
$path = Join-Path -Path $currentPath -ChildPath "savedDates.txt"

$SubmitButton.Add_Click({

	$startDate = $StartDate.SelectedDate
	$endDate = $EndDate.SelectedDate

    if($startDate -lt $endDate){
        $formattedStartDate = $startDate.ToString('yyyy-MM-dd')
        $formattedEndDate = $endDate.ToString('yyyy-MM-dd')
    
        $content = "$formattedStartDate `nto:$formattedEndDate"
        
        Set-Content -Path $path -Value $content
        Start-Process -FilePath "powershell.exe" -ArgumentList  "-File .\ReportScript.ps1"
        
        $progressStats.Content = "Retrieving emails: `nfrom: $content"
    }else{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Start date cannot be greater than End date", "OKOnly","Error With Date")
    }

})

$window.ShowDialog()

