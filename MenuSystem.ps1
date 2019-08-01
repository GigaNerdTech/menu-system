# Written by Joshua Woleben
# Date 6/17/19
# Methodist Health System

$menu_config_path = "\\networkservers\Menus\"
$menu_header = "Name","Command","Description"

$global:current_menu_obects = New-Object -TypeName PSObject -Property @{Name='';Command='';Description=''}
$global:previous_item = "Main.menu"
$global:previous_name = "Main Menu"

function populate_menu {
    param($menu_object)

    $menu_object | ForEach-Object -Process {
         $MenuItems.Items.Add(($_ | Select -ExpandProperty 'Name')) | out-null  
    }
}

function load_config {
    param([string]$menu_config_path,
    [string]$menu_name)

    $menu_object = (Get-Content ($menu_config_path + $menu_name) | ConvertFrom-Csv -Header $menu_header)
    return ($menu_object | Sort-Object -Property Name)
}
function initialize {
    $main_menu_object = load_config $menu_config_path "Main.menu"
    populate_menu $main_menu_object 
    return $main_menu_object   
}
function ItemSelect {
    param([string]$command,
    [string]$name)

        if ($command -match "\.menu") {

        # Clear current items
        $MenuItems.Items.Clear()

        # Load new config
        $global:current_menu_objects = load_config $menu_config_path $command

        # Populate menu
        populate_menu $global:current_menu_objects

        # Set Menu Title
        $CurrentMenuLabel.Content = $name

      
    }
    else {
        # Launch command

        if ($command -match "http") {
            # Launch default browser
            Start $command

        }
        elseif ($command -match "\.exe") {
            Invoke-Command -FilePath ($command).ToString()
        }
        elseif ($command -match "\.xls") {
            $Excel = New-Object -ComObject Excel.Application
            $Workbook = $Excel.Workbooks.Open($command) 
            $Excel.Visible = $true
        }
        elseif($command -match "\.doc") {
            $Word= New-Object –comobject Word.Application

            $WordDoc = $Word.Documents.Open($command)
            $Word.Visible = $true
            
        }
        elseif($command -match "\.mp4") {
            $proc = Start-process -FilePath wmplayer.exe -ArgumentList $command
        }
        elseif ($command -match "\.ps1") {
            & $command
        }
        else {
            Invoke-Item $command
        }
    }

}

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Menu System" Height="800" Width="450" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="CurrentMenu" Content="MAIN"/>
        <ListBox x:Name="MenuItems" Height = "500" AllowDrop="True" SelectionMode="Extended" ScrollViewer.VerticalScrollBarVisibility="Visible"/>
        <Button x:Name="BackButton" Content="[ Previous Menu ]" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Label x:Name="Descrip" Content="Description"/>
        <Label x:Name="Description" MinHeight = "100"/>
        <Button x:Name="SelectItem" Content="[ Select Item ]" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}

# Set up controls
$CurrentMenuLabel = $global:Form.FindName('CurrentMenu')
$MenuItems = $global:Form.FindName('MenuItems')
$SelectItem = $global:Form.FindName('SelectItem')
$Description = $global:Form.FindName('Description')
$BackButton = $global:Form.FindName('BackButton')


# Populate Menu List
$global:current_menu_objects = initialize
$current_menu_obects | ForEach-Object -Process { Write-Host $_ }

$SelectItem.Add_Click({
    $current_item = New-Object -TypeName PSObject -Property @{Name='';Command='';Description=''}
    Write-Host $MenuItems.SelectedItem
    $current_item = $global:current_menu_objects | Where Name -eq $MenuItems.SelectedItem

    Write-Host "Button clicked."

    $current_item | ForEach-Object -Process { Write-Host $_ }
    
    ItemSelect $current_item.Command $current_item.Name

})

$MenuItems.Add_MouseDoubleClick({

    $current_item = New-Object -TypeName PSObject -Property @{Name='';Command='';Description=''}
    Write-Host $MenuItems.SelectedItem
    $current_item = $global:current_menu_objects | Where Name -eq $MenuItems.SelectedItem

    Write-Host "Button clicked."

    $current_item | ForEach-Object -Process { Write-Host $_ }
    
    ItemSelect $current_item.Command $current_item.Name


})
$MenuItems.Add_SelectionChanged({
     $current_item = New-Object -TypeName PSObject -Property @{Name='';Command='';Description=''}
    Write-Host $MenuItems.SelectedItem
    $current_item = $global:current_menu_objects | Where Name -eq $MenuItems.SelectedItem
    $Description.Content = $current_item.Description
    $Description.UpdateLayout()
    
})

$BackButton.Add_Click({
    ItemSelect $global:previous_item $global:previous_name
})


$global:Form.Add_Loaded({
    $global:Form.Title = "Menu System"
})
# Show GUI
$global:Form.ShowDialog() | out-null