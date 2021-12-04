################################################################################
#                                                                              #
#                                BULLET BLENDER                                #
#                       Written By: MSgt Anthony V. Brechtel                   #
#                                                                              #
################################################################################
clear-host
#Set-StrictMode -version latest
#Set-StrictMode -Off
$script:memBefore = (Get-Process -id $PID | Sort-Object WorkingSet64 | Select-Object Name,@{Name='WorkingSet';Expression={($_.WorkingSet64)}})
$script:memBefore = [System.Math]::Round(($script:memBefore.WorkingSet)/1mb, 2)
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Set-Location $dir
################################################################################
######Load Assemblies###########################################################
Add-Type -AssemblyName 'System.Windows.Forms'
Add-Type -AssemblyName 'System.Drawing'
Add-Type -AssemblyName 'PresentationFramework'
[System.Windows.Forms.Application]::EnableVisualStyles();

################################################################################
######Load Console Scaling Support##############################################
# Dummy WPF window (prevents auto scaling).
[xml]$Xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    x:Name="Window">
</Window>
"@
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml)
$Window = [Windows.Markup.XamlReader]::Load($Reader)
################################################################################
######Global Variables##########################################################

##System Vars
$script:program_title = "Bullet Blender"
$script:program_version = "1.4 (Beta - 4 Dec 2021)"
$script:settings = @{};                    #Contains System Settings
$script:return = 0;                        #Catches return from certain functions
$script:logfile = "$dir\Resources\Required\Log.txt"; if(Test-Path -literalpath $script:logfile){Remove-Item -literalpath $script:logfile}
$script:log_mem_change = $script:memBefore #Logs Difference In time
$script:print_to_console = 0;              #Turns on Console Log           1=On 0=Off
$script:print_to_log = 0;                  #Turns on File Log              1=On 0=Off
$script:var_size_detection = 0;            #Turns on Variable size tracker 1=On 0=Off
$script:var_sizes = New-Object system.collections.hashtable #Debugging Function

##Timer Vars
$script:reload_function = "";              #Forces a Reload independant of current window
$script:reload_function_arg_a = "";        #Arg A
$script:reload_function_arg_b = "";        #Arg B

##Display Vars
$script:zoom = 1;                          #Tracks users Zoom Level
$script:user_resizing = 0;                 #Tracks window Resizing
$script:user_resizing_starting_height = 0; #Tracks the most recent height prior to modification 
$script:Form_height = 1000                 #Default Window Height
$script:Form_width = 1480                  #Default Window Width
$script:sizer_box_width = 48               #Calculate Tex Size Box Width      
$script:Sidekick_width = 300               #Sidekick Window Width

##Editor Vars
$Script:recent_editor_text = "";           #Tracks Changes to Editor
$script:clicked_right = 0;                 #Signal to send left Click to move caret Position
$Script:LockInterface = 0;                 #Prevents Movement Timer Functions from Interupting caret Positions

##Location Vars
$Script:caret_position = 0;                #Tracks the Caret Position
$script:current_bullet = "";               #Tracks what bullet the  user is on
$script:current_line = 1                   #Tracks What line the user is on

##Dictionary Vars
$script:dictionary = @{};                  #The Dictionary
$script:dictionary_index = @{};            #Tracks status of correct/incorrect words & acronyms C = Correct, M = Misspelled, S = Shorthand Acro, E = Longhand Acro

##Acronym Vars
$script:acronym_list = New-Object system.collections.hashtable             #List of all acronyms loaded
$script:acro_index = @{};                  #Tracks the location & Information of all active acronyms
$Script:acronym_lists = @{};               #Tracks Acronym Lists Enabled/Disabled Status
$script:active_lists = @{};                #Remembers Which Lists the User is adding acronyms too

##Thesaurus Vars
$script:thesaurus_job = "";                #Tracks Job for Thesaurus lookups
$script:global_thesaurus = @{};            #Hand-off for Thesaurus
$script:thesaurus_menu = "";               #Hand-off menu to functions

##Word Hippo
$script:global_word_hippo =                #Hand-off for Word Hippo
$script:word_hippo_job = ""                #Tracks Job for Word Hippo lookups
$script:word_hippo_menu = "";               #Hand-off menu to functions

##Bullet Vars
$Script:bullet_banks = @{};                #Tracks Bullet Bank Lists Enabled/Disabled Status
$Script:bullet_bank = @{};                 #All Currently Loaded Bullets for Feeder

##Package Vars
$script:package_list = @{};                #Tracks packages Enabled/Disabled Status

##Feeder Vars
$script:feeder_job = "";                   #Tracks Job for bullet Feeder

##Theme Settings
$script:theme_settings = @{};              #Contains Theme Colors
$script:theme_original = @{};              #Tracks any changes made during theme changes
$script:color_picker = "";                 #Transfers color to Colorpicker

##Sidekick Vars
$script:sidekick_job = "";                 #Tracks the job for Sidekick metrics/activity
$script:sidekick_results = "";             #Retrieves Sidekick Job Results
$script:sidekickgui = "New";               #Tracks Rebuild of Sidekick Window
$script:package_name_value = "";           #Form Object Tracks Right Pane Package Name
$script:bullets_loaded_name_value = "";    #Form Object Tracks Right Pane Loaded Bullet Count
$script:acronyms_loaded_name_value = "";   #Form Object Tracks Right Pane Acronym Count
$script:location_value = "";               #Form Object Tracks Right Pane Users Current Line
$script:headers_count_name_value = "";     #Form Object Tracks Right Pane Header Count
$script:bullets_count_name_value = "";     #Form Object Tracks Right Pane Package Bullet Count
$script:word_count_name_value = "";        #Form Object Tracks Right Pane Word Count
$script:unique_acro_count_name_value = ""; #Form Object Tracks Right Pane Unique Acronyms
$script:formating_errors_combo = "";       #Form Object Tracks Right Pane Formatting Errors
$format_error = "";            #Hash Containing Formatting Errors
$script:consistency_errors_combo = "";     #Form Object Tracks Right Pane Consistency Errors
$script:top_used_acros_combo = "";         #Form Object Tracks Right Pane Top Used Acronyms
$script:top_used_words_combo = "";         #Form Object Tracks Right Pane Top Used Words
$script:metrics_used_combo = "";           #Form Object Tracks Right Pane Metrics Used
$script:compression_trackbar_label = "";   #Form Object Displays Trackbar Setting

##History Vars
$script:old_text = "";                     #Used for tracking changes between new editor text and previous text
$script:text_lock = -1;                    #Locks the position of users movement between Undo/Redo actions
$script:history = @{};                     #Contains the history of all keypresses
$script:lock_history = 0;                  #Prevents changes to history during package loads
$script:history_system_location = 0;       #Latest entry in history (Always last)
$script:history_user_location = 0;         #Where the user is during Undo/Redo movements
$script:history_replace_text = "";         #Contents of selected text tracks overwritten text
$script:history_replace_text_start = 0     #Position Start of selected text during text overwrites
$script:history_replace_text_end = 0       #Lengh of selected text during text overwrites
$script:save_history_job = "";             #Tracks the job of saving history
$script:save_history_timer = Get-Date      #Tracks the last time a job was ran
$script:save_history_tracker = 0;          #Tracks the last place in history that was saved

##Calculate Text Size Vars
$script:bullets_and_sizes  = new-object System.Collections.Hashtable #Tracks current bullet sizes
$script:bullets_and_lines  = new-object System.Collections.Hashtable #Tracks location of bullets
$script:bullets_compressed = new-object System.Collections.Hashtable #Tracks whether a bullet has already been compressed
$script:space_hash         = new-object System.Collections.Hashtable #Text Compression characters
$character_blocks          = new-object System.Collections.Hashtable #Sizes list for all characters

##Idle Timer
if(Test-Path variable:Script:Timer){$Script:Timer.Dispose();}
$Script:Timer = New-Object System.Windows.Forms.Timer                #Main system timer, most functions load through this timer
$Script:Timer.Interval = 1000
$Script:CountDown = 1

##Variable Flushing
$script:main_vars = @()                                                 #Contains List of Startup Variables


################################################################################
######Main######################################################################
function main
{
    #################################################################################
    ##Build Main Form
    #$script:Form                                           = New-Object system.Windows.Forms.Form   
    $script:Form.Size                                      = New-Object Drawing.Size(1475, 1000)
    $script:Form.ClientSize.Width                          = 2800
    $script:Form.ClientSize.Height                         = 1000
    $script:Form.BackColor                                 = $script:theme_settings['MAIN_BACKGROUND_COLOR']
    $script:Form.Text                                      = "$script:program_title"
    $script:Form.MinimumSize                               = New-Object Drawing.Size(1265,200)
    $script:Form.keypreview                                = $true
    $script:Form.Add_KeyDown({
        #################################################################################
        ##Override Key System Key Presses (Ctrl-Z, Ctrl-Y,Ctrl-Left,Ctrl-Right)
        if(($_.Control) -and ($_.keycode -match "z"))
        {
            #write-host "Undo"
            $_.handled = $true
            if($script:text_lock -eq -1)
            {
                $script:text_lock = $script:history_system_location;
            }
            undo_history
        }
        if(($_.Control) -and ($_.keycode -match "^y"))
        {
            #write-host "Redo"
            $_.handled = $true
            if($script:text_lock -eq -1)
            {
                $script:text_lock = $script:history_system_location;
            }
            redo_history
        }
        if(($_.Control) -and ($_.keycode -match "RIGHT"))
        {
            $pattern = " "
            $matches = [regex]::Matches($editor.text, $pattern)
            if($matches.Success)
            {  
                foreach($match in $matches)
                {
                    if($match.index -gt $editor.SelectionStart)
                    {
                        $editor.SelectionStart = $match.index + 1;
                        break
                    }
                }
            }
            $_.handled = $true
        }
        if(($_.Control) -and ($_.keycode -match "LEFT"))
        {
            $pattern = " "
            $matches = [regex]::Matches($editor.text, $pattern)
            if($matches.Success)
            {  
                $final = 0;
                foreach($match in $matches)
                {   
                    if(($match.index + 1) -ge $editor.SelectionStart)
                    {
                        break
                    }
                    $final = $match.index
                }
                $editor.SelectionStart = $final + 1;
            }
            $_.handled = $true
        }
    })
    #################################################################################
    ##Make Gradient Background
    $script:Form.add_paint({
        if($Form.WindowState -ne "Minimized")
        {
            $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point $this.clientrectangle.width,0),
            (new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),$script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'],$script:theme_settings['MAIN_BACKGROUND_COLOR'])
            $_.graphics.fillrectangle($brush,$this.clientrectangle)
        }
        
    })

    #################################################################################
    ##Build System Menu
    #$MenuBar                                        = New-Object System.Windows.Forms.MenuStrip
    $MenuBar.backcolor                              = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $MenuBar.forecolor                              = $script:theme_settings['MENU_TEXT_COLOR']
    $script:Form.Controls.Add($MenuBar)
    $FileMenu                                       = New-Object System.Windows.Forms.ToolStripMenuItem
    $EditMenu                                       = New-Object System.Windows.Forms.ToolStripMenuItem
    $BulletMenu                                     = New-Object System.Windows.Forms.ToolStripMenuItem
    $script:AcronymMenu                             = New-Object System.Windows.Forms.ToolStripMenuItem
    $OptionsMenu                                    = New-Object System.Windows.Forms.ToolStripMenuItem
    $AboutMenu                                      = New-Object System.Windows.Forms.ToolStripMenuItem
    $FileMenu.Font                                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $EditMenu.Font                                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $BulletMenu.Font                                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $script:AcronymMenu.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $OptionsMenu.Font                               = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $AboutMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $MenuBar.Items.Add($FileMenu) | Out-Null
    $MenuBar.Items.Add($EditMenu) | Out-Null
    $MenuBar.Items.Add($BulletMenu) | Out-Null
    $MenuBar.Items.Add($script:AcronymMenu) | Out-Null
    $MenuBar.Items.Add($OptionsMenu) | Out-Null
    $MenuBar.Items.Add($AboutMenu) | Out-Null
    $FileMenu.Text = "&File"
    $EditMenu.Text = "&Edit"      
    $BulletMenu.Text = "&Bullets"
    $script:AcronymMenu.Text = "&Acronyms"
    $OptionsMenu.Text = "&Options"
    $AboutMenu.Text = "&About"
       
    #################################################################################
    ##Build File Menu
    build_file_menu

    #################################################################################
    ##Edit Menu
    build_edit_menu

    #################################################################################
    ##Bullet Menu
    build_bullet_menu
    
    #################################################################################
    ##Acronym Menu
    build_acronym_menu

    #################################################################################
    ##Options Menu
    build_options_menu

    #################################################################################
    ##Options Menu
    build_about_menu
    
    #################################################################################
    ##Editor
    #$editor                                         = New-Object CustomRichTextBox
    $editor.Font                                    = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    load_package
    $editor.Font                                    = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    $editor.Size                                    = New-Object System.Drawing.Size(1200,650)
    $editor.Location                                = New-Object System.Drawing.Size(5,40)    
    $editor.ReadOnly                                = $False
    $editor.WordWrap                                = $False
    $editor.Multiline                               = $True
    $editor.BackColor                               = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
    $editor.ScrollBars                              = "vertical"

    #####Force Scaling
    #$editor.SelectionStart                        = 0
    #$editor.SelectionLength                       = 1
    #$editor.SelectionColor                        = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_FONT_COLOR'])
    #$editor.Font                                  = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    #$editor.DeselectAll()

    #################################################################################
    ##Editor Text Change
    $editor.Add_TextChanged({
        if($script:lock_history -ne 1)
        {
            write_history
        }
        $this.CustomVScroll()
        #Reset Idle Timer
        $Script:CountDown = 3         
    })
    
    #################################################################################
    ##Editor Right Mouse Click
    $editor.Add_MouseDown({  
        if($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) 
        {
            if($editor.SelectedText.length -eq 0)
            {
                [Clicker]::LeftClickAtPoint([System.Windows.Forms.Cursor]::Position.X,[System.Windows.Forms.Cursor]::Position.Y) #Sends Left Click to move Caret
                $script:clicked_right = 1;
            }
            else
            {
                if($Script:LockInterface -ne 1)
                {
                    right_click_menu
                }
            }
        }
        #################################################################################
        ##Editor Left Mouse Click
        if($_.Button -eq [System.Windows.Forms.MouseButtons]::Left ) 
        {  
            if($script:clicked_right -eq 1)
            {
                if($Script:LockInterface -ne 1)
                {
                    right_click_menu
                    $script:clicked_right = 0
                }
            }
        }
    })

    #################################################################################
    ##Editor Mouse Up
    $editor.Add_MouseUp({
        if($_.Button -eq [System.Windows.Forms.MouseButtons]::Left ) 
        {
            $script:history_replace_text = $editor.SelectedText
            $script:history_replace_text_start = $editor.SelectionStart
            $script:history_replace_text_end = $editor.Selectionlength
        }
    })

    #################################################################################
    ##Editor KeyDown 
    $editor.Add_Keydown({
        #Protects Text from Undo/Redo Actions
        $script:text_lock = -1
    })
    $script:Form.Controls.Add($editor)

    #################################################################################
    ##Ghost Editor
    #$ghost_editor                                   = New-Object CustomRichTextBox
    $ghost_editor.Font                              = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    $ghost_editor.Size                              = New-Object System.Drawing.Size(1200,650)
    $ghost_editor.Location                          = New-Object System.Drawing.Size(($editor.width + 50),40)
    $ghost_editor.WordWrap                          = $false
    $ghost_editor.Multiline                         = $true

    #################################################################################
    ##Bullet Feed Panel
    #$bullet_feeder_panel                     = New-Object system.Windows.Forms.Panel
    $bullet_feeder_panel.height              = ($script:Form.height - $editor.height - 100)
    $bullet_feeder_panel.width               = $editor.Width
    $bullet_feeder_panel.BackColor           = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
    $bullet_feeder_panel.Anchor              = 'top,left'
    $bullet_feeder_panel.Location            = New-Object System.Drawing.Point($editor.Location.x,($editor.Height + $editor.Location.y + 2))
    $bullet_feeder_panel.Add_MouseDown({

        $script:user_resizing = [System.Windows.Forms.Cursor]::Position.Y;
        $script:user_resizing_starting_height = $editor.height
        #write-host Moving Started

    })
    $bullet_feeder_panel.Add_Mouseup({

        $script:user_resizing = 0;
        $Script:Timer.Interval = 100;
        #write-host Moving Ended

    })

    #################################################################################
    ##Bullet Feeder Box
    #$feeder_box                                     = New-Object System.Windows.Forms.RichTextBox
    $feeder_box.Text                                = ""
    $feeder_box.Font                                = [Drawing.Font]::New($script:theme_settings['FEEDER_FONT'], [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'])
    $feeder_box.Size                                = New-Object System.Drawing.Size($bullet_feeder_panel.width,($bullet_feeder_panel.height - 5))
    $feeder_box.Location                            = New-Object System.Drawing.Size(0,5)    
    $feeder_box.ReadOnly                            = $True
    $feeder_box.WordWrap                            = $True
    $feeder_box.Multiline                           = $True
    $feeder_box.BackColor                           = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
    $feeder_Box.ForeColor                           = $script:theme_settings['FEEDER_FONT_COLOR']
    $feeder_box.ScrollBars                          = "Vertical"
    $feeder_box.Add_MouseDown({     
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) 
        {
            if($feeder_box.SelectedText.Length -ge 1)
            {
                $contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
                $feeder_box.ContextMenuStrip = $contextMenuStrip1        
                $contextMenuStrip1.Items.Add("Copy").add_Click({clipboard_copy_2})
            }
        }
    })
    $feeder_box.Add_KeyUp({
        if(($_.control) -and ($_.keycode -match "c"))
        {
            #write-host Copy Feeder
            clipboard_copy_2
        }
    })
    $bullet_feeder_panel.Controls.Add($feeder_box)

    #################################################################################
    ##Text Sizer
    #$sizer_box                                      = New-Object System.Windows.Forms.RichTextBox
    $sizer_box.Size                                 = New-Object System.Drawing.Size($script:sizer_box_width,($editor.Height - 4))
    $sizer_box.Location                             = New-Object System.Drawing.Size(($editor.Location.x + $editor.width),($editor.Location.y + 3))    
    $sizer_box.ReadOnly                             = $true
    $sizer_box.SelectionAlignment                   = "Right"
    $sizer_box.WordWrap                             = $false
    $sizer_box.ScrollBars                           = 'None'
    $sizer_box.Multiline                            = $true
    $sizer_box.BorderStyle                          = "none"
    $sizer_box.Forecolor                            = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'])
    $sizer_box.BackColor                            = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
    $sizer_box.Font                                 = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE']) 
    $sizer_box.text = " "
    $sizer_box.Font                                 = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])

    #################################################################################
    ##Sizer Art
    #$sizer_art                                      = new-object system.windows.forms.label
    $sizer_art.Width                                = $sizer_box.width
    $sizer_art.height                               = $feeder_box.height + 5
    $sizer_art.Location                             = New-Object System.Drawing.Size(($feeder_box.location.x + $feeder_box.width),($sizer_box.Location.y + $sizer_box.height))  
    $sizer_art.add_paint({
        $brush = new-object System.Drawing.Drawing2D.LinearGradientBrush((new-object system.drawing.point 0,0),
        (new-object system.drawing.point($this.clientrectangle.width,$this.clientrectangle.height)),$script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'],$script:theme_settings['MAIN_BACKGROUND_COLOR'])
        $_.graphics.fillrectangle($brush,$this.clientrectangle)
    })

    #################################################################################
    ##Sidekick Console
    #$sidekick_panel                          = New-Object system.Windows.Forms.Panel
    $sidekick_panel.height                   = ($script:Form.Height - 100)
    $sidekick_panel.width                    = $script:Sidekick_width
    $sidekick_panel.BackColor                = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
    $sidekick_panel.Anchor                   = 'top,left'
    $sidekick_panel.Location                 = New-Object System.Drawing.Point(($sizer_box.Location.x + $sizer_box.width + 10),$editor.Location.y)
    $sidekick_panel.Add_MouseDown({
        if($sidekick_panel.width -eq 5)
        {
            $sidekick_panel.width = $script:Sidekick_width;
            $script:Form_Width = ($script:Form_Width + 1); #Force a Shrink change
        }
        else
        {
            $sidekick_panel.width = 5;
            $script:Form_Width = ($script:Form_Width - 1); #Force a Grow change
        }
        $script:zoom = "Changed"
    })
    $script:Form.Controls.Add($sidekick_panel)

    #$left_panel                             = New-Object system.Windows.Forms.Panel
    $left_panel.height                       = ($sidekick_panel.height)
    $left_panel.width                        = ($sidekick_panel.width - 5)
    $left_panel.BackColor                    = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
    $left_panel.Anchor                       = 'top,left'
    $left_panel.Location                     = New-Object System.Drawing.Point(5,-5) #Modified Y pos to -5 from 0 (Ver 1.3 Update)
    $sidekick_panel.Controls.Add($left_panel)
    
    #################################################################################
    ##Mirror Scrolling System
    $editor.Buddy = $sizer_box
    $editor.Add_VScroll({
        $this.CustomVScroll()
    })
    #################################################################################
    ##Finalize Form
    $script:Form.controls.Add($bullet_feeder_panel)
    $script:Form.Controls.Add($sizer_art)
    $script:Form.Controls.Add($sizer_box)
    $script:Form.Controls.Add($sizer_box)
    
    Log "Main End"
    Log "BLANK"
    $script:Form.ShowDialog()
}
################################################################################
#####Build File Menu############################################################
function build_file_menu
{
    $FileMenu.DropDownItems.clear();
    $new_package = New-Object System.Windows.Forms.ToolStripMenuItem
    $new_package.Text = "New"
    $new_package.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $new_package.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $FileMenu.DropDownItems.Add($new_package)  | Out-Null
    $new_package.Add_Click({
        $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        $Script:Timer.Stop()
        save_history
        $script:return = 0;
        if(($editor.text -ne "") -and ($script:settings['PACKAGE'] -eq "Current"))
        {
            $message = "Would you like to save your current package first?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Save Current Work?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                $script:return = save_package_dialog
            }
            else
            {
                if(test-path -literalpath "$dir\Resources\Packages\Current")
                {
                    Remove-Item -literalpath "$dir\Resources\Packages\Current" -recurse -Force
                }
                $script:settings['PACKAGE'] = "Current"
                load_package
            }
        }
        else
        {
            $script:return = 1;
        }
        if($script:return -eq 1)
        {
            if(test-path -literalpath "$dir\Resources\Packages\Current")
            {
                Remove-Item -literalpath "$dir\Resources\Packages\Current" -recurse -Force
            }
            $script:settings['PACKAGE'] = "Current"
            load_package
        }
        
    })
    $save_package = New-Object System.Windows.Forms.ToolStripMenuItem
    $save_package.Text = "Save"
    $save_package.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $save_package.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $FileMenu.DropDownItems.Add($Save_package) | Out-Null
    $Save_package.Add_Click({
        $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        if($script:settings['PACKAGE'] -eq "Current")
        {
            save_package_dialog
        }
        else
        {
            save_history
        }
    })

    $saveas_package = New-Object System.Windows.Forms.ToolStripMenuItem
    $saveas_package.Text = "Save As"
    $saveas_package.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $saveas_package.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']

    $FileMenu.DropDownItems.Add($Saveas_package) | Out-Null
    $Saveas_package.Add_Click({
        $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        save_package_dialog
    })

    $open_package = New-Object System.Windows.Forms.ToolStripMenuItem
    $open_package.Text = "Open / Manage"
    $open_package.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $open_package.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $FileMenu.DropDownItems.Add($open_package) | Out-Null
    $open_package.Add_Click({
        $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        manage_package_dialog
    })
}
################################################################################
#####Build Edit Menu############################################################
function build_edit_menu
{
    $EditMenu.DropDownItems.clear();
    $undo_option = New-Object System.Windows.Forms.ToolStripMenuItem
    $undo_option.Text = "Undo         (Ctrl + Z)"
    $undo_option.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $undo_option.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $EditMenu.DropDownItems.Add($undo_option)  | Out-Null
    $undo_option.Add_Click({
        if($script:text_lock -eq -1)
        {
            $script:text_lock = $script:history_system_location;
        }
        undo_history
    })
    $redo_option = New-Object System.Windows.Forms.ToolStripMenuItem
    $redo_option.Text = "Redo         (Ctrl + Y)"
    $redo_option.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $redo_option.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $EditMenu.DropDownItems.Add($redo_option)  | Out-Null
    $redo_option.Add_Click({
        if($script:text_lock -eq -1)
        {
            $script:text_lock = $script:history_system_location;
        }
        redo_history
    })
}
################################################################################
#####Build Bullet Menu##########################################################
function build_bullet_menu
{
    $BulletMenu.DropDownItems.clear();
    $manage_bullets = New-Object System.Windows.Forms.ToolStripMenuItem
    $manage_bullets.Text = "Manage Bullet Banks"
    $manage_bullets.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $manage_bullets.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $BulletMenu.DropDownItems.Add($manage_bullets) | Out-Null
    $manage_bullets.Add_Click({
        $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        manage_bullets_dialog
    })

    $import_bullets = New-Object System.Windows.Forms.ToolStripMenuItem
    $import_bullets.Text = "Import Bullet Bank"
    $import_bullets.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $import_bullets.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $BulletMenu.DropDownItems.Add($import_bullets) | Out-Null
    $import_bullets.Add_Click({
        $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        import_bullet_form
    })


    $separator = [System.Windows.Forms.ToolStripSeparator]::new()
	$BulletMenu.DropDownItems.Add($separator) | Out-Null

    if($script:package_list.get_Count() -gt 1)
    {      
        $internal_packages = New-Object System.Windows.Forms.ToolStripMenuItem
        $internal_packages.Text = "Internal Packages"
        $internal_packages.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $internal_packages.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        if($script:settings['LOAD_PACKAGES_AS_BULLETS'] -eq 1)
        {
            $internal_packages.checked = $true
        }
        else
        {
            $internal_packages.checked = $false
        }
        $internal_packages.Add_Click({
            $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $MenuBar.Refresh();
            if($this.Checked -eq $true)
            {
                $this.Checked = $false
                $script:settings['LOAD_PACKAGES_AS_BULLETS'] = 0;
            }
            else
            {
                $this.Checked = $true
                $script:settings['LOAD_PACKAGES_AS_BULLETS'] = 1;
            }
            update_settings
            load_bullets
        })
	    $BulletMenu.DropDownItems.Add($internal_packages) | Out-Null
    }

    if($script:Bullet_banks.count -ne 0)
    {
        foreach($bank in $script:Bullet_banks.getEnumerator() | Sort Key)
        {
            $bullet_list = New-Object System.Windows.Forms.ToolStripMenuItem
            $bullet_list.Text = $bank.key -replace ".txt$",""
            $bullet_list.Name = $bank.key
            $bullet_list.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $bullet_list.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']

            if($bank.value -eq 1)
            {
                $bullet_list.checked = $true
            }
            else
            {
                $bullet_list.checked = $false
            }
            $BulletMenu.DropDownItems.Add($bullet_list) | Out-Null
            $bullet_list.Add_Click({
                $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $MenuBar.Refresh();
                if($this.Checked -eq $true)
                {
                    $this.Checked = $false
                    $script:Bullet_banks[$this.name] = 0;
                }
                else
                {
                    $this.Checked = $true
                    $script:Bullet_banks[$this.name] = 1;
                }
                save_bullet_tracker
                load_bullets
            })
        }
    }
    save_bullet_tracker
    load_bullets
    $Script:recent_editor_text = "Changed"
}
################################################################################
#####Build Acronym Menu#########################################################
function build_acronym_menu
{
    
    $script:AcronymMenu.DropDownItems.clear();
    
    $manage_acronym_list = New-Object System.Windows.Forms.ToolStripMenuItem
    $manage_acronym_list.Text = "Manage Acronyms && Abbreviations"
    $manage_acronym_list.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $manage_acronym_list.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $script:AcronymMenu.DropDownItems.Add($manage_acronym_list) | Out-Null
    $manage_acronym_list.Add_Click({
        $script:AcronymMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        manage_acronyms_dialog
    })

    $import_acronym_list = New-Object System.Windows.Forms.ToolStripMenuItem
    $import_acronym_list.Text = "Import Acronyms && Abbreviations List"
    $import_acronym_list.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $import_acronym_list.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $script:AcronymMenu.DropDownItems.Add($import_acronym_list) | Out-Null
    $import_acronym_list.Add_Click({
        $script:AcronymMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        import_acronym_form
    })

    $separator = [System.Windows.Forms.ToolStripSeparator]::new()
    $script:AcronymMenu.DropDownItems.Add($separator) | Out-Null

    if($Script:acronym_lists.count -ne 0)
    {   
        foreach($list in $Script:acronym_lists.getEnumerator() | Sort Key)
        {
            $acronym_list = New-Object System.Windows.Forms.ToolStripMenuItem
            $acronym_list.Text = $list.key -replace ".csv$",""
            $acronym_list.Name = $list.key
            $acronym_list.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $acronym_list.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            if($list.value -eq 1)
            {
                $acronym_list.checked = $true
            }
            else
            {
                $acronym_list.checked = $false
            }
            $script:AcronymMenu.DropDownItems.Add($acronym_list) | Out-Null
            $acronym_list.Add_Click({
                $script:AcronymMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $MenuBar.Refresh();
                if($this.Checked -eq $true)
                {
                    $this.Checked = $false
                    $Script:acronym_lists[$this.name] = 0;
                }
                else
                {
                    $this.Checked = $true
                    $Script:acronym_lists[$this.name] = 1;
                }
                
                save_acronym_tracker
                load_acronyms
                $Script:recent_editor_text = "Changed"
            })
        }
    }
    save_acronym_tracker
    load_acronyms
    $Script:recent_editor_text = "Changed"
}
################################################################################
#####Build Options Menu#########################################################
function build_options_menu
{
    $OptionsMenu.DropDownItems.clear();
    $theme_settings = New-Object System.Windows.Forms.ToolStripMenuItem
    $theme_settings.Text = "Theme Settings"
    $theme_settings.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $theme_settings.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $OptionsMenu.DropDownItems.Add($theme_settings) | Out-Null
    $theme_settings.Add_Click({
        $OptionsMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        interface_dialog
    })

    $system_settings = New-Object System.Windows.Forms.ToolStripMenuItem
    $system_settings.Text = "System Settings"
    $system_settings.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $system_settings.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $OptionsMenu.DropDownItems.Add($system_settings) | Out-Null
    $system_settings.Add_Click({
        $OptionsMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.Refresh();
        system_settings_dialog
    })
}
################################################################################
#####Build About Menu###########################################################
function build_about_menu
{
    $AboutMenu.DropDownItems.clear();
    $FAQ = New-Object System.Windows.Forms.ToolStripMenuItem
    $FAQ.Text = "FAQ"
    $FAQ.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $FAQ.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $AboutMenu.DropDownItems.Add($FAQ) | Out-Null
    $FAQ.Add_Click({
        FAQ_dialog
    })

    $about = New-Object System.Windows.Forms.ToolStripMenuItem
    $about.Text = "About"
    $about.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $about.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    $AboutMenu.DropDownItems.Add($about) | Out-Null
    $about.Add_Click({
        about_dialog
    }) 

}
################################################################################
######Manage Bullets Dialog#####################################################
function manage_bullets_dialog
{
    
    $item_number = $script:Bullet_banks.get_count()
    #$item_number = 20
    
    $spacer = 0;
    $edit_bullets_form = New-Object System.Windows.Forms.Form
    $edit_bullets_form.FormBorderStyle = 'Fixed3D'
    $edit_bullets_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $edit_bullets_form.Location = new-object System.Drawing.Point(0, 0)
    $edit_bullets_form.MaximizeBox = $false
    $edit_bullets_form.SizeGripStyle = "Hide"
    $edit_bullets_form.Width = 800
    if($item_number -eq 0)
    {
        $edit_bullets_form.Height = 200;
    }
    elseif((($item_number * 65) + 140) -ge 600)
    {
        $edit_bullets_form.Height = 600;
        $edit_bullets_form.Autoscroll = $true
        $spacer = 20
    }
    else
    {
        $edit_bullets_form.Height = (($item_number * 65) + 140)
    }
    $edit_bullets_form.Text = "Manage Bullet Banks"
    #$edit_bullets_form.TopMost = $True
    $edit_bullets_form.TabIndex = 0
    $edit_bullets_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    ################################################################################################
    $y_pos = 10;


    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "Manage Bullet Banks";
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($edit_bullets_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point((($edit_bullets_form.width / 2) - ($title_label.width / 2)),$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $edit_bullets_form.controls.Add($title_label);

    $y_pos = $y_pos + 40;
    $create_bank_button           = New-Object System.Windows.Forms.Button
    $create_bank_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $create_bank_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $create_bank_button.Width     = 200
    $create_bank_button.height     = 25
    $create_bank_button.Location  = New-Object System.Drawing.Point((($edit_bullets_form.width / 2) - $create_bank_button.width - 10),$y_pos);
    $create_bank_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $create_bank_button.Text      ="Create Bullet Bank"
    $create_bank_button.Name = ""
    $create_bank_button.Add_Click({
        create_bullet_bank
        $script:reload_function = "manage_bullets_dialog" 
        $edit_bullets_form.close();
    })
    $edit_bullets_form.controls.Add($create_bank_button)


    
    $import_bank_button           = New-Object System.Windows.Forms.Button
    $import_bank_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $import_bank_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $import_bank_button.Width     = 200
    $import_bank_button.height     = 25
    $import_bank_button.Location  = New-Object System.Drawing.Point((($edit_bullets_form.width / 2) + 10),$y_pos);
    $import_bank_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $import_bank_button.Text      ="Import Bullet Bank"
    $import_bank_button.Name = ""
    $import_bank_button.Add_Click({ 
        import_bullet_form
        $script:reload_function = "manage_bullets_dialog" 
        $edit_bullets_form.close();
        
    })
    $edit_bullets_form.controls.Add($import_bank_button)
    

    $y_pos = $y_pos + 35;
    $separator_bar                             = New-Object system.Windows.Forms.Label
    $separator_bar.text                        = ""
    $separator_bar.AutoSize                    = $false
    $separator_bar.BorderStyle                 = "fixed3d"
    #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
    $separator_bar.Anchor                      = 'top,left'
    $separator_bar.width                       = (($edit_bullets_form.width - 50) - $spacer)
    $separator_bar.height                      = 1
    $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
    $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $separator_bar.TextAlign                   = 'MiddleLeft'
    $edit_bullets_form.controls.Add($separator_bar);

    $y_pos = $y_pos + 5;

    #write-host Header $y_pos

    if($item_number -ne 0)
    {
        #####################################################################################
        foreach($bank in $script:Bullet_banks.getEnumerator() | sort Key)
        {
            $bank_file = "$dir\Resources\Bullet Banks\" + $bank.Key
            $bank_name = $bank.Key -replace ".txt$", ""


            $bank_name_label                          = New-Object system.Windows.Forms.Label
            $bank_name_label.text                     = "$bank_name";
            $bank_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $bank_name_label.Anchor                   = 'top,right'
            $bank_name_label.width                    = (($edit_bullets_form.width - 50) - $spacer)
            $bank_name_label.height                   = 30
            $bank_name_label.location                 = New-Object System.Drawing.Point(20,$y_pos)
            $bank_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $edit_bullets_form.controls.Add($bank_name_label);

            $y_pos = $y_pos + 30;
                    
            $edit_button           = New-Object System.Windows.Forms.Button
            $edit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $edit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $edit_button.Width     = 120
            $edit_button.height     = 25
            $edit_button.Location  = New-Object System.Drawing.Point(20,$y_pos);
            $edit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $edit_button.Text      = "Manual Edit"
            $edit_button.Name      = $bank_file 
            $edit_button.Add_Click({
                $message = "Making edits to a Bullet Banks:`n - Must be kept in .txt file format`n - Must remain one bullet per line`n"
                [System.Windows.MessageBox]::Show($message,"!!!WARNING!!!",'Ok')
                explorer.exe $this.name
            });
            $edit_bullets_form.controls.Add($edit_button) 

            $delete_button           = New-Object System.Windows.Forms.Button
            $delete_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $delete_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $delete_button.Width     = 90
            $delete_button.height     = 25
            $delete_button.Location  = New-Object System.Drawing.Point(($edit_button.Location.x + $edit_button.width + 5),$y_pos);
            $delete_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $delete_button.Text      ="Delete"
            $delete_button.Name      = $bank_file 
            $delete_button.Add_Click({
                $file = [System.IO.Path]::GetFileNameWithoutExtension($this.name)
                $message = "Are you sure you want to delete the `"$file`" bank? You cannot revert this action.`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    if(Test-path -LiteralPath $this.name)
                    {
                        Remove-Item -LiteralPath $this.name
                    }
                    $script:Bullet_banks.remove("$file.txt");
                    build_bullet_menu
                    $script:reload_function = "manage_bullets_dialog"
                    $edit_bullets_form.close();         
                }

            });
            $edit_bullets_form.controls.Add($delete_button)

            $rename_button           = New-Object System.Windows.Forms.Button
            $rename_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $rename_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $rename_button.Width     = 90
            $rename_button.height     = 25
            $rename_button.Location  = New-Object System.Drawing.Point(($delete_button.Location.x + $delete_button.width + 5),$y_pos);
            $rename_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $rename_button.Text      ="Rename"
            $rename_button.Name      = $bank_file 
            $rename_button.Add_Click({
                $old_name = $this.name
                $new_name = rename_dialog $old_name

                #write-host ON $old_name
                #write-host NN $new_name
                
                if(($new_name -cne $old_name) -and ($new_name -ne ""))
                {
                    $old_key = [System.IO.Path]::GetFileNameWithoutExtension($old_name)
                    $new_key = [System.IO.Path]::GetFileNameWithoutExtension($new_name)
                    $old_value = $script:Bullet_banks["$old_key.txt"]
                    #write-host OV $old_value
                    $script:Bullet_banks.remove("$old_key.txt");
                    $script:Bullet_banks.add("$new_key.txt",$old_value);
                    build_bullet_menu    
                    $script:reload_function = "manage_bullets_dialog"
                    $edit_bullets_form.close();
                }
            });
            $edit_bullets_form.controls.Add($rename_button)
            


            $enable_checkbox = new-object System.Windows.Forms.checkbox
            $enable_checkbox.Location = new-object System.Drawing.Size(($rename_button.Location.x + $rename_button.width + 5),$y_pos);
            $enable_checkbox.Size = new-object System.Drawing.Size(200,30)
            $enable_checkbox.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $enable_checkbox.name = $bank.key          
            $enable_checkbox.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            if($bank.value -eq "0")
            {
                $enable_checkbox.Checked = $false
                $enable_checkbox.text = "Disabled"
            }
            else
            {
                $enable_checkbox.Checked = $true
                $enable_checkbox.text = "Enabled"
            }
            $enable_checkbox.Add_CheckStateChanged({
                if($this.Checked -eq $true)
                {
                    $this.text = "Enabled"
                    $script:Bullet_banks[$this.name] = 1;
                    build_bullet_menu
                }
                else
                {
                    $this.text = "Disabled"
                    $script:Bullet_banks[$this.name] = 0;
                    build_bullet_menu
                }
            })
            $edit_bullets_form.controls.Add($enable_checkbox);


            #######################################################
            $line_count = 0
            $reader = New-Object IO.StreamReader $bank_file
            while($null -ne ($line = $reader.ReadLine()))
            {
                if(($line -match "^-") -and ($line.Length -gt 50))
                {
                    #write-host $line
                    $line_count++;
                }
            }
            $reader.Close() 
            $item_count_label                          = New-Object system.Windows.Forms.Label
            $item_count_label.text                     = "$line_count Bullets";
            $item_count_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $item_count_label.Anchor                   = 'top,right'
            $item_count_label.TextAlign = "MiddleRight"
            $item_count_label.width                    = 180
            $item_count_label.height                   = 30
            $item_count_label.location                 = New-Object System.Drawing.Point((($edit_bullets_form.width - 210) - $spacer),$y_pos);
            $item_count_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            $edit_bullets_form.controls.Add($item_count_label);

            $y_pos = $y_pos + 30
            $separator_bar                             = New-Object system.Windows.Forms.Label
            $separator_bar.text                        = ""
            $separator_bar.AutoSize                    = $false
            $separator_bar.BorderStyle                 = "fixed3d"
            #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
            $separator_bar.Anchor                      = 'top,left'
            $separator_bar.width                       = (($edit_bullets_form.width - 50) - $spacer)
            $separator_bar.height                      = 1
            $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
            $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $separator_bar.TextAlign                   = 'MiddleLeft'
            $edit_bullets_form.controls.Add($separator_bar);
            $y_pos = $y_pos + 5
        }
    
        $edit_bullets_form.ShowDialog()
    }
    else
    {
        $message = "You have no Bullet Banks to edit.`nYou must create or import a Bullet Bank first."
        #[System.Windows.MessageBox]::Show($message,"No bank",'Ok')

        $error_label                          = New-Object system.Windows.Forms.Label
        $error_label.text                     = "$message";
        $error_label.ForeColor                = "Red"
        $error_label.Anchor                   = 'top,right'
        $error_label.width                    = ($edit_bullets_form.width - 10)
        $error_label.height                   = 50
        $error_label.TextAlign = "MiddleCenter"
        $error_label.location                 = New-Object System.Drawing.Point(10,$y_pos)
        $error_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $edit_bullets_form.controls.Add($error_label);
        $edit_bullets_form.ShowDialog()
    }
    
}
################################################################################
######Import Bullet Processing##################################################
function import_bullet_processing($input_file,$output_file)
{
    $success = 0;

    $input_ext = [System.IO.Path]::GetExtension("$input_file")
    
    $processing_area = "$dir\Resources\Required\Processing\" + [io.path]::GetFileNameWithoutExtension($output_file) + [System.IO.Path]::GetExtension("$output_file")
    $processing_file = "";
    if($input_ext -match ".xls$|.xlsx$")
    {    
        [Array]$output = xls_to_csv $input_file $processing_area 3
        $processing_file = $output.Keys | % ToString #Convert Hash Return to String
    }
    elseif($input_ext -match ".csv$|.txt$")
    {
        Copy-Item -LiteralPath $input_file $processing_area
        $processing_file = $processing_area
    }
    elseif($input_ext -match ".doc$|.docx$")
    {
        $output = doc_to_txt $input_file $processing_area
        $processing_file = $processing_area
    }
    else
    {
        #write-host "Invalid File type"
        return 0;
    }

    #write-host Processing: $processing_file
    $reader = New-Object IO.StreamReader "$processing_file"
    $line_counter = 0;
    $bullets = new-object System.Collections.Hashtable
    $bullets.clear();
    $duplicates = 0;

    $space_hash2 = new-object System.Collections.Hashtable
    $space_hash2.add(" ",11.3886113886114)
    $space_hash2.add(" ",9.47735191637631)
    $space_hash2.add(" ",4.75524475524475)
    while($null -ne ($line = $reader.ReadLine()))
    {
        $line_counter++;
        $go = 1;
        $line = $line -replace "’|â€™","'"
        $line = $line -replace " | | ", " "
        $line = $line -replace "  "," "
        $line = $line -replace "\udfd7|\udbc2|\u00a0", " "
        $line = $line -replace "\u2014|\u2013", "-"
        $line = $line -replace "\u201c|\u201d", "`""
        #write-host $line
       
        #####################################
        ###Don't convert Txt to CSV
        [System.Collections.ArrayList]$line_split = @()
        if(!($input_ext -match ".txt$|.docx$|.doc"))
        {
            [Array]$line_split = csv_line_to_array $line
        }
        else
        {
            $line_split.Add($line);
        }
        ##############################@@@@@
        foreach($section in $line_split)
        {
            $section = $section.trim()
            if(($section -match "^-") -and ($section.length -gt 70))
            {
                

                #############################################################################
                ##Calculate Bullet Size
                [double]$size = 0;
                for ($i = 0; $i -lt $section.length; $i++)
                {
                    $character = $section[$i]
                    if($character_blocks.Contains("$character"))
                    {
                        $size = $size + ($character_blocks["$character"])
                    }
                    else
                    {
                        $utf = '{0:x4}' -f [int][char]$character + ""
                        #write-host "Missing Character Block For `"$character`" ($utf)"
                        #exit
                    }
                }
                ##############################################################################
                ##Compress/Expand Bullet
                $position = 0;
                $matches = [regex]::Matches("$section"," ")
                $stop = 0;
                if($matches.Success)
                {
                    foreach($space_type in $space_hash2.GetEnumerator() | sort value -Descending) #Loop 1
                    {
                        if($stop -eq 1)
                        {
                            break
                        }
                        foreach($match in $matches) #Loop2
                        {
                            if($match.index -ne 1)
                            {
                                if(($stop -ne 1) -and ($size -gt 2718))
                                {
                                    $current_space = $section.substring($match.index,1)
                                    $current_size = $character_blocks["$current_space"]
                                    $size = $size - $current_size
                                    $size = $size + $space_type.Value
                                    $section = $section.remove($match.index,1)
                                    $section = $section.insert($match.index,$space_type.key)
                                }
                                if($size -le 2718)
                                {
                                    $stop = 1;
                                    #$size = $size
                                    break
                                }
                            }
        
                        }
                    }
                }
               ##############################################################################
                $size = [int][math]::floor($size)
                if(($size -gt 2000) -and ($size -le 2720))
                {
                    
                    [String]$section = "$section"
                    if(!($bullets.Contains($section)))
                    {
                        $bullets.Add($section,$size);
                            #write-host "Added Bullet:         $line_counter = $section ($size)"
                    }
                    else
                    {
                        $duplicates++; 
                            write-host "Duplicate:            $line_counter = $section ($size)"
                    }
                }
                else
                {
                        if($size -gt 2720)
                        {
                            write-host "Too Long Eliminated:  $line_counter = $section ($size)"
                        }
                        else
                        {
                            write-host "Too Short Eliminated: $line_counter = $section ($size)"
                        }
                }

            }
            else
            {
                #write-host Removed: $line_counter = $section
            }

        }
    }
    $reader.Close();
    if(Test-Path -LiteralPath "$processing_file")
    {
        Remove-Item -LiteralPath "$processing_file"
    }


    #################Write the Bullets

    $writer = [System.IO.StreamWriter]::new($output_file)
    foreach($bullet in $bullets.GetEnumerator() | sort key)
    {
        $writer.WriteLine($bullet.key)
        #write-host A= $bullet.key = $bullet.value
    }
    $writer.close()
    #write-host
    #write-host "Duplicates Found: $duplicates"
    #write-host Bullets Found: $bullets.get_count();
    $bullet_count = $bullets.get_count();


    if(Test-Path -LiteralPath $output_file)
    {
        $file = [io.path]::GetFileNameWithoutExtension($output_file) + ".txt"
        $script:Bullet_banks.add($file,1);
        build_bullet_menu
        
        $Script:recent_editor_text = "Changed"
        $message = "$bullet_count Bullets Found`n$duplicates Duplicates Found"
        [System.Windows.MessageBox]::Show($message,"Bullets",'Ok')
        #Write-host Complete
        $success = 1;
    }
  
    return $success

























}
################################################################################
######Create Bullet Bank########################################################
function create_bullet_bank
{
    $create_bullet_bank_form = New-Object System.Windows.Forms.Form
    $create_bullet_bank_form.FormBorderStyle = 'Fixed3D'
    $create_bullet_bank_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $create_bullet_bank_form.Location = new-object System.Drawing.Point(0, 0)
    $create_bullet_bank_form.Size = new-object System.Drawing.Size(305, 120)
    $create_bullet_bank_form.MaximizeBox = $false
    $create_bullet_bank_form.SizeGripStyle = "Hide"
    $create_bullet_bank_form.Text = "Create Bullet Bank"
    #$create_bullet_bank_form.TopMost = $True
    $create_bullet_bank_form.TabIndex = 0
    $create_bullet_bank_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $bullet_bank_name_label                          = New-Object system.Windows.Forms.Label
    $bullet_bank_name_label.text                     = "Bank Name:";
    #$bullet_bank_name_label.AutoSize                 = $true
    #$bullet_bank_name_label.BackColor                = 'Green'
    $bullet_bank_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $bullet_bank_name_label.Anchor                   = 'top,right'
    $bullet_bank_name_label.width                    = 125
    $bullet_bank_name_label.height                   = 30
    $bullet_bank_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $bullet_bank_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $create_bullet_bank_form.controls.Add($bullet_bank_name_label);

    $bullet_bank_name_input                         = New-Object system.Windows.Forms.TextBox                       
    $bullet_bank_name_input.AutoSize                 = $true
    $bullet_bank_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $bullet_bank_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $bullet_bank_name_input.Anchor                   = 'top,left'
    $bullet_bank_name_input.width                    = 150
    $bullet_bank_name_input.height                   = 30
    $bullet_bank_name_input.location                 = New-Object System.Drawing.Point(140,12)
    $bullet_bank_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $bullet_bank_name_input.text                     = ""
    $create_bullet_bank_form.controls.Add($bullet_bank_name_input);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($create_bullet_bank_form.width / 2) - ($submit_button.width)),45);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Submit"
    $submit_button.Add_Click({ 
        [array]$errors = "";
        $file = $bullet_bank_name_input.text + ".txt"
        if($bullet_bank_name_input.text -eq "")
        {
            $errors += "You must provide a name."
        }
        #write-host "$dir\Resources\Bullet Banks\$file"
        if(Test-path "$dir\Resources\Bullet Banks\$file")
        {
            $errors += "Bullet Bank already exists."
        }
        if($errors.count -eq 1)
        {
            $message = "Are you sure you want to save changes?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Create?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                Add-Content -literalpath "$dir\Resources\Bullet Banks\$file" ""
                $script:Bullet_banks.add($file,1);
                build_bullet_menu
                $create_bullet_bank_form.close();                      
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }


    });
    $create_bullet_bank_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($create_bullet_bank_form.width / 2)),45);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $create_bullet_bank_form.close();
    });
    $create_bullet_bank_form.controls.Add($cancel_button) 


    $create_bullet_bank_form.ShowDialog()

}
################################################################################
######Update Feeder#############################################################
function update_feeder
{
    if($script:feeder_job -eq "")
    {
        ############################################
        #Start Job
        $script:feeder_job = Start-Job -ScriptBlock {
            $current_bullet = $using:current_bullet
            $script:Bullet_bank = $using:Bullet_bank


            $rack_and_stack = @{};
            $current_words = $current_bullet.ToLower() -replace "[^a-z0-9]| | ",' '
            

            $current_bullet_modified = $current_bullet.ToLower() -replace "[^a-z0-9]| | ",' '
            $current_bullet_modified_wordsplit = $current_bullet_modified -split ' ';
            $current_bullet_modified_wordsplit = ($current_bullet_modified_wordsplit | sort length -desc | select -first 6)

        
            foreach($bullet in $script:Bullet_bank.getEnumerator())
            {
                $bullet = $bullet.key
                $match_score = 0;
                $match_bullet = $bullet.ToLower() -replace "[^a-z0-9]| | ",' '
                foreach($word in $current_bullet_modified_wordsplit)
                {
                    if($match_bullet -match "$word")
                    {
                        $match_score = $match_score + $word.length;
                    }
                }
                if($match_score -ne 0)
                {
                    if(($rack_and_stack.Get_Count()) -le 14)
                    {
                        if(!($rack_and_stack.Contains($bullet)))
                        {
                            
                            $rack_and_stack.Add("$bullet",$match_score);
                        } 
                    }
                    else
                    { 
                        if(!($rack_and_stack.Contains("$bullet")))
                        {
                            foreach($scored_bullet in $rack_and_stack.getEnumerator() | Sort Value -Descending) 
                            {
                                $scored_bullet1 = $scored_bullet.key
                                if($match_score -gt $scored_bullet.value)
                                {
                                    $rack_and_stack.Remove("$scored_bullet1")
                                    $rack_and_stack.Add("$bullet",$match_score);
                                    break;
                                }
                            }

                        }
                    }
                }
            }
            $text = "";
            foreach($scored_bullet in $rack_and_stack.getEnumerator() | Sort Value -Descending) 
            {
                $text = $text + $scored_bullet.key + "`n";
            }
            return "$text"
        }
    }
    else
    {   
        if($script:feeder_job.state -eq "Completed")
        {   
            $text = Receive-Job -Job $script:feeder_job
            if($text -ne $feeder_box.text)
            {
                $feeder_box.text = $text
                #$feeder_box.Font = [Drawing.Font]::New('Times New Roman', 14)
                $feeder_box.Font = [Drawing.Font]::New($script:theme_settings['FEEDER_FONT'], [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'])



                ##############################################
                ############Fix Missing Half Spaces
                $pattern = " | | "
                $matches = [regex]::Matches($feeder_box.text, $pattern)
                if($matches)
                {  
                    foreach($match in $matches)
                    {
                        #write-host ---------------------------------
                        #write-host MV $match.value.length
                        #write-host MI $match.index
                        $feeder_box.SelectionStart = $match.index
                        $feeder_box.SelectionLength = $match.value.length
                        $feeder_box.SelectionFont = [Drawing.Font]::New('Times New Roman', [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
                        $feeder_box.DeselectAll()

                    }
                }
                ##############################################

                while($feeder_box.ZoomFactor -ne $script:zoom) 
                {
                    #Zoom Changes during RTF replace, but won't change in time... this is a work around.
                    $feeder_box.ZoomFactor = $script:zoom
                }  
            }
            $script:feeder_job = "";
        }
    }
}
################################################################################
######Load Bullets##############################################################
function load_bullets
{
    $Script:bullet_bank = @{};
    #############Load Package Bullets
    if($script:settings['LOAD_PACKAGES_AS_BULLETS'] -eq 1)
    {
        foreach($package in $script:package_list.getEnumerator())
        {
            $file = $package.key
            $status = $package.value
            if($file -ne $script:settings['PACKAGE'])
            {
                if((Test-Path -LiteralPath "$dir\Resources\Packages\$file\Snapshot.txt") -and ($status -eq 1))
                {    
                    $line_number = 0;
                    $reader = [System.IO.File]::OpenText("$dir\Resources\Packages\$file\Snapshot.txt")
                    while($null -ne ($line = $reader.ReadLine()))
                    {
                        $line_number++;
                        if($line -match "^- |^- |^- |^- " -and $line.length -ge 85)
                        {
                            if(!($Script:bullet_bank.contains("$line")))
                            {
                                $data = $file + "::" + $line_number
                                $Script:bullet_bank.add("$line",$data)
                            }
                        }
                    }
                    $reader.Close();
                }
            }
        }
    }
    ##############Load Bullet Banks
    foreach($bank in $script:Bullet_banks.getEnumerator())
    {
        $file = $bank.key
        $status = $bank.value

        if((Test-Path -LiteralPath "$dir\Resources\Bullet Banks\$file") -and ($status -eq 1))
        {    
            $line_number = 0;
            $reader = [System.IO.File]::OpenText("$dir\Resources\Bullet Banks\$file")
            while($null -ne ($line = $reader.ReadLine()))
            {
                $line_number++;
                if($line -match "^- |^- |^- |^- " -and $line.length -ge 85)
                {
                    if(!($Script:bullet_bank.contains("$line")))
                    {
                        $data = $file + "::" + $line_number
                        $Script:bullet_bank.add("$line",$data)
                    }
                }
            }
            $reader.Close();
        }
    }
    #write-host $Script:bullet_bank.Get_Count() Bullets Loaded
    sidekick_display
}
################################################################################
######Load Acronyms#############################################################
function load_acronyms
{
    
    $script:acronym_list = New-Object system.collections.hashtable
    foreach($list in $script:acronym_lists.getEnumerator())
    {
        $file = $list.key
        $status = $list.value

        if((Test-Path -LiteralPath "$dir\Resources\Acronym Lists\$file") -and ($status -eq 1))
        {    
            $reader = [System.IO.File]::OpenText("$dir\Resources\Acronym Lists\$file")
            while($null -ne ($line = $reader.ReadLine()))
            {
                [Array]$line_split = csv_line_to_array $line
                if($line_split[0] -and $line_split[1])
                {
                    [string]$key = $line_split[0] + "::" + $line_split[1]
                    if($line_split[0] -and $line_split[1])
                    {
                        if(!($script:acronym_list.ContainsKey("$key")))
                        {
                                    $script:acronym_list.Add("$key","");
                        }
                        else
                        {
                            #write-host "Duplicate: $line"
                        }
                    }
                }
            }
            $reader.Close();
        }
    }
    #write-host $script:acronym_list.Get_Count() Acronyms Loaded
}
################################################################################
######Create Acronym List#######################################################
function create_acronym_list
{
    $create_acronym_list_form = New-Object System.Windows.Forms.Form
    $create_acronym_list_form.FormBorderStyle = 'Fixed3D'
    $create_acronym_list_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $create_acronym_list_form.Location = new-object System.Drawing.Point(0, 0)
    $create_acronym_list_form.Size = new-object System.Drawing.Size(290, 120)
    $create_acronym_list_form.MaximizeBox = $false
    $create_acronym_list_form.SizeGripStyle = "Hide"
    $create_acronym_list_form.Text = "Create New Acronym or Abbreviation List"
    #$create_acronym_list_form.TopMost = $True
    $create_acronym_list_form.TabIndex = 0
    $create_acronym_list_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $acronym_list_name_label                          = New-Object system.Windows.Forms.Label
    $acronym_list_name_label.text                     = "List Name:";
    #$acronym_list_name_label.AutoSize                 = $true
    #$acronym_list_name_label.BackColor                = 'Green'
    $acronym_list_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $acronym_list_name_label.Anchor                   = 'top,right'
    $acronym_list_name_label.width                    = 110
    $acronym_list_name_label.height                   = 30
    $acronym_list_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $acronym_list_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $create_acronym_list_form.controls.Add($acronym_list_name_label);

    $acronym_list_name_input                         = New-Object system.Windows.Forms.TextBox                       
    $acronym_list_name_input.AutoSize                 = $true
    $acronym_list_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $acronym_list_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $acronym_list_name_input.Anchor                   = 'top,left'
    $acronym_list_name_input.width                    = 150
    $acronym_list_name_input.height                   = 30
    $acronym_list_name_input.location                 = New-Object System.Drawing.Point(125,12)
    $acronym_list_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $create_acronym_list_form.controls.Add($acronym_list_name_input);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($create_acronym_list_form.width / 2) - ($submit_button.width)),45);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Submit"
    $submit_button.Add_Click({ 
        [array]$errors = "";
        $file = $acronym_list_name_input.text + ".csv"
        if($acronym_list_name_input.text -eq "")
        {
            $errors += "You must provide a name."
        }
        #write-host "$dir\Resources\Acronym Lists\$file"
        if(Test-path "$dir\Resources\Acronym Lists\$file")
        {
            $errors += "Acronym list already exists."
        }
        if($errors.count -eq 1)
        {
            $message = "Are you sure you want to save changes?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Create?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                Add-Content -literalpath "$dir\Resources\Acronym Lists\$file" ""
                $Script:acronym_lists.add($file,1);
                build_acronym_menu
                $create_acronym_list_form.close();                      
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }


    });
    $create_acronym_list_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($create_acronym_list_form.width / 2)),45);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $create_acronym_list_form.close();
    });
    $create_acronym_list_form.controls.Add($cancel_button) 


    $create_acronym_list_form.ShowDialog()

}
################################################################################
######Import Bullet Bank Form###################################################
function import_bullet_form
{
    $import_bullet_form = New-Object System.Windows.Forms.Form
    $import_bullet_form.FormBorderStyle = 'Fixed3D'
    $import_bullet_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $import_bullet_form.Location = new-object System.Drawing.Point(0, 0)
    $import_bullet_form.MaximizeBox = $false
    $import_bullet_form.SizeGripStyle = "Hide"
    $import_bullet_form.Size='500,170'
    $import_bullet_form.Text = "Import Bullet Bank"
    #$import_bullet_form.TopMost = $True
    $import_bullet_form.TabIndex = 0
    $import_bullet_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $bank_name_label                          = New-Object system.Windows.Forms.Label
    $bank_name_label.text                     = "Bank Name:";
    $bank_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $bank_name_label.Anchor                   = 'top,right'
    $bank_name_label.width                    = 125
    $bank_name_label.height                   = 30
    $bank_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $bank_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $import_bullet_form.controls.Add($bank_name_label);

    $bank_file_label                          = New-Object system.Windows.Forms.Label
    $bank_file_label.text                     = "File Location:";
    $bank_file_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $bank_file_label.Anchor                   = 'top,right'
    $bank_file_label.width                    = 140
    $bank_file_label.height                   = 30
    $bank_file_label.location                 = New-Object System.Drawing.Point(10,45)
    $bank_file_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $import_bullet_form.controls.Add($bank_file_label);

    $bank_name_input                          = New-Object system.Windows.Forms.TextBox                       
    $bank_name_input.AutoSize                 = $true
    $bank_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $bank_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $bank_name_input.Anchor                   = 'top,left'
    $bank_name_input.width                    = 150
    $bank_name_input.height                   = 30
    $bank_name_input.location                 = New-Object System.Drawing.Point(140,12)
    $bank_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $bank_name_input.text                     = ""
    $bank_name_input.Add_TextChanged({
        $caret = $bank_name_input.SelectionStart;
        $bank_name_input.text = $bank_name_input.text -replace '[^0-9A-Za-z ,]', ''
        $bank_name_input.text = (Get-Culture).TextInfo.ToTitleCase($bank_name_input.text)
        $bank_name_input.SelectionStart = $caret

    });
    $import_bullet_form.controls.Add($bank_name_input);

    $bank_file_input                          = New-Object system.Windows.Forms.TextBox                       
    $bank_file_input.AutoSize                 = $true
    $bank_file_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $bank_file_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $bank_file_input.Anchor                   = 'top,left'
    $bank_file_input.width                    = 220
    $bank_file_input.height                   = 30
    $bank_file_input.location                 = New-Object System.Drawing.Point(150,47)
    $bank_file_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $bank_file_input.text                     = ""
    $bank_file_input.Add_TextChanged({
        $prompt_return = $this.text
        $prompt_return = $prompt_return -replace "^`"|`"$" 
        if($prompt_return -ne $Null -and $prompt_return -like "*.xls*" -or $prompt_return -like "*.csv*" -or $prompt_return -like "*.txt*" -and (Test-Path $prompt_return) -eq $True)
        {
            $bank_file_input.Text="$prompt_return"
            $submit_button.Enabled = $true
        }
        else
        {
            $submit_button.Enabled = $false
        }

    })
    $import_bullet_form.controls.Add($bank_file_input);

    $browse_button          = New-Object System.Windows.Forms.Button
    $browse_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $browse_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $browse_button.Width     = 110
    $browse_button.height     = 25
    $browse_button.Location  = New-Object System.Drawing.Point(($import_bullet_form.width - $browse_button.width - 20),46);
    $browse_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $browse_button.Text      ="Browse"
    $browse_button.Add_Click({

        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
        #$OpenFileDialog.filter = "All files (*.*)| *.*"
        $OpenFileDialog.filter = "Files (*.xls, *.xlsx, *.csv, *.txt, *.doc, *.docx)|*.xls;*.xlsx;*.csv;*.txt;*.doc;*.docx"
        $OpenFileDialog.ShowDialog() | Out-Null
        $prompt_return = $OpenFileDialog.filename

        if($prompt_return -ne $Null -and $prompt_return -like "*.xls*" -or $prompt_return -like "*.csv*" -or $prompt_return -like "*.txt*" -or $prompt_return -like "*.doc*" -and (Test-Path $prompt_return) -eq $True)
        {
            $bank_file_input.Text="$prompt_return"
            $submit_button.Enabled = $true
        }
        else
        {
            $submit_button.Enabled = $false
        }       
    });
    $import_bullet_form.controls.Add($browse_button) 

    $note_lablel1                          = New-Object system.Windows.Forms.Label
    $note_lablel1.text                     = "*Note:  Import will scan all Worksheets in a Workbook*";
    #$note_lablel1.AutoSize                = $true
    #$note_lablel1.BackColor               = 'Green'
    $note_lablel1.ForeColor                = "Yellow"
    $note_lablel1.Anchor                   = 'top,right'
    $note_lablel1.TextAlign                = "MiddleCenter"
    $note_lablel1.width                    = ($import_bullet_form.width - 30)
    $note_lablel1.height                   = 20
    $note_lablel1.location                 = New-Object System.Drawing.Point(15,75)
    $note_lablel1.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $import_bullet_form.controls.Add($note_lablel1);

  

    $submit_button                         = New-Object System.Windows.Forms.Button
    $submit_button.BackColor               = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor               = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width                   = 110
    $submit_button.height                  = 25
    $submit_button.Location                = New-Object System.Drawing.Point((($import_bullet_form.width / 2) - ($submit_button.width)),100);
    $submit_button.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text                    = "Submit"
    if($bank_name_input.text -eq "")
    {
        $submit_button.Enabled = $false
    }
    $submit_button.Add_Click({ 
        [array]$errors = "";

        $bank_name = $bank_name_input.text 
        $output_file = "$dir\Resources\Bullet Banks\" + $bank_name_input.text + ".txt"
        $input_file = $bank_file_input.text

        if($bank_name.Length -le 2)
        {
            $errors += "Bullet Bank name too short"
        }
        if($bank_name.Length -ge 40)
        {
            $errors += "Bullet Bank name too long"
        }
        if(Test-path -literalpath "$output_file")
        {
            $errors += "bank name already exists"
        }
        if($errors.count -eq 1)
        {
            $message = "Are you sure you want to save changes?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Import?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                $success = import_bullet_processing $input_file $output_file
                if($success -eq 1)
                {
                    $import_bullet_form.close();
                }
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }
    });
    $import_bullet_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($import_bullet_form.width / 2)),100);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $import_bullet_form.close();
    });
    $import_bullet_form.controls.Add($cancel_button) 

    $import_bullet_form.ShowDialog()
}
################################################################################
######Import Acronym Form#######################################################
function import_acronym_form
{
    $import_acronym_form = New-Object System.Windows.Forms.Form
    $import_acronym_form.FormBorderStyle = 'Fixed3D'
    $import_acronym_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $import_acronym_form.Location = new-object System.Drawing.Point(0, 0)
    $import_acronym_form.MaximizeBox = $false
    $import_acronym_form.SizeGripStyle = "Hide"
    $import_acronym_form.Size='500,170'
    $import_acronym_form.Text = "Import Acronyms and/or Abbreviations"
    #$import_acronym_form.TopMost = $True
    $import_acronym_form.TabIndex = 0
    $import_acronym_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $list_name_label                          = New-Object system.Windows.Forms.Label
    $list_name_label.text                     = "List Name:";
    $list_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $list_name_label.Anchor                   = 'top,right'
    $list_name_label.width                    = 110
    $list_name_label.height                   = 30
    $list_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $list_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $import_acronym_form.controls.Add($list_name_label);

    $list_file_label                          = New-Object system.Windows.Forms.Label
    $list_file_label.text                     = "File Location:";
    $list_file_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $list_file_label.Anchor                   = 'top,right'
    $list_file_label.width                    = 140
    $list_file_label.height                   = 30
    $list_file_label.location                 = New-Object System.Drawing.Point(10,45)
    $list_file_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $import_acronym_form.controls.Add($list_file_label);

    $list_name_input                          = New-Object system.Windows.Forms.TextBox                       
    $list_name_input.AutoSize                 = $true
    $list_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $list_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $list_name_input.Anchor                   = 'top,left'
    $list_name_input.width                    = 150
    $list_name_input.height                   = 30
    $list_name_input.location                 = New-Object System.Drawing.Point(125,12)
    $list_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $list_name_input.text                     = ""
    $list_name_input.Add_TextChanged({
        $caret = $list_name_input.SelectionStart;
        $list_name_input.text = $list_name_input.text -replace '[^0-9A-Za-z ,]', ''
        $list_name_input.text = (Get-Culture).TextInfo.ToTitleCase($list_name_input.text)
        $list_name_input.SelectionStart = $caret

    });
    $import_acronym_form.controls.Add($list_name_input);

    $list_file_input                          = New-Object system.Windows.Forms.TextBox                       
    $list_file_input.AutoSize                 = $true
    $list_file_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $list_file_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $list_file_input.Anchor                   = 'top,left'
    $list_file_input.width                    = 220
    $list_file_input.height                   = 30
    $list_file_input.location                 = New-Object System.Drawing.Point(150,47)
    $list_file_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $list_file_input.text                     = ""
    $list_file_input.Add_TextChanged({
        $prompt_return = $this.text
        $prompt_return = $prompt_return -replace "^`"|`"$" 
        if($prompt_return -ne $Null -and $prompt_return -like "*.xls*" -or $prompt_return -like "*.csv*" -and (Test-Path $prompt_return) -eq $True)
        {
            $list_file_input.Text="$prompt_return"
            $submit_button.Enabled = $true
        }
        else
        {
            $submit_button.Enabled = $false
        }

    })
    $import_acronym_form.controls.Add($list_file_input);

    $browse_button          = New-Object System.Windows.Forms.Button
    $browse_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $browse_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $browse_button.Width     = 110
    $browse_button.height     = 25
    $browse_button.Location  = New-Object System.Drawing.Point(($import_acronym_form.width - $browse_button.width - 20),46);
    $browse_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $browse_button.Text      ="Browse"
    $browse_button.Add_Click({

        $prompt_return = prompt_for_file

        if($prompt_return -ne $Null -and $prompt_return -like "*.xls*" -or $prompt_return -like "*.csv*" -and (Test-Path $prompt_return) -eq $True)
        {
            $list_file_input.Text="$prompt_return"
            $submit_button.Enabled = $true
        }
        else
        {
            $submit_button.Enabled = $false
        }       
    });
    $import_acronym_form.controls.Add($browse_button) 

    $note_lablel1                          = New-Object system.Windows.Forms.Label
    $note_lablel1.text                     = "*Note:  File must contain only two columns.*";
    #$note_lablel1.AutoSize                 = $true
    #$note_lablel1.BackColor                = 'Green'
    $note_lablel1.ForeColor                = "yellow"
    $note_lablel1.Anchor                   = 'top,right'
    $note_lablel1.TextAlign = "MiddleCenter"
    $note_lablel1.width                    = ($import_acronym_form.width - 30)
    $note_lablel1.height                   = 20
    $note_lablel1.location                 = New-Object System.Drawing.Point(15,75)
    $note_lablel1.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $import_acronym_form.controls.Add($note_lablel1);

  

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($import_acronym_form.width / 2) - ($submit_button.width)),100);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Submit"
    if($list_name_input.text -eq "")
    {
        $submit_button.Enabled = $false
    }
    $submit_button.Add_Click({ 
        [array]$errors = "";

        $list_name = $list_name_input.text 
        $output_file = "$dir\Resources\Acronym Lists\" + $list_name_input.text + ".csv"
        $input_file = $list_file_input.text

        if($list_name.Length -le 2)
        {
            $errors += "Acronym and/or Abbreviation list name too short"
        }
        if($list_name.Length -ge 40)
        {
            $errors += "Acronym and/or Abbreviation list name too long"
        }
        if(Test-path -literalpath "$output_file")
        {
            $errors += "List name already exists"
        }
        if($errors.count -eq 1)
        {
            $message = "Are you sure you want to save changes?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Overwrite?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                $success = import_acronym_processing $input_file $output_file
                if($success -eq 1)
                {
                    $import_acronym_form.close();
                }
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }
    });
    $import_acronym_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($import_acronym_form.width / 2)),100);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $import_acronym_form.close();
    });
    $import_acronym_form.controls.Add($cancel_button) 

    $import_acronym_form.ShowDialog()
}
################################################################################
######Import Acronym Processing#################################################
function import_acronym_processing($input_file,$output_file)
{
    $success = 0;
    #write-host Input: $input_file
    #write-host Output: $output_file
    #write-host Success: $success
    $input_ext = [System.IO.Path]::GetExtension("$input_file")
    
    $processing_area = "$dir\Resources\Required\Processing\" + [io.path]::GetFileNameWithoutExtension($output_file) + [System.IO.Path]::GetExtension("$output_file")
    $processing_file = "";
    if($input_ext -match ".xls|.xlsx")
    {    
        [Array]$output = xls_to_csv $input_file $processing_area 1
        $processing_file = $output.Keys | % ToString #Convert Hash Return to String
    }
    elseif($input_ext -match ".csv")
    {
        Copy-Item -LiteralPath $input_file $processing_area
        $processing_file = $processing_area
    }
    else
    {
        #write-host "Invalid File type"
        return 0;
    }

    #write-host Processing: $processing_file
    $reader = [System.IO.File]::OpenText("$processing_file")
    $line_counter = 0;
    $acro_count = 0;
    $tracker = new-object System.Collections.Hashtable
    while($null -ne ($line = $reader.ReadLine()))
    {
        $line_counter++;
        $go = 1;
        $multi = @{};
        $line = $line -replace "’|â€™","'"
        $line = $line -replace " | | ", " "
        if($line_counter -eq 1)
        {
            #Remove Header
            if($line -match "Acronym|Abbreviation|Meaning")
            {
                $go = 0;
                #write-host "REMOVED Line $line_counter = $line"
            }
        }
        [Array]$line_split = csv_line_to_array $line

        ################Basic Elimination
        $line_split[0] = $line_split[0].trim()
        if(($line_split[0] -eq "") -or ($line_split[0] -match "^- "))
        {
            $go = 0;
        }
        if($line_split.count -ge 2)
        {
            $line_split[1] = $line_split[1].trim();
            if(($line_split[1] -eq "") -or ($line_split[1] -match "^- "))
            {
                $go = 0;
            }
        }
        else
        {
            $go = 0;
        }
        if($line_split.count -ge 3)
        {
            $go = 0; #More than two columns
            #write-host "REMOVED $line_counter = $line"
        }
        #######################################
        if(($line_split[0]) -and ($line_split[1]) -and ($go -eq 1))
        {   
            $par_split1 = $line_split[0] -replace "\)",''
            $par_split1 = $par_split1 -split "\(".trim();
            $par_split2 = $line_split[1] -replace "\)",''
            $par_split2 = $par_split2 -split "\(".trim();
            $core1, $extras1 = $par_split1
            $core2, $extras2 = $par_split2
            if($core1.length -gt $core2.length)
            {
                $buffer = $line_split[0]
                $line_split[0] = $line_split[1]
                $line_split[1] = $buffer
                $line = $line_split[0] +"," + $line_split[1]
                $buffer = $core1
                $core1 = $core2
                $core2 = $buffer

                #write-host "SWAPPED Line $line_counter = $line"
            }
            else
            {
                #write-host "NO SWAP $line_counter = $line"
            }
            if(($core1.length -lt 1) -or ($core2.length -lt 3))
            {
                #write-host "INVALID Line $line_counter = $line"
                $go = 0;
            }
        }

        #############Parenthesis Matching
        if(($go -eq 1) -and (($line_split[0] -match "\(") -and ($line_split[0] -match "\)")) -or (($go -eq 1) -and ($line_split[1] -match "\(") -and ($line_split[1] -match "\)")))
        {
            $par_split1 = $line_split[0] -replace "\)",''
            $par_split1 = $par_split1 -split "\(".trim();
            $par_split2 = $line_split[1] -replace "\)",''
            $par_split2 = $par_split2 -split "\(".trim();
            $core1, $extras1 = $par_split1
            $core2, $extras2 = $par_split2
            #write-host ---------------------------------------------------------
            #write-host C1 $core1
            #write-host C2 $core2
            if(!($tracker.contains($core2)))
            {
                $tracker.Add($core2,$core1);
                #write-host V0 = $core2, $core1
            }
            foreach($split1 in $extras1)
            {
                $found = 0;
                foreach($split2 in $extras2)
                {
                    if($split1 -eq $split2)
                    {
                        $word1 = "$core1$split1"
                        $word2 = "$core2$split2"                      
                        $found = 1;
                        if(!($tracker.contains($word2)))
                        {
                            $tracker.Add($word2,$word1);
                            #write-host V1 = $word2,$word1
                        }
                        break;
                    }
                    elseif(($split1.substring(($split1.length -1),1)) -match ($split2.substring(($split2.length -1),1)))
                    {
                        $word1 = "$core1$split1"
                        $word2 = "$core2$split2"
                        $word2 = $word2 -replace "yied$","ied"
                        $found = 1;
                        if(!($tracker.contains($word2)))
                        {
                            $tracker.Add($word2,$word1);
                            #write-host V2 = $word2,$word1
                        }
                        break;
                    }
                    else
                    {
                        $word1 = "$core1$split1"
                        $word2 = "$core2$split2"
                        if(!($tracker.contains($word2)))
                        {
                            $tracker.Add($word2,$word1);
                            #write-host V3 = $word2,$word1
                        }
                        $found = 1;
                    }

                }
                if($found -eq 0)
                {
                    $word1 = "$core1$split2"
                    $word2 = "$core2$split2"
                    #write-host P $word1 $split2
                    #write-host P $word2 $split2

                    $word1 = $word1 -replace "ss$","s"
                    $word1 = $word1 -replace "eded$","ed"
                    $word2 = $word2 -replace "ss$","s"
                    $word2 = $word2 -replace "eded$","ed"  
                    if(!($tracker.contains($word2)))
                    {
                        $tracker.Add($word2,$word1);
                        #write-host NF1 = $word1 $word2
                    }
                }
            }
            #############################################################################
            foreach($split2 in $extras2)
            {
                $found = 0;
                foreach($split1 in $extras1)
                {
                    if($split2 -eq $split1)
                    {
                        $word2 = "$core2$split2"
                        $word1 = "$core1$split1"
                        $found = 1; 
                        if(!($tracker.contains($word2)))
                        {
                            $tracker.Add($word2,$word1);
                            #write-host V4 = $word2,$word1
                        }
                        break;
                    }
                    elseif(($split2.substring(($split2.length -1),1)) -match ($split1.substring(($split1.length -1),1)))
                    {
                        $word2 = "$core2$split2"
                        $word1 = "$core1$split1"
                        $word2 = $word2 -replace "yied$","ied"
                        $found = 1;
                        if(!($tracker.contains($word2)))
                        {
                            $tracker.Add($word2,$word1);
                            #write-host V5 = $word2,$word1
                        }
                        break;
                    }
                    else
                    {
                        $word2 = "$core2$split2"
                        $word1 = "$core1$split1"
                        #write-host V6 = $word2,$word1 #Don't use almost always wrong
                        $found = 1;
                    }

                }
                if($found -eq 0)
                {
                    #write-host NOT FOUND2 $split2 #Don't use almost always wrong

                }
            }
        }
        else
        {
            if($go -eq 1)
            {
                if(!($tracker.contains($line_split[1])))
                {
                    $tracker.Add($line_split[1],$line_split[0]);
                    $acro_count++;
                    #write-host "Added $line_counter = $line"
                }
                else
                {
                    #write-host "DUPLICATE $line_counter = $line"
                }
            }
            else
            {
                #write-host "FAILED $line_counter = $line"
            }
        }      
    }
    $reader.Close();
    if(Test-Path -LiteralPath "$processing_file")
    {
        Remove-Item -LiteralPath "$processing_file"
    }


    #################Write the Acros

    $writer = [System.IO.StreamWriter]::new($output_file)
    foreach($acro in $tracker.GetEnumerator() | sort value)
    {
        $line = $acro.value
        $line = csv_write_line $line $acro.key.ToLower()
        #write-host A = $line
        $writer.WriteLine($line)
    }
    $writer.close()



    if(Test-Path -LiteralPath $output_file)
    {
        $file = [io.path]::GetFileNameWithoutExtension($output_file) + ".csv"
        $Script:acronym_lists.add($file,1);
        build_acronym_menu
        
        $Script:recent_editor_text = "Changed"
        $message = "$acro_count Acronyms"
        [System.Windows.MessageBox]::Show($message,"Acronyms Added",'Ok')
        $success = 1;
        
    }
  
    return $success
}
################################################################################
######Edit Acronyms Dialog######################################################
function manage_acronyms_dialog
{
    
    $item_number = $Script:acronym_lists.get_count()
    #$item_number = 20
    
    $spacer = 0;
    $edit_acronym_form = New-Object System.Windows.Forms.Form
    $edit_acronym_form.FormBorderStyle = 'Fixed3D'
    $edit_acronym_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $edit_acronym_form.Location = new-object System.Drawing.Point(0, 0)
    $edit_acronym_form.MaximizeBox = $false
    $edit_acronym_form.SizeGripStyle = "Hide"
    $edit_acronym_form.Width = 800
    if($item_number -eq 0)
    {
        $edit_acronym_form.Height = 200;
    }
    elseif((($item_number * 65) + 140) -ge 600)
    {
        $edit_acronym_form.Height = 600;
        $edit_acronym_form.Autoscroll = $true
        $spacer = 20
    }
    else
    {
        $edit_acronym_form.Height = (($item_number * 65) + 140)
    }
    $edit_acronym_form.Text = "Manage Acronyms & Abbreviations"
    #$edit_acronym_form.TopMost = $True
    $edit_acronym_form.TabIndex = 0
    $edit_acronym_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    ################################################################################################
    $y_pos = 10;


    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "Manage Acronyms && Abbreviations";
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.BackColor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = $edit_acronym_form.Width
    $title_label.height                   = 30
    $title_label.TextAlign                = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point((($edit_acronym_form.Width / 2) - ($title_label.Width / 2) + 35),$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $edit_acronym_form.controls.Add($title_label);

    $y_pos = $y_pos + 40;
    $create_list_button           = New-Object System.Windows.Forms.Button
    $create_list_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $create_list_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $create_list_button.Width     = 150
    $create_list_button.height     = 25
    $create_list_button.Location  = New-Object System.Drawing.Point((($edit_acronym_form.width / 3) -80),$y_pos);
    $create_list_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $create_list_button.Text      ="Create List"
    $create_list_button.Name = ""
    $create_list_button.Add_Click({ 
        create_acronym_list
        $script:reload_function = "manage_acronyms_dialog" 
        $edit_acronym_form.close();
    })
    $edit_acronym_form.controls.Add($create_list_button)


    
    $import_list_button           = New-Object System.Windows.Forms.Button
    $import_list_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $import_list_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $import_list_button.Width     = 150
    $import_list_button.height     = 25
    $import_list_button.Location  = New-Object System.Drawing.Point((($edit_acronym_form.width / 3) + 75),$y_pos);
    $import_list_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $import_list_button.Text      ="Import List"
    $import_list_button.Name = ""
    $import_list_button.Add_Click({ 
        import_acronym_form
        $script:reload_function = "manage_acronyms_dialog" 
        $edit_acronym_form.close();
    })
    $edit_acronym_form.controls.Add($import_list_button)
    

    $add_acronym_button           = New-Object System.Windows.Forms.Button
    $add_acronym_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $add_acronym_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $add_acronym_button.Width     = 150
    $add_acronym_button.height     = 25
    $add_acronym_button.Location  = New-Object System.Drawing.Point((($edit_acronym_form.width / 3) + $add_acronym_button.Width + 80),$y_pos);
    $add_acronym_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $add_acronym_button.Text      ="Add Acronym"
    $add_acronym_button.Name = ""
    $add_acronym_button.Add_Click({add_to_acronyms})
    $edit_acronym_form.controls.Add($add_acronym_button)


    $y_pos = $y_pos + 35;
    $separator_bar                             = New-Object system.Windows.Forms.Label
    $separator_bar.text                        = ""
    $separator_bar.AutoSize                    = $false
    $separator_bar.BorderStyle                 = "fixed3d"
    #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
    $separator_bar.Anchor                      = 'top,left'
    $separator_bar.width                       = (($edit_acronym_form.width - 50) - $spacer)
    $separator_bar.height                      = 1
    $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
    $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $separator_bar.TextAlign                   = 'MiddleLeft'
    $edit_acronym_form.controls.Add($separator_bar);

    $y_pos = $y_pos + 5;

    #write-host Header $y_pos

    if($item_number -ne 0)
    {
        #####################################################################################
        foreach($list in $Script:acronym_lists.getEnumerator() | sort Key)
        {
            $list_file = "$dir\Resources\Acronym Lists\" + $list.Key
            $list_name = $list.Key -replace ".csv$", ""


            $list_name_label                          = New-Object system.Windows.Forms.Label
            $list_name_label.text                     = "$list_name";
            $list_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $list_name_label.Anchor                   = 'top,right'
            $list_name_label.width                    = (($edit_acronym_form.width - 50) - $spacer)
            $list_name_label.height                   = 30
            $list_name_label.location                 = New-Object System.Drawing.Point(20,$y_pos)
            $list_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $edit_acronym_form.controls.Add($list_name_label);

            $y_pos = $y_pos + 30;
                    
            $edit_button           = New-Object System.Windows.Forms.Button
            $edit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $edit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $edit_button.Width     = 120
            $edit_button.height     = 25
            $edit_button.Location  = New-Object System.Drawing.Point(20,$y_pos);
            $edit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $edit_button.Text      = "Manual Edit"
            $edit_button.Name      = $list_file 
            $edit_button.Add_Click({
                $message = "Making edits to Acronym and/or Abbreviations lists:`n - Must be kept in .csv file format`n - Must remain only two columns`n - Shorthand in column 1`n - Longhand in column 2`n - Is case sensitive (lower case ideal in most cases)"
                [System.Windows.MessageBox]::Show($message,"!!!WARNING!!!",'Ok')
                explorer.exe $this.name
            });
            $edit_acronym_form.controls.Add($edit_button) 

            $delete_button           = New-Object System.Windows.Forms.Button
            $delete_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $delete_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $delete_button.Width     = 90
            $delete_button.height     = 25
            $delete_button.Location  = New-Object System.Drawing.Point(($edit_button.Location.x + $edit_button.width + 5),$y_pos);
            $delete_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $delete_button.Text      ="Delete"
            $delete_button.Name      = $list_file 
            $delete_button.Add_Click({
                $file = [System.IO.Path]::GetFileNameWithoutExtension($this.name)
                $message = "Are you sure you want to delete the `"$file`" list? You cannot revert this action.`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    if(Test-Path -LiteralPath $this.name)
                    {
                        Remove-Item -LiteralPath $this.name
                    }
                    $Script:acronym_lists.remove("$file.csv");
                    build_acronym_menu
                    $script:reload_function = "manage_acronyms_dialog"
                    $edit_acronym_form.close();         
                }

            });
            $edit_acronym_form.controls.Add($delete_button)

            $rename_button           = New-Object System.Windows.Forms.Button
            $rename_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $rename_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $rename_button.Width     = 90
            $rename_button.height     = 25
            $rename_button.Location  = New-Object System.Drawing.Point(($delete_button.Location.x + $delete_button.width + 5),$y_pos);
            $rename_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $rename_button.Text      ="Rename"
            $rename_button.Name      = $list_file 
            $rename_button.Add_Click({
                $old_name = $this.name
                $new_name = rename_dialog $old_name

                #write-host ON $old_name
                #write-host NN $new_name
                
                if(($new_name -cne $old_name) -and ($new_name -ne ""))
                {
                    $old_key = [System.IO.Path]::GetFileNameWithoutExtension($old_name)
                    $new_key = [System.IO.Path]::GetFileNameWithoutExtension($new_name)
                    $old_value = $Script:acronym_lists["$old_key.csv"]
                    #write-host OV $old_value
                    $Script:acronym_lists.remove("$old_key.csv");
                    $Script:acronym_lists.add("$new_key.csv",$old_value);
                    build_acronym_menu    
                    $script:reload_function = "manage_acronyms_dialog"
                    $edit_acronym_form.close();
                }
            });
            $edit_acronym_form.controls.Add($rename_button)
            


            $enable_checkbox = new-object System.Windows.Forms.checkbox
            $enable_checkbox.Location = new-object System.Drawing.Size(($rename_button.Location.x + $rename_button.width + 5),$y_pos);
            $enable_checkbox.Size = new-object System.Drawing.Size(100,30)
            $enable_checkbox.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $enable_checkbox.name = $list.key          
            $enable_checkbox.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            if($list.value -eq "0")
            {
                $enable_checkbox.Checked = $false
                $enable_checkbox.text = "Disabled"
            }
            else
            {
                $enable_checkbox.Checked = $true
                $enable_checkbox.text = "Enabled"
            }
            $enable_checkbox.Add_CheckStateChanged({
                if($this.Checked -eq $true)
                {
                    $this.text = "Enabled"
                    $Script:acronym_lists[$this.name] = 1;
                    build_acronym_menu
                }
                else
                {
                    $this.text = "Disabled"
                    $Script:acronym_lists[$this.name] = 0;
                    build_acronym_menu
                }
            })
            $edit_acronym_form.controls.Add($enable_checkbox);


            #######################################################
            $line_count = 0
            $reader = New-Object IO.StreamReader $list_file
            while($null -ne ($line = $reader.ReadLine()))
            {
                if((!($line -match "Acronym|Meaning|Abbrevation")) -and ($line -match ","))
                {
                    #write-host $line
                    $line_count++;
                }
            }
            $reader.Close() 
            $item_count_label                          = New-Object system.Windows.Forms.Label
            $item_count_label.text                     = "$line_count Items";
            $item_count_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $item_count_label.Anchor                   = 'top,right'
            $item_count_label.TextAlign = "MiddleRight"
            $item_count_label.width                    = 110
            $item_count_label.height                   = 30
            $item_count_label.location                 = New-Object System.Drawing.Point((($edit_acronym_form.width - 140) - $spacer),$y_pos);
            $item_count_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            $edit_acronym_form.controls.Add($item_count_label);

            $y_pos = $y_pos + 30
            $separator_bar                             = New-Object system.Windows.Forms.Label
            $separator_bar.text                        = ""
            $separator_bar.AutoSize                    = $false
            $separator_bar.BorderStyle                 = "fixed3d"
            #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
            $separator_bar.Anchor                      = 'top,left'
            $separator_bar.width                       = (($edit_acronym_form.width - 50) - $spacer)
            $separator_bar.height                      = 1
            $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
            $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $separator_bar.TextAlign                   = 'MiddleLeft'
            $edit_acronym_form.controls.Add($separator_bar);
            $y_pos = $y_pos + 5
        }
    
        $edit_acronym_form.ShowDialog()
    }
    else
    {
        $message = "You have no Acronyms and/or Abbreviations lists to edit.`nYou must create or import a list first."
        #[System.Windows.MessageBox]::Show($message,"No List",'Ok')

        $error_label                          = New-Object system.Windows.Forms.Label
        $error_label.text                     = "$message";
        $error_label.ForeColor                = "Red"
        $error_label.Anchor                   = 'top,right'
        $error_label.width                    = ($edit_acronym_form.width - 10)
        $error_label.height                   = 50
        $error_label.TextAlign = "MiddleCenter"
        $error_label.location                 = New-Object System.Drawing.Point(10,$y_pos)
        $error_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $edit_acronym_form.controls.Add($error_label);
        $edit_acronym_form.ShowDialog()
    }
    
}
################################################################################
######Rename Dialog#############################################################
function rename_dialog($input_file)
{
    $input_name = [System.IO.Path]::GetFileNameWithoutExtension($input_file)

    $rename_form = New-Object System.Windows.Forms.Form
    $rename_form.FormBorderStyle = 'Fixed3D'
    $rename_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $rename_form.Location = new-object System.Drawing.Point(0, 0)
    $rename_form.Size = new-object System.Drawing.Size(400, 120)
    $rename_form.MaximizeBox = $false
    $rename_form.SizeGripStyle = "Hide"
    $rename_form.Text = "Rename `"$input_name`""
    #$rename_form.TopMost = $True
    $rename_form.TabIndex = 0
    $rename_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $rename_name_label                          = New-Object system.Windows.Forms.Label
    $rename_name_label.text                     = "New Name:";
    $rename_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $rename_name_label.Anchor                   = 'top,right'
    $rename_name_label.width                    = 120
    $rename_name_label.height                   = 30
    $rename_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $rename_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))

    $rename_form.controls.Add($rename_name_label);

    $rename_name_input                         = New-Object system.Windows.Forms.TextBox                       
    $rename_name_input.AutoSize                 = $true
    $rename_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $rename_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $rename_name_input.Anchor                   = 'top,left'
    $rename_name_input.width                    = 250
    $rename_name_input.height                   = 30
    $rename_name_input.location                 = New-Object System.Drawing.Point(($rename_name_label.Location.x + $rename_name_label.Width + 5) ,12)
    $rename_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $rename_name_input.text                     = "$input_name"
    $rename_name_input.name                     = "$input_file"
    $rename_name_input.Add_TextChanged({
        $caret = $rename_name_input.SelectionStart;
        #$rename_name_input.text = $rename_name_input.text -replace '[^0-9A-Za-z ,-]', ''
        $rename_name_input.text = $rename_name_input.text.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '

        #$rename_name_input.text = (Get-Culture).TextInfo.ToTitleCase($rename_name_input.text)
        $rename_name_input.SelectionStart = $caret
    });
    $rename_form.controls.Add($rename_name_input);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($rename_form.width / 2) - ($submit_button.width)),45);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Save"
    $submit_button.Name = ""
    $submit_button.Add_Click({ 
        [array]$errors = "";
        $original = $rename_name_input.name
        $file_base = [System.IO.Path]::GetDirectoryName($rename_name_input.name)
        $file_ext = [System.IO.Path]::GetExtension($rename_name_input.name)
        $full_path = "$file_base\" +  $rename_name_input.text + $file_ext
        #write-host $full_path
        #write-host $file_base
        #write-host $file_ext

        if(!($full_path -ceq "$original"))
        { 
            if($full_path -eq "$original")
            {
                #User only changed case of files
            }
            elseif(Test-path "$full_path")
            {
                $errors += "File already exists."
            }
            

            if($rename_name_input.text -eq "")
            {
                $errors += "You must provide a name."
            }

            
            if($errors.count -eq 1)
            {
                $message = "Are you sure you want to save changes?`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Overwrite?", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    Rename-Item -LiteralPath $original -NewName $full_path -Force
                    $submit_button.Name = "$full_path"
                    $null = $rename_form.close();
                    
                    
                }
            }
            else
            {
                $message = "Please fix the following errors:`n`n"
                $counter = 0;
                foreach($error in $errors)
                {
                    if($error -ne "")
                    {
                        $counter++;
                        $message = $message + "$counter - $error`n"
                    } 
                }
                [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
            }
        }
        else
        {
            $submit_button.Name = "$full_path"
            $null = $rename_form.close();
        }


    });
    $rename_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($rename_form.width / 2)),45);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $rename_form.close();
    });
    $rename_form.controls.Add($cancel_button) 


    $null = $rename_form.ShowDialog()
    return $submit_button.Name
    
}
################################################################################
######XLS to CSV Conversion#####################################################
function xls_to_csv($input_file,$output_file,$mode)
{
    #write-host Input = $input_file
    #write-host Output = $output_file
    #write-host Mode = $mode
    #Mode 1 = First Worksheet Only
    #Mode 2 = Export Each Spreadsheet individually
    #Mode 3 = Merge All sheets into 1 Excel sheet
    if($mode -eq "")
    {
             $mode = 1;
    }
    $base_name = [io.path]::GetFileNameWithoutExtension($output_file)
    $base_directory = [System.IO.Path]::GetDirectoryName($output_file);
    $output_names = @{};
    #write-host Mode: $mode
    #write-host Input: $input_file
    #write-host Output: $output_file
    #write-host Base Name $base_name
    #write-host Base DIR  $base_directory 

    if(Test-path "$input_file")
    {
        #######Get Active Processes#######
        $user_excels = Get-Process EXCEL -ErrorAction SilentlyContinue
        #write-host Process: $user_excels.id
        ##################################

        $objExcel = New-Object -ComObject Excel.Application
        $workbook = $objExcel.Workbooks.Open("$input_file") 

        $objExcel.Visible=$false
        $objExcel.DisplayAlerts = $False

        if($mode -eq 1)
        {
            #$workbook.SaveAs($output_file,62)
            #####################
            try
            {
                $workbook.SaveAs($output_file,62); #Unicode CSV
            }
            catch
            {
                write-host "WARNING: Failed Unicode Saving: $save_name"
                try
                {
                    $workbook.SaveAs($output_file,6);#Standard CSV
                }
                catch
                {
                    write-host "FATAL ERROR: Failed Unicode Saving: $save_name"
                }
            }
            #####################
            $objExcel.DisplayAlerts = $True
            $output_names.add($output_file,"");
            $objExcel.Quit()
            ##############################
            ######Exit Excel Forcefully
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
            #[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)

            Remove-Variable objExcel
            Remove-Variable workbook
            #Remove-Variable ws

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            ##############################
             
        }
        if($mode -match "^2$|^3$")
        {
            foreach ($ws in $workbook.Worksheets)
            {
                $save_name = "$base_directory\" + $base_name + "_" + $ws.Name + ".csv"
                #####################
                try
                {
                    $ws.SaveAs($save_name,62); #Unicode CSV
                }
                catch
                {
                    write-host "WARNING: Failed Unicode Saving: $save_name"
                    try
                    {
                        $ws.SaveAs($save_name,6);#Standard CSV
                    }
                    catch
                    {
                        write-host "FATAL ERROR: Failed Unicode Saving: $save_name"
                    }
                }
                #####################
                $output_names.add($save_name,"");
            }
            $objExcel.Quit()
            ######Exit Excel Forcefully
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)

            Remove-Variable objExcel
            Remove-Variable workbook
            Remove-Variable ws

            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
            ##############################         
        }
        if($mode -eq 3)
        {
            if(Test-Path $output_file)
            {
                Remove-Item -LiteralPath $output_file
            }
            get-childItem "$base_directory\*.csv" | foreach {
                $reader = [System.IO.File]::AppendAllText("$output_file",[System.IO.File]::ReadAllText($_.FullName))
                if($reader)
                {
                    $reader.close();
                }
            }
            foreach($file in $output_names.GetEnumerator())
            {
                remove-item -LiteralPath $file.key
            }
            $output_names.clear();
            $output_names.add($output_file,"");
        }
        #######Kill Process##############
        #write-host "Killing Excels"
        $current_excels = Get-Process EXCEL -ErrorAction SilentlyContinue
        foreach($open_excel in $current_excels)
        {     
            if(!($user_excels.id -contains $open_excel.id))
            {
               $killer = taskkill /PID $open_excel.id /F
            }
        }
        ##################################
        #write-host "finished Killing"
    }
    else
    {
        write-host "Critical Error: Input Spreadsheet does not exist"
    }

    ############Fix Excel UTF to Unicode
    #foreach($file in $output_names.GetEnumerator())
    #{    
    #    $content = Get-Content -literalpath $file.Key -Encoding UTF8
    #    Set-Content -literalpath $file.key $content -Encoding Ascii   
    #}

    #write-host Finished Excel Processing

    return $output_names
}
################################################################################
######DOC to TXT Conversion#####################################################
function doc_to_txt($input_file,$output_file)
{
    #######Get Active Processes#######
    $user_words = Get-Process WINWORD -ErrorAction SilentlyContinue
    #write-host Process: $user_words.id
    ##################################

    $Word = New-Object -ComObject Word.Application
    $Document = $Word.Documents.Open($input_file)
    $def = [Type]::Missing
    $Document.SaveAs([ref]$output_file,[ref] 7,$def,$def,$def,$def,$def,$def,$def,$def,$def,65001)
    $Document.Close()
    $Word.Quit()
    ##############################
    ######Exit Word Forcefully
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Document)

    Remove-Variable Word
    Remove-Variable Document

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    ##############################

    ###Convert UTF8 to Unicode
    $MyRawString = Get-Content -Raw $output_file
    $Encoding = New-Object System.Text.UnicodeEncoding
    [System.IO.File]::WriteAllLines($output_file, $MyRawString,$Encoding)

    #######Kill Process##############
    #write-host "Killing words"
    $current_words = Get-Process WINWORD -ErrorAction SilentlyContinue
    foreach($open_word in $current_words)
    {     
        if(!($user_words.id -contains $open_word.id))
        {
            $killer = taskkill /PID $open_word.id /F
        }
    }
    ##################################
}
################################################################################
######Prompt for File###########################################################
function prompt_for_file()
{  
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")|Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
 #$OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.filter = "Excel Worksheets (*.xls, *.xlsx, *.csv)|*.xls;*.xlsx;*.csv"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
}
################################################################################
######Save Bullet Tracker#######################################################
function save_bullet_tracker
{
    ###Update Tracker File
    foreach($bank in $script:Bullet_banks.getEnumerator())
    {
        $file = $bank.key
        $status = $bank.value
        if(Test-Path -LiteralPath "$dir\Resources\Bullet Banks\$file")
        {
            $line = $file + "::" + $status
            Add-Content "$dir\Resources\Required\Bullet_lists_temp.txt" $line
        }
    }
    if(Test-Path -LiteralPath "$dir\Resources\Required\Bullet_lists_temp.txt")
    {
        if(Test-Path -LiteralPath "$dir\Resources\Required\Bullet_lists.txt")
        {
            Remove-Item -LiteralPath "$dir\Resources\Required\Bullet_lists.txt"
        }
        Rename-Item -LiteralPath "$dir\Resources\Required\Bullet_lists_temp.txt" "$dir\Resources\Required\Bullet_lists.txt"
    }
}
################################################################################
######Save Acronym Tracker######################################################
function save_acronym_tracker
{
    
    ###Update Tracker File
    foreach($list in $Script:acronym_lists.getEnumerator())
    {
        $file = $list.key
        $status = $list.value
        if(Test-Path -LiteralPath "$dir\Resources\Acronym Lists\$file")
        {
            $line = $file + "::" + $status
            Add-Content "$dir\Resources\Required\Acronym_lists_temp.txt" $line
        }
    }
    if(Test-Path -LiteralPath "$dir\Resources\Required\Acronym_lists_temp.txt")
    {
        if(Test-path -LiteralPath "$dir\Resources\Required\Acronym_lists.txt")
        {
            Remove-Item -LiteralPath "$dir\Resources\Required\Acronym_lists.txt"
        }
        Rename-Item -LiteralPath "$dir\Resources\Required\Acronym_lists_temp.txt" "$dir\Resources\Required\Acronym_lists.txt"
    }
}
################################################################################
######Save Bullet Tracker#######################################################
function save_package_tracker
{
    ###Update Tracker File
    foreach($package in $script:package_list.getEnumerator())
    {
        $file = $package.key
        $status = $package.value
        if(Test-Path -LiteralPath "$dir\Resources\Packages\$file")
        {
            $line = $file + "::" + $status
            Add-Content "$dir\Resources\Required\Package_list_temp.txt" $line
        }
    }
    if(Test-Path -LiteralPath "$dir\Resources\Required\Package_list_temp.txt")
    {
        if(Test-Path -LiteralPath "$dir\Resources\Required\Package_list.txt")
        {
            Remove-Item -LiteralPath "$dir\Resources\Required\Package_list.txt"
        }
        Rename-Item -LiteralPath "$dir\Resources\Required\Package_list_temp.txt" "$dir\Resources\Required\Package_list.txt"
    }
}
################################################################################
######Right Click Menu##########################################################
function right_click_menu
{
    $contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
    $editor.ContextMenuStrip = $contextMenuStrip1
    $contextMenuStrip1.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
    $contextMenuStrip1.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
    if($editor.SelectedText.Length -ge 1)
    {
        $contextMenuStrip1.Items.Add("Cut").add_Click({clipboard_cut})
                 
        $contextMenuStrip1.Items.Add("Copy").add_Click({clipboard_copy})
    }       
    $contextMenuStrip1.Items.Add("Paste").add_Click({clipboard_paste})
    
    
    ############################################################################
    ###Find What word the user is clicking on
    $simplified_text = $editor.text -replace "[^a-z0-9]| | | ",' '
    #write-host $simplified_text
    $front = "";
    $index_end = $editor.SelectionStart;
    $back = ""
    $index_start = $editor.SelectionStart;
    For ($i=0; $i -le ($editor.text.Length - $editor.SelectionStart); $i++) 
    {
        $temp = "";
        if(!(($editor.SelectionStart + $i) -ge $editor.text.Length))
        {
            $temp = $simplified_text.Substring(($editor.SelectionStart + $i),1);
        }
        $front = $front + $temp
        $index_end = $editor.SelectionStart + $i
        if($front -match " $")
        {
            break
        } 
    }        
    For ($i = $editor.SelectionStart - 1; $i -ge 0; $i--) 
    {
        $temp = $simplified_text.Substring($i,1);
        $back =  $temp + $back
        $index_start = $i
        if($back -match "^ ")
        {
            $index_start++
            break
        } 
    }       
    ############################################################################
    $word = $simplified_text.substring($index_start,($index_end -$index_start))
    #write-host "$index_start = $back -- " $editor.SelectionStart "-- $front = $index_end"
    #write-host WORD -$word-       
            
    ###################################################################
    ##Acronyms
    $first = 0; #Build menu only on first find
    $thesaurus_word = "";
    foreach($entry in $script:acro_index.getEnumerator())
    {
                
        if(($entry.value -ge $index_start) -and ($entry.value -le $index_end))
        {
            ($mode,$index,$acronym,$meaning) = $entry.key -split "::"
            #write-host $mode,$index,$acronym,$meaning
            if($mode -eq "E")
            {
                if($first -eq 0)
                {
                    $separator = [System.Windows.Forms.ToolStripSeparator]::new()
	                $contextMenuStrip1.Items.Add($separator)
                    $first++;
	                $acronym_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                    $acronym_menu.text = "Extend"
                    $acronym_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                    $acronym_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
	                $contextMenuStrip1.Items.Add($acronym_menu)
                    $thesaurus_word = $meaning
                }
                $word_item = [System.Windows.Forms.ToolStripMenuItem]::new()
                $word_item.text = "$meaning" -replace '&','&&';
                $word_item.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $word_item.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $word_item.name = "$index_start" + "::" + ($index_start + $acronym.length) + "::" +  "$meaning"
                $word_item.add_Click({
                    replace_text $this.name
                })
	            $acronym_menu.DropDownItems.Add($word_item)
                        
            }
            if($mode -eq "S")
            {
                if($first -eq 0)
                {
                    $separator = [System.Windows.Forms.ToolStripSeparator]::new()
	                $contextMenuStrip1.Items.Add($separator)
                    $first++;
	                $acronym_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                    $acronym_menu.text = "Shorten"
                    $acronym_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                    $acronym_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
	                $contextMenuStrip1.Items.Add($acronym_menu)
                    $thesaurus_word = $acronym
                }
                $word_item = [System.Windows.Forms.ToolStripMenuItem]::new()
                $word_item.text = "$meaning"  -replace '&','&&';
                $word_item.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $word_item.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $word_item.name = "$index_start" + "::" + ($index_start + $acronym.length) + "::" +  "$meaning"
                $word_item.add_Click({
                    replace_text $this.name
                })
	            $acronym_menu.DropDownItems.Add($word_item)
            }             
        }
    }
    $add_acromenu = 0;
    if(($editor.SelectedText.Length -ge 1) -and ($editor.SelectedText.Length -le 50))
    {
        $dic_menu = [System.Windows.Forms.ToolStripSeparator]::new()
	    $contextMenuStrip1.Items.Add($dic_menu)

        $dic_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
        $dic_menu.text = "Add to Acronyms/Abbreviations"
        $dic_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $dic_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        $dic_menu.name = $editor.SelectedText
        $dic_menu.add_Click({
            add_to_acronyms $this.name
        })
	    $contextMenuStrip1.Items.Add($dic_menu)
        $add_acromenu = 1;
    }
    ###################################################################
    ##Dictionary

    #Find and Rack & Stack Closest Words
    $dic_rack_an_stack = @{};
    if($script:dictionary_index.containskey($word))
    {

        if($script:dictionary_index[$word] -ne "C")
        {
            if($add_acromenu -eq 0)
            {
                $dic_menu = [System.Windows.Forms.ToolStripSeparator]::new()
	            $contextMenuStrip1.Items.Add($dic_menu)

                $dic_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $dic_menu.text = "Add to Acronyms/Abbreviations"
                $dic_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $dic_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $dic_menu.name = $word
                $dic_menu.add_Click({
                    add_to_acronyms $this.name
                })
	            $contextMenuStrip1.Items.Add($dic_menu)
            }
            
            $dic_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $dic_menu.text = "Add to Dictionary"
            $dic_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $dic_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $dic_menu.name = $word
            $dic_menu.add_Click({
                add_to_dictionary $this.name
            })
	        $contextMenuStrip1.Items.Add($dic_menu)

            $dic_menu = [System.Windows.Forms.ToolStripSeparator]::new()
	        $contextMenuStrip1.Items.Add($dic_menu)

            if($script:dictionary_index[$word] -ne "M")
            {
                $word_match_array = $script:dictionary_index[$word] -split '::'

                foreach($match in $word_match_array)
                {
                    $dic_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                    $dic_menu.text = $match
                    $dic_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                    $dic_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                    $dic_menu.name = "$index_start" + "::" + ($index_start + $word.length)  + "::" + $match
                    $dic_menu.add_Click({
                        replace_text $this.name
                    })
	                $contextMenuStrip1.Items.Add($dic_menu)
                }   
            }
            else
            {
                foreach($line in $script:dictionary.getEnumerator()) 
                {
                    if(($line.key.Substring(0,1) -eq $word.Substring(0,1)) -and ($line.key.Substring(($line.key.length -1),1) -eq $word.Substring(($word.length -1),1)))
                    {     
                        $score = levenshtein $line.key $word
                        if(($line.key.length -ge 2) -and ($word.length -ge 2) -and ($line.key.Substring(1,1) -eq $word.Substring(1,1)))
                        {
                            $score = $score - 2
                        }
                        if(($line.key.length -ge 3) -and ($word.length -ge 3) -and ($line.key.Substring(2,1) -eq $word.Substring(2,1)))
                        {
                            $score = $score - 2
                        }
                        if(($line.key.length -ge 4) -and ($word.length -ge 4) -and ($line.key.Substring(3,1) -eq $word.Substring(3,1)))
                        {
                            $score = $score - 2
                        }
                        if(($line.key.length -ge 2) -and ($word.length -ge 2) -and ($line.key.Substring(($line.key.length -2),1) -eq $word.Substring(($word.length -2),1)))
                        {
                            $score = $score - 2
                        }
                        [int]$distance = (($word.length - $line.key.length)  + 1)
                        $score = ((($distance * $distance) / 2) + $score)


                        if(!($dic_rack_an_stack.containskey($line.key)))
                        {
                            #write-host $line.key = $word = $score = $distance
                            $dic_rack_an_stack.add($line.key,$score);
                        }    
                    }
                }
            }
            ################
            $capacity = 16
            foreach($find in $dic_rack_an_stack.getEnumerator() | Sort Value) 
            {
                $capacity--;
                if($capacity -eq 0)
                {
                    break;
                }
                $dic_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $dic_menu.text = $find.key
                $dic_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $dic_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $dic_menu.name = "$index_start" + "::" + ($index_start + $word.length)  + "::" + $find.key
                $dic_menu.add_Click({
                    replace_text $this.name
                })
	            $contextMenuStrip1.Items.Add($dic_menu)

                if($script:dictionary_index[$word] -eq "M")
                {
                    $script:dictionary_index[$word] = "";
                }

                $script:dictionary_index[$word] = $script:dictionary_index[$word] + $find.key + "::"

                #write-host $find.key = $find.value = $word 
            }
            $script:dictionary_index[$word] = $script:dictionary_index[$word] -replace '::$',''

        } 
    }
    ############################################################################
    #Thesaurus
    $script:global_thesaurus = @{};
    $script:global_word_hippo = @{};
    if($editor.SelectedText.length -eq 0)
    {
        if(($script:dictionary_index.containskey($word)) -and ($script:dictionary_index[$word] -eq "C") -or ($thesaurus_word -ne ""))
        {
            if($thesaurus_word -ne "")
            {
                $thesaurus_word = $thesaurus_word -split ' '
                $word = $thesaurus_word[0]
            }
            $words = thesaurus_lookup $word
            $separator = [System.Windows.Forms.ToolStripSeparator]::new()
	        $contextMenuStrip1.Items.Add($separator)
            $script:thesaurus_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $thesaurus_menu.text = "Thesaurus"
            $thesaurus_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $thesaurus_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $thesaurus_menu.name = "$index_start::$word"
	        $contextMenuStrip1.Items.Add($thesaurus_menu)
            $thesaurus_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $thesaurus_sub_menu.text = "Loading..."
            $thesaurus_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $thesaurus_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $thesaurus_sub_menu.name = "Loading..."
            $thesaurus_sub_menu.enabled = $false
	        $thesaurus_menu.DropDownItems.Add($thesaurus_sub_menu)

            ##############Word Hippo Lookup

            $words = word_hippo_lookup $word
            $separator = [System.Windows.Forms.ToolStripSeparator]::new()
	        $contextMenuStrip1.Items.Add($separator)
            $script:word_hippo_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $word_hippo_menu.text = "Word Hippo"
            $word_hippo_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $word_hippo_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $word_hippo_menu.name = "$index_start::$word"
	        $contextMenuStrip1.Items.Add($word_hippo_menu)
            $word_hippo_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
            $word_hippo_sub_menu.text = "Loading..."
            $word_hippo_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $word_hippo_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
            $word_hippo_sub_menu.name = "Loading..."
            $word_hippo_sub_menu.enabled = $false
	        $word_hippo_menu.DropDownItems.Add($word_hippo_sub_menu)

        }
    }
}
################################################################################
#######Thesaurus Lookup#########################################################
function thesaurus_lookup($lookup_word)
{
    if(($script:thesaurus_job -eq "") -and ($lookup_word -ne ""))
    {
        #write-host Looking up $lookup_word
        #########################################################
        ###Start Job
        $script:thesaurus_job = Start-Job -ScriptBlock {

            $local_thesaurus = $using:global_thesaurus
            $word = $using:lookup_word
            $dir = $using:dir

            #####################################################################
            ###########Duplicated Code Because Jobs cannot see outside of scope
            function csv_line_to_array ($line)
            {
                if($line -match "^,")
                {
                    $line = ",$line"; 
                }
                Select-String '(?:^|,)(?=[^"]|(")?)"?((?(1)[^"]*|[^,"]*))"?(?=,|$)' -input $line -AllMatches | Foreach { $line_split = $_.matches -replace '^,|"',''}
                [System.Collections.ArrayList]$line_split = $line_split
                return $line_split
            }
            #####################################################################
            if(Test-Path "$dir\Resources\Required\Thesaurus.csv")
            {
                [int]$scan_attempts = 0;
                $search_word = $word.ToLower();
                $search_word2 = $search_word + "s"
                $search_word3 = $search_word + "ed"
                $search_word4 = $search_word -replace 'ed$', 'e';
                
                $max_tries = 8;       
                while($scan_attempts -le $max_tries)
                {
                    #write-host $scan_attempts
                    #write-host "-$search_word-"
                    $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Thesaurus.csv")
                    while($null -ne ($line = $reader.ReadLine()))
                    {
                        if(($line -match ",$search_word,|,$search_word2,|,$search_word3,|,$search_word4,") -or (($scan_attempts -ge 4) -and ($line -match ",$word")))
                        {
                            [Array]$line_split = csv_line_to_array $line
                            foreach($line_word in $line_split)
                            {
                                if(($line_word -ne "") -and ($line_word -ne "$word") -and ($line_word -ne "$search_word2") -and ($line_word -ne "$search_word3") -and ($line_word -ne "$search_word4"))
                                {
                                    if($local_thesaurus.contains($line_word))
                                    {
                                        $local_thesaurus[$line_word] = $local_thesaurus[$line_word] + 1;
                                        #write-host Double Down
                                    }
                                    else
                                    {
                                        #write-host Found $line_word
                                        $local_thesaurus.Add($line_word,0)
                                        $scan_attempts = $max_tries;
                                        #write-host Found
                                    }
                                }
                            }
                        }
                    }
                    $reader.Close();
                    if(($scan_attempts -ne $max_tries) -and ($scan_attempts -ge 2))
			        {
                        
                        $search_word = $search_word.substring(0,($search_word.length - 1));
                        if($search_word.Length -le 2)
                        {
                            $scan_attempts = $max_tries

                        }
                    }
                    $scan_attempts++;
                }
            
            }
            $max_words = 25
            $counter = 0;
            $small_list = @{};
            foreach($synonym  in $local_thesaurus.getEnumerator() | Sort Value -Descending) 
            {
                $counter++;
                $small_list.add($synonym.key,$synonym.value);
                if($counter -eq $max_words)
                {
                    break;
                } 
            }

        return $small_list
        }
    }
    else
    {
        #########################################################
        ###Job Finished
        if($script:thesaurus_job.state -eq "Completed")
        {
            $script:global_thesaurus = Receive-Job -Job $script:thesaurus_job
            if(($script:thesaurus_job.state -eq "Completed") -and ($script:global_thesaurus.Get_Count() -ne 0))
            {   
                ([int]$index_start,$word) = $thesaurus_menu.name -split "::"
                #write-host Thesuarus $index_start - $word
                $thesaurus_menu.DropDownItems.clear();
                $max_words = 25
                $counter = 0;
                foreach($synonym  in $script:global_thesaurus.getEnumerator() | Sort Value -Descending) 
                {
                    #write-host $synonym.key - $synonym.value
                    $counter++;

                    $thesaurus_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                    $thesaurus_sub_menu.text = $synonym.key
                    $thesaurus_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                    $thesaurus_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                    $thesaurus_sub_menu.name = "$index_start" + "::" + ($index_start + $word.length)  + "::" + $synonym.key

                    #write-host $thesaurus_sub_menu.name
                    $thesaurus_sub_menu.add_Click({
                        replace_text $this.name
                    })
	                $thesaurus_menu.DropDownItems.Add($thesaurus_sub_menu)

                    if($counter -eq $max_words)
                    {
                        break;
                    } 
                }          
                $script:thesaurus_job = "";

            }
            elseif(($status -eq "Completed") -and ($script:global_thesaurus.Get_Count() -eq 0))
            {

                $thesaurus_menu.DropDownItems.clear();
                $thesaurus_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $thesaurus_sub_menu.text = "No Synonyms Found"
                $thesaurus_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $thesaurus_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $thesaurus_sub_menu.name = "No Synonyms Found"
                $thesaurus_sub_menu.enabled = $false
                $thesaurus_menu.DropDownItems.Add($thesaurus_sub_menu)
                ###Duplicate fixes glitch where menu pops on the top right hand corner of screen... dum...
                $thesaurus_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $thesaurus_sub_menu.text = "No Synonyms Found"
                $thesaurus_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $thesaurus_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $thesaurus_sub_menu.name = "No Synonyms Found"
                $thesaurus_sub_menu.enabled = $false
                $thesaurus_menu.DropDownItems.Add($thesaurus_sub_menu)
                $thesaurus_menu.DropDownItems.remove($thesaurus_sub_menu)

                $script:thesaurus_job = "";
            }
        }
    }

}
################################################################################
#######Word Hippo Lookup########################################################
function word_hippo_lookup($lookup_word)
{
    if(($script:word_hippo_job -eq "") -and ($lookup_word -ne ""))
    {
        #########################################################
        ###Start Job
        $script:word_hippo_job = Start-Job -ScriptBlock {

            $word = $using:lookup_word
            $word_list = New-Object system.collections.hashtable
            $word = $word.ToLower();
            $full_url = "https://www.wordhippo.com/what-is/another-word-for/$word.html"

            try
            {
                $response = Invoke-WebRequest -Uri $full_url
                $response = $response -split "`n"
                $found = 0;
                foreach($line in $response)
                {
                    $pattern = '<a href="'
                    if($line -match [regex]::Escape($pattern) -and $line.Length -lt 70)
                    {
                        $line_split = $line -split ('><a href="|.html">|</a></div>')

                        if(($line_split[1] -ne "") -and ($line_split[2] -ne ""))
                        {
                            if($line_split[1] -eq $line_split[2])
                            {
                                if(!($word_list.contains($line_split[1])))
                                {
                                    $word_list.add($line_split[1],"");
                                    $found++
                                    if($found -gt 25)
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                $word_list.add("Failed");
            }
            return $word_list
        }
    }
    else
    {
        #########################################################
        ###Job Finished
        if($script:word_hippo_job.state -eq "Completed")
        {


            $script:global_word_hippo = Receive-Job -Job $script:word_hippo_job
            if(($script:word_hippo_job.state -eq "Completed") -and ($script:global_word_hippo.Get_Count() -ne 0))
            {   
                
                ([int]$index_start,$word) = $word_hippo_menu.name -split "::"
                #write-host Word_hippo $index_start - $word
                $word_hippo_menu.DropDownItems.clear();
                $max_words = 25
                $counter = 0;
                foreach($synonym in $script:global_word_hippo.getEnumerator() | Sort Value -Descending) 
                {
                    #write-host $synonym.key - $synonym.value
                    $counter++;

                    $word_hippo_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                    $word_hippo_sub_menu.text = $synonym.key
                    $word_hippo_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                    $word_hippo_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                    $word_hippo_sub_menu.name = "$index_start" + "::" + ($index_start + $word.length)  + "::" + $synonym.key

                    #write-host $word_hippo_sub_menu.name
                    $word_hippo_sub_menu.add_Click({
                        replace_text $this.name
                    })
	                $word_hippo_menu.DropDownItems.Add($word_hippo_sub_menu)

                    if($counter -eq $max_words)
                    {
                        break;
                    } 
                }          
                $script:word_hippo_job = "";

            }
            elseif(($status -eq "Completed") -and ($script:global_word_hippo.Get_Count() -eq 0))
            {

                $word_hippo_menu.DropDownItems.clear();
                $word_hippo_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $word_hippo_sub_menu.text = "No Synonyms Found"
                $word_hippo_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $word_hippo_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $word_hippo_sub_menu.name = "No Synonyms Found"
                $word_hippo_sub_menu.enabled = $false
                $word_hippo_menu.DropDownItems.Add($word_hippo_sub_menu)
                ###Duplicate fixes glitch where menu pops on the top right hand corner of screen... dum...
                $word_hippo_sub_menu = [System.Windows.Forms.ToolStripMenuItem]::new()
                $word_hippo_sub_menu.text = "No Synonyms Found"
                $word_hippo_sub_menu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
                $word_hippo_sub_menu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
                $word_hippo_sub_menu.name = "No Synonyms Found"
                $word_hippo_sub_menu.enabled = $false
                $word_hippo_menu.DropDownItems.Add($word_hippo_sub_menu)
                $word_hippo_menu.DropDownItems.remove($word_hippo_sub_menu)

                $script:word_hippo_job = "";
            }
        }
    }
}
################################################################################
######Add to Dictionary#########################################################
function add_to_dictionary($word)
{
    $message = "Are you sure you want to add `"$word`" to the dictionary?`n`n"
    $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Add `"$word`" to dictionary?", "YesNo" , "Information" , "Button1")
    if($yesno -eq "Yes")
    {
        $word = $word.Tolower();
        Add-Content "$dir\Resources\Required\Dictionary.txt" "$word"
        $script:dictionary.Add($word,"");
        $script:dictionary_index[$word] = "C"
        $Script:recent_editor_text = "" #Force Reload
    }
}
################################################################################
######Add to Acronyms###########################################################
function add_to_acronyms($acronym,$meaning)
{
    if($acronym)
    {
        $acronym = $acronym.trim();
        $acronym = $acronym -replace " | | |`n`r|`n|`r",' '
    }
    if($meaning)
    {
        $meaning = $meaning.trim();
        $meaning = $meaning -replace " | | |`n`r|`n|`r",' '
    }

    $add_acronym_form = New-Object System.Windows.Forms.Form
    $add_acronym_form.FormBorderStyle = 'Fixed3D'
    $add_acronym_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $add_acronym_form.Location = new-object System.Drawing.Point(0, 0)
    $add_acronym_form.MaximizeBox = $false
    $add_acronym_form.SizeGripStyle = "Hide"
    $add_acronym_form.Size='285,340'
    $add_acronym_form.Text = "Add Acronym or Abbreviation"
    #$add_acronym_form.TopMost = $True
    $add_acronym_form.TabIndex = 0
    $add_acronym_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $acronym_shorthand_label                          = New-Object system.Windows.Forms.Label
    $acronym_shorthand_label.text                     = "Acronym:";
    #$acronym_shorthand_label.AutoSize                 = $true
    #$acronym_shorthand_label.BackColor                = 'Green'
    $acronym_shorthand_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $acronym_shorthand_label.Anchor                   = 'top,right'
    $acronym_shorthand_label.width                    = 100
    $acronym_shorthand_label.height                   = 30
    $acronym_shorthand_label.location                 = New-Object System.Drawing.Point(10,10)
    $acronym_shorthand_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $add_acronym_form.controls.Add($acronym_shorthand_label);

    $acronym_shorthand_label                          = New-Object system.Windows.Forms.Label
    $acronym_shorthand_label.text                     = "Meaning:";
    #$acronym_shorthand_label.AutoSize                 = $true
    #$acronym_shorthand_label.BackColor                = 'Green'
    $acronym_shorthand_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $acronym_shorthand_label.Anchor                   = 'top,right'
    $acronym_shorthand_label.width                    = 100
    $acronym_shorthand_label.height                   = 30
    $acronym_shorthand_label.location                 = New-Object System.Drawing.Point(10,45)
    $acronym_shorthand_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $add_acronym_form.controls.Add($acronym_shorthand_label);


    $acronym_shorthand_input                          = New-Object system.Windows.Forms.TextBox                       
    $acronym_shorthand_input.AutoSize                 = $true
    $acronym_shorthand_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $acronym_shorthand_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $acronym_shorthand_input.Anchor                   = 'top,left'
    $acronym_shorthand_input.width                    = 150
    $acronym_shorthand_input.height                   = 30
    $acronym_shorthand_input.location                 = New-Object System.Drawing.Point(115,12)
    $acronym_shorthand_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $acronym_shorthand_input.text                     = "$acronym"
    $add_acronym_form.controls.Add($acronym_shorthand_input);

    $acronym_longhand_input                          = New-Object system.Windows.Forms.TextBox                       
    $acronym_longhand_input.AutoSize                 = $true
    $acronym_longhand_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $acronym_longhand_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $acronym_longhand_input.Anchor                   = 'top,left'
    $acronym_longhand_input.width                    = 150
    $acronym_longhand_input.height                   = 30
    $acronym_longhand_input.location                 = New-Object System.Drawing.Point(115,47)
    $acronym_longhand_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $acronym_longhand_input.text                     = "$meaning"
    $add_acronym_form.controls.Add($acronym_longhand_input);

    $acronym_list_selection_label                          = New-Object system.Windows.Forms.Label
    $acronym_list_selection_label.text                     = "Applies to:";
    #$acronym_list_selection_label.AutoSize                 = $true
    #$acronym_list_selection_label.BackColor                = 'Green'
    $acronym_list_selection_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $acronym_list_selection_label.Anchor                   = 'top,right'
    $acronym_list_selection_label.width                    = 110
    $acronym_list_selection_label.height                   = 30
    $acronym_list_selection_label.location                 = New-Object System.Drawing.Point(10,80)
    $acronym_list_selection_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $add_acronym_form.controls.Add($acronym_list_selection_label);

    $acronym_list_selection                                = New-Object -TypeName System.Windows.Forms.CheckedListBox
    $acronym_list_selection.width                          = 255
    $acronym_list_selection.height                         = (100)
    $acronym_list_selection.Anchor                         = 'top,right'
    $acronym_list_selection.location                       = New-Object System.Drawing.Point(10,115)
    $acronym_list_selection.Items.Clear();
    #$acronym_list_selection.Backcolor                      = $script:settings['MEMBERS_LIST_BACKGROUND_COLOR']
    #$acronym_list_selection.Forecolor                      = $script:settings['MEMBERS_LIST_TEXT_COLOR']
	$acronym_list_selection.FormattingEnabled              = $True

    if($Script:acronym_lists.count -ne 0)
    {
        foreach($list in $Script:acronym_lists.getEnumerator())
        {
            $enabled = $list.value
            $list = $list.key -replace ".csv$",""
            [void] $acronym_list_selection.Items.add("$list")
            if($script:active_lists.contains("$list")) #Check the items that the user had checked last
            {
                $acronym_list_selection.SetItemChecked($acronym_list_selection.Items.IndexOf("$list"), $true);
            }
        }
    }
    else
    {
        $message = "You have no acronym list! You must create one before you can add any acronyms."
        [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
    }
    $acronym_list_selection.Add_ItemCheck({
        if($this.text)
        {
            #write-host $this.text
            $index = $acronym_list_selection.Items.IndexOf($this.text);
            #write-host $acronym_list_selection.GetItemChecked($index)
        }
    })

    $add_acronym_form.controls.Add($acronym_list_selection);

    $new_list_button           = New-Object System.Windows.Forms.Button
    $new_list_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $new_list_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $new_list_button.Width     = 155
    $new_list_button.height     = 25
    $new_list_button.Location  = New-Object System.Drawing.Point(10,210);
    $new_list_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $new_list_button.Text      ="Create New List"
    $new_list_button.Name      = "$acronym::$meaning"

    $new_list_button.Add_Click({
        create_acronym_list
       ($acronym,$meaning) =  $this.name -split "::"
       $script:reload_function = "add_to_acronyms"
       $script:reload_function_arg_a = $acronym;
       $script:reload_function_arg_b = $meaning; 
       $add_acronym_form.close();
        

    })
    $add_acronym_form.controls.Add($new_list_button);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($add_acronym_form.width / 2) - ($submit_button.width)),260);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Submit"
    $submit_button.Add_Click({ 
        [array]$errors = "";

        $acronym = $acronym_shorthand_input.text
        $acronym = $acronym -replace " | | |`n`r|`n|`r",' '
        $meaning = $acronym_longhand_input.text
        $meaning = $meaning -replace " | | |`n`r|`n|`r",' '
        if($Script:acronym_lists.count -eq 0)
        {
            $errors += "You must create an acronym list before you can add acronyms"
        }
        if($acronym.length -le 1)
        {
            $errors += "Acronym or Abbreviation too Short"
        }
        if($acronym.length -ge 40)
        {
            $errors += "Acronym or Abbreviation too Long"
        }
        if($meaning.length -le 2)
        {
            $errors += "Acronym or Abbreviation meaning too Short"
        }
        if($meaning.length -ge 150)
        {
            $errors += "Acronym or Abbreviation meaning too Long"
        }

        $script:active_lists = @{};
        foreach($list in $acronym_list_selection.Items)
        {
            $index = $acronym_list_selection.Items.IndexOf($list);
            if($acronym_list_selection.GetItemChecked($index) -eq $true)
            {
                $script:active_lists.add($list,"");
            }
        }

        if($script:active_lists.count -eq 0)
        {
            $errors += "You must select at least one list to add too"
        }
        if($errors.count -eq 1)
        {
            $message = "Are you sure you want to save changes?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Overwrite?", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                foreach($list in $script:active_lists.getEnumerator())
                {
                    $file= $list.key + ".csv";
                    $line = $acronym
                    $line = csv_write_line $line $meaning
                    if(test-path -literalpath "$dir\Resources\Acronym Lists\$file")
                    {     
                        Add-Content "$dir\Resources\Acronym Lists\$file" $line
                    }   
                }
                load_acronyms
                $Script:recent_editor_text = "Changed"
                $add_acronym_form.close();
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }
    });
    $add_acronym_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($add_acronym_form.width / 2)),260);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
            $add_acronym_form.close();
    });
    $add_acronym_form.controls.Add($cancel_button) 
    $add_acronym_form.ShowDialog()
}
################################################################################
######Replace Text##############################################################
function replace_text($request)
{
    $Script:LockInterface = 1;
    ($start, $end, $replace) = $request -split "::"

    $editor.SelectionStart = $start
    $editor.SelectionLength = ($end -$start)
    $editor.SelectedText = "$replace"

    $Script:CountDown=1

}
################################################################################
######Update Sizer Box##########################################################
function update_sizer_box
{
    #$script:bullets_and_sizes  = new-object System.Collections.Hashtable #Tracks bullet lengths
    $sizer_box.SelectionStart = 0;
    ###############################################
    ##Clean Up Lists
    $bullet_line_count  = $script:bullets_and_lines.get_count();
    $bullet_size_count  = $script:bullets_and_sizes.get_count();
    $editor_line_count = (([regex]::Matches($editor.text, "`n" )).count  + 1)
    $sizer_line_count = ([regex]::Matches($sizer_box.text, "`n" )).count

    #write-host BLC = $bullet_line_count
    #write-host BSC = $bullet_size_count
    #write-host ELC = $editor_line_count
    #write-host SLC = $sizer_line_count
    #write-host

    #################################################
    ###Remove Old Bullets From Bullets_and_Lines
    if($bullet_line_count -gt $editor_line_count)
    {
        #write-host Lines Deleted
        for($bullet_line_count; $bullet_line_count -gt $editor_line_count; $bullet_line_count--)
        {
            #write-host Delete $bullet_line_count
            $script:bullets_and_lines.remove($bullet_line_count);
        } 
    }
    #################################################
    ###Remove Old Bullets From Bullets_and_Sizes
    $remove_bullets = @{};
    foreach($bullet_size in $script:bullets_and_sizes.GetEnumerator())
    {
        $found = 0;
        foreach($bullet_lines in $script:bullets_and_lines.GetEnumerator())
        {
            if($bullet_lines.value -eq $bullet_size.key)
            {
                $found = 1;
                break;
            }
        }
        if($found -eq 0)
        {
            if(!($remove_bullets.contains($bullet_size.key)))
            {
                $remove_bullets.add($bullet_size.key,"");
            }
        }
    }
    foreach($remove in $remove_bullets.GetEnumerator())
    {
        $script:bullets_and_sizes.remove($remove.key);
        #write-host Removed $remove.key
    }

    ################################################
    ##Add Returns
    if($editor_line_count -gt $sizer_line_count)
    {
        $sizer_box.SelectionStart = 0
        $sizer_box.SelectionLength = 0;
        while($editor_line_count -gt $sizer_line_count)
        {
            $sizer_line_count++
            $sizer_box.SelectedText = "$sizer_line_count`n"
        }   
    }
    ###############################################
    ##Transfer Sizes
    $sizer_text_split = $sizer_box.text -split '\n'
    $counter = 0;
    $location = 0;
    $last_location = 0;
    foreach($sizer_line in $sizer_text_split)
    { 
        $counter++
        if($script:bullets_and_lines[$counter])
        {
            [int]$real_size = $script:bullets_and_sizes[$script:bullets_and_lines[$counter]]
            $sizer_box.SelectionStart = $location
            $sizer_box.SelectionLength = $sizer_line.length   
            if($real_size -le 2718)
            {
                $sizer_box.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'])
                #$sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            }
            else
            {
                $sizer_box.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'])
                #$sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            }
            
            if($script:settings['SIZER_BOX_INVERTED'] -eq 1)
            {
                $real_size = 2718 - $real_size
                $sizer_box.SelectedText = "$real_size"
            }
            else
            {
                $sizer_box.SelectedText = "$real_size"
            }
            $sizer_line = "$real_size"

        }
        else
        {
            $sizer_box.SelectionStart = $location
            $sizer_box.SelectionLength = $sizer_line.length
            $sizer_box.SelectedText = " "
            $sizer_line = " "  
        }
        if($counter -gt $editor_line_count)
        {
            ##Remove extra lines from Sizer Box
            $last_location = $location + 1;
            break;
        }
        $location = $location + $sizer_line.length + 1;

    }
    ##Finalize removal
    #$sizer_box.SelectionStart = $last_location
    #$sizer_box.SelectionLength = $sizer_box.text.length - $last_location
    #$sizer_box.SelectedText = " "

    for($i = 0;$i -lt 20;$i++)
    {
        ##Force Sizer Box Sync
        $editor.CustomVScroll()
    }
    while($sizer_box.ZoomFactor -ne $script:zoom) 
    {
        #Zoom Changes during RTF replace, but won't change in time... this is a work around.
        $sizer_box.ZoomFactor = $script:zoom
    }  
}
################################################################################
######Calculate Text Size New###################################################
function calculate_text_size_new($line)
{
    ##############################################################
    ##Get Line Length
    [double]$size = 0;
    for ($i = 0; $i -lt $line.length; $i++)
    {
        $character = $line[$i]
        if($character_blocks.Contains("$character"))
        {
            $size = $size + ($character_blocks["$character"])
        }
        else
        {
            write-host "Missing Character Block For `"$character`""
            #exit
        }
    }
    return $size
}
################################################################################
######Calculate Text Size#######################################################
function calculate_text_size
{

    
    #$editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE']) #added to fix font resets
    
    #############################################################################################################################################################
    ######Calculate Text Size


    ########################################################Text Compression setup
    if($script:space_hash.get_count() -ne 0)
    {
        $ghost_editor.text = $ghost_editor.text -replace " | | ", " "
    }


    $lines = $ghost_editor.text -split "`n";
    $text = "";
    $compressed_lines = new-object System.Collections.Hashtable
    $position = 0;
    foreach($line in $lines)
    {   
        ########################################################Get Current Line Size
        [double]$size = 0;
        for ($i = 0; $i -lt $line.length; $i++)
        {
            $character = $line[$i]
            if($character_blocks.Contains("$character"))
            {
                $size = $size + ($character_blocks["$character"])
            }
            else
            {
                write-host "Missing Character Block For `"$character`""
                #exit
            }
        }

        #######################################################Text Compression
        [int]$length = 2718 - $size
        if($script:space_hash.get_count() -ne 0)
        {
            #write-host $size = $length
            $matches = [regex]::Matches("$line"," ")
            $stop = 0;
            if($matches.Success)
            {  
                foreach($space_type in $space_hash.GetEnumerator() | sort value -Descending) #Loop 1
                {
                    #write-host Type: $space_type.value
                    if($stop -eq 1)
                    {
                        break
                    }
                    foreach($match in $matches) #Loop2
                    {
                        if($match.index -ne 1)
                        {
                            if(($stop -ne 1) -and ($size -gt 2718))
                            {
                                #$current_space = $line.substring($match.index,1)
                                #$current_size = $character_blocks["$current_space"]
                                #$size = $size - $current_size
                                #$size = $size + $space_type.Value
                                #$line = $line.remove($match.index,1)
                                #$line = $line.insert($match.index,$space_type.key)
                                        
                                $current_space = $ghost_editor.text.substring(($position + $match.index),1)                #Applies to entire block of text
                                $current_size = $character_blocks["$current_space"]                                        #Get size of space currently there
                                $size = $size - $current_size                                                              #Applies to line only
                                $size = $size + $space_type.Value                                                          #Applies to line only
                                $ghost_editor.text = $ghost_editor.text.remove(($position + $match.index),1)               #Applies to entire block of text
                                $ghost_editor.text = $ghost_editor.text.insert(($position + $match.index),$space_type.key) #Applies to entire block of text

                                #write-host $line
                            }
                            if($size -le 2718)
                            {
                                #write-host Compression Success
                                #write-host $line
                                $stop = 1;
                                [int]$size = $size
                                #write-host New Size: $size
                                break
                            }
                        }
                    }
                }
                if($stop -eq 0)
                {
                    #write-host Compression Failed
                    #write-host $line
                }
            }
        }
        #############################################################Shorthand 
        [int]$size = 2718 - $size
        if(($size -gt 2718))
        {
            [string]$size = ""
        }
        $text = $text + $size + "`n"
        $position = $position + $line.Length + 1;
    }

    #############################################################################
    ######Build a Ghost RTB
            
    $ghost_sizer_box.Multiline = $True
    $ghost_sizer_box.text = $text
    $ghost_sizer_box.SelectAll();
    $ghost_sizer_box.Forecolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'])
    $ghost_sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])          
    $ghost_sizer_box.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'])    
                     
    $pattern = "-\d+`n"
    $matches = [regex]::Matches("$text", $pattern)
    if($matches.Success)
    {  
        foreach($match in $matches)
        {
            $ghost_sizer_box.SelectionStart = $match.index
            $ghost_sizer_box.SelectionLength = $match.value.length
            $ghost_sizer_box.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'])
            $ghost_sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            $ghost_sizer_box.DeselectAll()
        }
    }
    $ghost_sizer_box.SelectAll()
    $ghost_sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    ##############################################################################
    #######Transfer to Real Sizer
    if($ghost_sizer_box.rtf -ne $sizer_box.Selectedrtf)
    {   
        $sizer_box_caret = $sizer_box.SelectionStart;
        $start = $sizer_box.SelectionStart
        $length = $sizer_box.SelectionLength
        $sizer_box.rtf = $ghost_sizer_box.rtf
        $sizer_box.SelectionStart = $sizer_box_caret
        $sizer_box.SelectionStart = $start;
        $sizer_box.SelectionLength = $length;
        
        while($sizer_box.ZoomFactor -ne $script:zoom) 
        {
            #Zoom Changes during RTF replace, but won't change in time... this is a work around.
            $sizer_box.ZoomFactor = $script:zoom
        }                             
    }
    ##############################################################################################################################################################
}
################################################################################
######Scan Text#################################################################
function scan_text
{
    $ghost_editor.Rtf = $editor.Rtf
    

    ########################################################
    ##Compression
    if($script:space_hash.get_count() -ne 0)
    {
        $lines = $ghost_editor.text -split "`n";
        $position = 0;
        $line_count = 0;
        foreach($line in $lines)
        {
            $line_count++;
            if($line -ne "")
            {
                if(!($script:bullets_compressed.contains($line)))
                {
                    ############################################
                    ##Get Current Bullet Size
                    $original = $line;
                    $line = $line -replace " | | ", " "
                    $ghost_editor.SelectionStart = $position
                    $ghost_editor.SelectionLength = $line.length
                    $ghost_editor.SelectedText = $line

                    $size = calculate_text_size_new $line
                
                    #write-host RL = $line = $size
                    ############################################
                    ##Attempt to Compress
                
                    $matches = [regex]::Matches("$line"," ")
                    $stop = 0;
                    if($matches)
                    {  
                        foreach($space_type in $space_hash.GetEnumerator() | sort value -Descending) #Loop 1
                        {
                            #write-host Type: $space_type.value
                            if($stop -eq 1)
                            {
                                break
                            }
                            foreach($match in $matches) #Loop2
                            {
                                if($match.index -ne 1)
                                {
                                    if(($stop -ne 1) -and ($size -gt 2718))
                                    {
                                        $current_space = $ghost_editor.text.substring(($position + $match.index),1)                #Applies to entire block of text
                                        $current_size = $character_blocks["$current_space"]   
                                        if($current_size -ne $space_type.value)
                                        {
                                            #Space is different size
                                            $line = $line.remove($match.index,1)
                                            $line = $line.insert($match.index,$space_type.key)
                                        
                                        
                                            $current_size = $character_blocks["$current_space"]                                        #Get size of space currently there
                                            #write-host CS = $current_size
                                            #write-host Pre $size
                                       
                                            $size = $size - $current_size                                                              #Applies to line only
                                            $size = $size + $space_type.Value    
                                            #write-host Pos $size                                                      #Applies to line only
                                            $ghost_editor.text = $ghost_editor.text.remove(($position + $match.index),1)               #Applies to entire block of text
                                            $ghost_editor.text = $ghost_editor.text.insert(($position + $match.index),$space_type.key) #Applies to entire block of text
                                        }
                                        else
                                        {
                                            #write-host space match
                                        }
                                    }
                                    if($size -le 2718)
                                    {
                                        #write-host Compression Success
                                        #write-host $line
                                        $stop = 1;
                                        [int]$size = $size
                                        #write-host New Size: $size
                                        break
                                    }
                                }
                            }#Foreach Match
                        }#Foreach Type
                    }#Matches Success
                    ############################################
                    ##Index Bullet in 3 Bullet tracking hashes
                    $script:bullets_compressed[$line] = $size
                    $script:bullets_and_lines[$line_count] = $line
                    if($script:bullets_and_sizes.contains($original))
                    {
                        $script:bullets_and_sizes.remove($original);
                        if(!($script:bullets_and_sizes.contains($line)))
                        {
                            $script:bullets_and_sizes.add($line,$size);
                        }
                    }             
                }#Already Compressed
                else
                {
                    #write-host Already Compressed
                }
            }#Blank Line
            $position = $position + $line.Length + 1;
        }#Foreach Line
        
        update_sizer_box
    }#Compression ON


    $ghost_editor.ZoomFactor = $editor.ZoomFactor
    $ghost_editor.SelectAll();
    $ghost_editor.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_FONT_COLOR'])
    $ghost_editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
    $ghost_editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])


    $simplified_text = $editor.Text -replace " | | |`n`r|`n|`r",' ' #Treats all spaces the same.
    $simplified_text2 = $simplified_text.ToLower();
    $simplified_text3 = $simplified_text2 -replace "[^a-z]`n`r|`n|`r",' '

    #write-host $simplified_text
    #write-host $simplified_text2
    #write-host $simplified_text3
    $script:acro_index = @{}; 
    foreach($entry in $script:acronym_list.getEnumerator())
    {
        $acro_split = $entry.key -split "::"
        $find = "";
        $find2 = "";
        $run_alternator = 0; #Switches from Short-hand & Long-hand acronym searches 
        foreach($find in $acro_split)
        {
            
            $matches = "";
            $meaning = "";    
            if($run_alternator -eq 0)
            {
                #write-host "Shorthand"
                $meaning = $acro_split[1];
                $find2 = $find.substring(0,1).toupper() + $find.substring(1)
                $pattern = "$([regex]::escape($find))|$([regex]::escape($find2))"
                $matches = [regex]::Matches($simplified_text, $pattern)
            }
            else
            {
                #write-host "Longhand"
                $meaning = $acro_split[0];
                $find2 = $find.ToLower(); #Search via Lower case
                #write-host ------------
                #write-host $find
                #write-host $find2
                $pattern = "$([regex]::escape($find2))"
                $matches = [regex]::Matches($simplified_text2, $pattern)
            }

            if($matches)
            {  
                foreach($match in $matches)
                {
                    $add_to_index = "On";
                    $index = $match.index;
                    ##########Get Characters Before & After Found acronyms
                    $before_acro = "";
                    $after_acro = "";
                    if(!(($match.index - 1) -le 0))
                    {
                        $before_acro = $simplified_text.substring($match.index - 1, 1);
                    }
                    if(!(($match.index + $find.length + 2) -ge $simplified_text.length))
                    {
                        $after_acro  = $simplified_text.substring($match.index + $find.length, 2);
                    }
                    elseif(!(($match.index + $find.length + 1) -ge $simplified_text.length))
                    {
                        $after_acro  = $simplified_text.substring($match.index + $find.length, 1);
                    }


                    ######Determine if find is valid
                    $modified_acro = $find
                    if(!(($before_acro -match " |/|-|\.|`r`n") -or ($before_acro -eq "")))
                    {
                        $add_to_index = "Off 1"
                    
                    } 
                    if(!(($after_acro -match "^ |^; |^--|^/|^\.\.|!|`n|^-") -or ($after_acro -eq "") -or ($find -match "^w/$|^f/$"))) #No Character After
                    {
                        if(($after_acro -cmatch "^s |^d |^s/|^d/|^s-|^d-|^s\.|^d\.|^s`n|^`d") -or ($after_acro -eq "")) #One Character After
                        {
                            $modified_acro = $modified_acro + $after_acro.substring(0,1);
                        }
                        elseif(($after_acro -cmatch "^'s|^'d|^'g") -or ($after_acro -eq "")) #Two Characters After
                        {
                            $modified_acro = $modified_acro + $after_acro.substring(0,2);
                        }
                        else
                        {
                            $add_to_index = "Off 2 $find = -$after_acro-"
                        }
                    }
                
                    ##############Add To Index                         
                    if($add_to_index -eq "On")
                    {   
                        $key = "0"
                        if($run_alternator -eq 0)
                        {
                            $key = "E" + "::" + $index + "::" + $modified_acro + "::" + $meaning #Shorten
                        }
                        else
                        {
                            $key = "S" + "::" + $index + "::" + $modified_acro + "::" + $meaning #Extend
                        }
                        #write-host $run_alternator $key
                        if(!($script:acro_index.containskey($key)))
                        {
                            $script:acro_index.Add($key,$index);
                        }      
                    }           
                }#Each Match
            }#Found matches
            $run_alternator++;
        }#Foreach Acrnym side (Shorthand/Longhand)
    }#Foreach Acronym Set
    ##################################################################

    $pattern = "\w+"
    $matches = [regex]::Matches("$simplified_text3", $pattern)

    if($matches)
    {  
        foreach($match in $matches)
        {
            $word = $match.value       
            $found = 0;
            $status = "C";
            ########Check Acronyms
            foreach($entry in $script:acro_index.getEnumerator()) #Check to make sure it's not an Acronym 
            {
                if($entry.value -eq $match.index)
                {
                    $found = 1;
                    break;
                }
            }
            #######Check if Number
            if($found -eq 0)
            {
                if($word -match "^\d+")
                {
                    $word = $editor.text.substring($match.index,$match.value.length) #Remove the lowercase
                    #write-host $word
                    if($word -cmatch "\d$|K$|M$|B$|T$|x$")
                    {
                        $found = 1;
                    }
                    elseif($word -cmatch "k$|m$|b$|t$")
                    {
                        $correct = $word.ToUpper();
                        if(!($script:dictionary_index.contains($word)))
                        {
                            $script:dictionary_index.Add($word,$correct);
                        }
                        #write-host "$word = $correct"
                    }
                }
            }
            #######Check Dictionary Index
            if($found -eq 0)
            {
                    
                if($script:dictionary_index.containskey($match.value))
                {
                        $found = 1;
                        if($script:dictionary_index[$match.value] -ne "C")
                        {
                            $status = "M" #Misspelled
                        } 
                } 
            }
            ######Check Dictionary
            if($word.length -eq 1)
            {
                if(($editor.text.substring($match.index,1) -cmatch "a|I") -or (($match.index -ge 1) -and ($editor.text.substring(($match.index -1),1) -cmatch "'")))
                {
                    $found = 1;
                    $status = "C"
                }
            }
            if(($found -eq 0) -and ($word.length -ge 2))
            {
                $word_find1 = $match.value;
                $word_find2 = $match.value + "s";
                $word_find3 = $match.value + "d";
                $word_find4 = $match.value + "ing";
                $word_find5 = $match.value -replace "s$|ed$",''
                $word_find6 = $match.value.substring(0,($match.value.length - 1))
                if($script:dictionary.containskey($word_find1))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                elseif($script:dictionary.containskey($word_find2))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                elseif($script:dictionary.containskey($word_find3))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                elseif($script:dictionary.containskey($word_find4))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                elseif($script:dictionary.containskey($word_find5))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                elseif($script:dictionary.containskey($word_find6))
                {
                    $found = 1;
                    $script:dictionary_index.Add($match.value,"C");
                }
                    
            }
            ######Not Found in any list (Word is Misspelled) 
            if($found -eq 0)
            {
                $status = "M" #Misspelled
                $script:dictionary_index.Add($match.value,"M");
            }
            if($status -eq "M")
            {
                $ghost_editor.SelectionStart = $match.index
                $ghost_editor.SelectionLength = $match.value.length
                $ghost_editor.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'])
                #$ghost_editor.SelectionFont = [Drawing.Font]::New('Times New Roman', 14)
                $ghost_editor.SelectionFont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
                $ghost_editor.DeselectAll()
            }     
            
        }#Foreach Matches
    }#If Matches
    ###################################################################
    foreach($entry in $script:acro_index.getEnumerator() | sort key)
    {
            ($mode,$index,$acronym,$meaning) = $entry.key -split "::"

            $ghost_editor.SelectionStart = $index
            $ghost_editor.SelectionLength = $acronym.length
            if($mode -eq "E")
            {
                $ghost_editor.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'])
            }
            else
            {
                $ghost_editor.SelectionColor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'])
            }
            #$ghost_editor.SelectionFont = [Drawing.Font]::New('Times New Roman', 14)
            $ghost_editor.SelectionFont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            $ghost_editor.DeselectAll()
    }
    ##############################################################################
    #Translate Half Spaces (This provides support for fonts that don't have halfspaces)

    $pattern = " | | "
    $matches = [regex]::Matches($ghost_editor.text, $pattern)

    if($matches)
    {  
        foreach($match in $matches)
        {
            #write-host ---------------------------------
            #write-host MV $match.value.length
            #write-host MI $match.index
            $ghost_editor.SelectionStart = $match.index
            $ghost_editor.SelectionLength = $match.value.length
            $ghost_editor.SelectionFont = [Drawing.Font]::New('Times New Roman', 14)
            $ghost_editor.SelectionFont = [Drawing.Font]::New('Times New Roman', [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            $ghost_editor.DeselectAll()
        }
    }
}
################################################################################
######Clipboard Functions#######################################################
function clipboard_cut
{
    [string]$text = $editor.SelectedText
    $text = $text.replace("`n","`r");
    if($text)
    {
        Set-Clipboard -Value $text
        $editor.SelectedText = ""
    }
}
function clipboard_copy
{
    [string]$text = $editor.SelectedText
    $text = $text.replace("`n","`r");
    if($text)
    {
        Set-Clipboard -Value $text
    }
}
function clipboard_paste
{
   $paste = Get-Clipboard -raw
   $editor.SelectedText = $paste
}
function clipboard_copy_2
{
    #Clipboard Menu for Feeder Only
    [string]$text = $feeder_box.SelectedText
    $text = $text.replace("`n","`r");
    if($text)
    {
        Set-Clipboard -Value $text
    }
}
function clipboard_copy_3
{
    #Clipboard Menu for Acronym Lists
    [string]$text = $acronym_box.SelectedText
    $text = $text.replace("`n","`r");
    if($text)
    {
        Set-Clipboard -Value $text
    }
}
################################################################################
Function flush_memory
{
    if($script:settings['MEMORY_FLUSHING'] -match "3|4")
    {
        $Script:Timer.Stop()
        Log "Flushing Memory Start"
        $temp = "SUBLOG     Flush mode " + $script:settings['MEMORY_FLUSHING']
        Log "$temp"
        Remove-Variable "temp"

        #$memory =  [System.Math]::Round((((Get-Process -Id $PID).PrivateMemorySize))/1mb, 1)
        #$memory = $script:memBefore = (Get-Process -id $PID | Sort-Object WorkingSet64 | Select-Object Name,@{Name='WorkingSet';Expression={($_.WorkingSet64/1KB)}})
        #$memory = [System.Math]::Round(($script:memBefore.WorkingSet)/1mb, 2)

        $user_vars = @();
        $user_vars = Get-Variable | Select-Object -ExpandProperty Name
        $user_vars = (Compare-Object -ReferenceObject $script:main_vars -DifferenceObject $user_vars)
        $flush_count = 0;
        foreach($item in $user_vars.GetEnumerator() | sort )
        {
        
            [string]$string = $item.InputObject
            if(!($string -match "^memory$|^PSItem$|^this$|^_$|^AboutMenu$|^AcronymMenu$|^BulletMenu$|^EditMenu$|^FileMenu$|^OptionsMenu$"))
            {
                if($item.InputObject -ne "")
                {
                    #write-host "     ", $item.InputObject
                    $flush_count++
                    Remove-Variable $string -ErrorAction SilentlyContinue
                }
            }
        }
        Log "SUBLOG     $flush_count Removed"
        Log "Flushing Memory End"
        Log "BLANK"
    }
    if($script:settings['MEMORY_FLUSHING'] -match "2|4")
    {
        $Script:Timer.Stop()
        Log "Garbage Collect Started"
        [System.GC]::GetTotalMemory(‘forcefullcollection’) | out-null
        Log "Garbage Collect Ended"
        Log "BLANK"
    }
    $Script:Timer.Start()
}
################################################################################
######Idle Timer################################################################
Function Idle_Timer
{
    #########################################################################################
    ##Track Ticks
    --$Script:CountDown

    #########################################################################################
    ##Fix Menu Color Glitch
    ##File Menu
    if($FileMenu.Pressed)
    {
        $FileMenu.Forecolor = "Black"
    }
    else
    {
        if($FileMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }
    ##Edit Menu
    if($EditMenu.Pressed)
    {
        $EditMenu.Forecolor = "Black"
    }
    else
    {
        if($EditMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            $EditMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }
    ###Bullet Menu
    if($BulletMenu.Pressed)
    {
        $BulletMenu.Forecolor = "Black"
    }
    else
    {
        if($BulletMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }
    ##Acronym Menu
    if($script:AcronymMenu.Pressed)
    {
        $script:AcronymMenu.Forecolor = "Black"
    }
    else
    {
        if($script:AcronymMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            $script:AcronymMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }
    ##Options Menu
    if($OptionsMenu.Pressed)
    {
        $OptionsMenu.Forecolor = "Black"
    }
    else
    {
        if($OptionsMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            
            $OptionsMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }
    ##About Menu
    if($AboutMenu.Pressed)
    {
        $AboutMenu.Forecolor = "Black"
    }
    else
    {
        if($AboutMenu.Forecolor -eq "Black") #Reduce GUI Changes
        {
            $AboutMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
        }
    }

    #########################################################################################
    ##Force Reload a function
    if($script:reload_function -ne "")
    {
        Log "Running: Reload Function"
        #write-host RELOAD FUNCTION $script:reload_function $script:reload_function_arg_a 
        $buffer = $script:reload_function #Prevent timer from executing too many.
        $arg_a = $script:reload_function_arg_a
        $arg_b = $script:reload_function_arg_b
        $script:reload_function = "";
        $script:reload_function_arg_a = "";
        $script:reload_function_arg_b = "";         
        &$buffer $arg_a $arg_b
        Log "Ended: Reload Function"
                            
    }
    #########################################################################################
    ##Check Thesaurus Jobs
    if(($script:thesaurus_job -ne "") -and ($script:thesaurus_job.state))
    {
        Log "Running: Thesaurus_job"
        thesaurus_lookup
        Log "Ended: Thesaurus_job"
    }
    #########################################################################################
    ##Check Word Hippo Jobs
    if(($script:word_hippo_job -ne "") -and ($script:word_hippo_job.state))
    {
        Log "Running: Word_hippo_job"
        word_hippo_lookup
        Log "Ended: Word_hippo_job"
    }
    #########################################################################################
    ##User Resizing Entire Form
    if(($script:Form_height -ne $script:Form.height) -and ($Form.WindowState -ne "Minimized"))
    {
        Log "Resizing Form (Height) Start"

        $ratio = (($editor.height * $script:Form.height) / $script:Form_height)
        $script:Form_height = $script:Form.height

        $overlock1 = (($script:Form.height - 100) - ($ratio))
        if($overlock1 -ge 33)
        {
            $editor.height = $ratio
        }
        else
        {
            $editor.height = ($script:Form.height - 133)
        }
        $sizer_box.height = ($editor.height - 4)
        $bullet_feeder_panel.Location = New-Object System.Drawing.Size(($editor.Location.x),($editor.location.Y + $editor.height))
        $bullet_feeder_panel.height = ($script:Form.height - $editor.height - 100)
        $feeder_box.height = ($bullet_feeder_panel.height - 5)
        $sidekick_panel.height = ($script:Form.Height - 100)
        $script:zoom = "Changed"
        Log "Resizing Form (Height) End"
        Log "BLANK"

    }
    if($script:Form_Width -ne $script:Form.Width)
    {
        Log "Resizing Form (Width) Start"
        $find_best_zoom = $editor.ZoomFactor
        if($script:Form_Width -lt $script:Form.Width)
        {
            #write-host "Grew"
            $zoom_incr = $editor.ZoomFactor
            While($zoom_incr -lt 3)
            { 
                $zoom_incr = $zoom_incr + 0.01
                #write-host ZI $zoom_incr
                if($sidekick_panel.width -ne 5)
                {
                    #write-host "EXISTS"
                    $contraints = (($editor.location.x) + ($zoom_incr * 1200) + ($sizer_box.width) + $script:Sidekick_width + 40)
                }
                else
                {
                    #write-host "DOESN'T"
                    $contraints = (($editor.location.x) + ($zoom_incr * 1200) + ($sizer_box.width) + 40)
                }                 
                if($contraints -le $script:Form.Width)
                {
                    $find_best_zoom = $zoom_incr
                    #break
                }
                   
            }
        }
        else
        {
            #write-host "Shrunk"
            $zoom_incr = $editor.ZoomFactor
            While($zoom_incr -ge 0)
            { 
                if($sidekick_panel.width -ne 5)
                {
                    $contraints = (($editor.location.x) + ($zoom_incr * 1200) + $sizer_box.width + $script:Sidekick_width + 40)
                }
                else
                {
                    $contraints = (($editor.location.x) + ($zoom_incr * 1200) + $sizer_box.width + 40)
                }  
                
                if($contraints -lt $script:Form.Width)
                {
                    $find_best_zoom = $zoom_incr
                    break
                }
                $zoom_incr = $zoom_incr - 0.01 
            }
        }
        #############Found Best Zoom
        $editor.ZoomFactor = $find_best_zoom
        $script:zoom = "Changed"
        $script:Form_width = $script:Form.Width
        Log "Resizing Form (Width) End"
        Log "BLANK"
    }
    
    #########################################################################################
    ##User Resizing Editor/Feeder
    if($script:user_resizing -ne 0)
    {
        Log "Resizing Editor/Feeder Start"
        $Script:Timer.Interval = 1;
        $distance = ($script:user_resizing - ([System.Windows.Forms.Cursor]::Position.Y))

        $overlock1 = (($script:Form.height - 100) - ($script:user_resizing_starting_height - $distance)) #Prevents Feeder Too Small
        $overlock2 = ($script:user_resizing_starting_height - $distance)                          #Prevents Editor Too Small

        if(($overlock1 -gt 33) -and ($overlock2 -gt 33))
        {
            $editor.height = $script:user_resizing_starting_height - $distance
            $sizer_box.height = ($editor.height - 4)
            $bullet_feeder_panel.Location = New-Object System.Drawing.Size(($editor.Location.x),($editor.location.Y + $editor.height))
            $bullet_feeder_panel.height = ($script:Form.height - $editor.height - 100)
            $feeder_box.height = ($bullet_feeder_panel.height - 5)
            $sizer_art.height = $feeder_box.height + 5
            $sizer_art.Location = New-Object System.Drawing.Size($sizer_box.location.x,($sizer_box.Location.y + $sizer_box.height)) 
        }
        Log "Resizing Editor/Feeder End"
        Log "BLANK"
    }
    #########################################################################################
    ###Zoom Changes
    if($Script:LockInterface -ne 1)
    {   
        
        if(($script:zoom -ne $editor.ZoomFactor) -or ($script:zoom -ne $feeder_box.ZoomFactor))
        {
            Log "Zoom Change Start"
            #write-host Zoom Changed $editor.zoomfactor = $script:zoom
            if($script:zoom -ne $editor.ZoomFactor)
            {
                $script:zoom = $editor.ZoomFactor
                $feeder_box.ZoomFactor = $script:zoom 
                #write-host "Editor Zoom"
            }
            else
            {
                $script:zoom = $feeder_box.ZoomFactor
                $editor.ZoomFactor = $script:zoom
            }

            $editor.Width = ($script:zoom) * 1200
            $bullet_feeder_panel.Width = $editor.Width
            $feeder_box.Width = $editor.Width
            #$script:Form.Width = $editor.Width + 30


            $sizer_box.Width = ($script:zoom + .1) * $script:sizer_box_width
            $sizer_box.ZoomFactor = $script:zoom 
            $sizer_box.Location = New-Object System.Drawing.Size(($editor.Location.x + $editor.width),($editor.Location.y + 3))
            
            
            
            if($sidekick_panel.width -ne 5)
            {
                $sidekick_panel.Location = New-Object System.Drawing.Point(($sizer_box.Location.x + $sizer_box.width),$editor.Location.y)
                $size = ($script:Form.width - ($sidekick_panel.Location.x + 25))
                $sidekick_panel.width = $size
                $left_panel.height                   = ($sidekick_panel.height)
                $left_panel.width                    = ($sidekick_panel.width - 5)
            }
            else
            {
                #$sidekick_panel.Location = New-Object System.Drawing.Point(($script:Form.width - 30),$editor.Location.y)
                $sidekick_panel.Location = New-Object System.Drawing.Point(($sizer_box.Location.x + $sizer_box.width),$editor.Location.y)
                $left_panel.height                   = ($sidekick_panel.height)
                $left_panel.width                    = ($sidekick_panel.width - 5)
            }

            #$script:Form.width =  ($sidekick_panel.Location.x + $sidekick_panel.width + 20)
            $script:Form.MinimumSize = New-Object Drawing.Size((1200 + $sizer_box.width + 35),200)
            #$script:Form.MaximumSize = New-Object Drawing.Size(($sidekick_panel.Location.x + $script:Sidekick_width + 35),5000)


            $sizer_art.Width = $sizer_box.width
            $sizer_art.height = $feeder_box.height + 5
            $sizer_art.Location = New-Object System.Drawing.Size($sizer_box.location.x,($sizer_box.Location.y + $sizer_box.height))
            
            $script:sidekickgui = "New"
            #sidekick_display
            Log "Zoom Change End"
            Log "BLANK"
        }
    }
    
    #########################################################################################
    ###Update Bullet Feed/Text Size Detection
    if((([string]$script:feeder_job -ne "") -and ([string]$script:feeder_job.state)) -or ($editor.selectionstart -ne $script:caret_position) -or ($editor.Text -cne "$Script:recent_editor_text"))
    {
        Log "Update Feeder Start"
        #write-host Feeder Update $script:caret_position = $editor.SelectionStart = $script:feeder_job.state
        
        ##Update Line Location
        if(((Test-path variable:script:location_value) -and ($script:caret_position -ne $editor.SelectionStart)) -or ($editor.Text -cne "$Script:recent_editor_text"))
        {
            #write-host Finding
            $lines = $editor.text -split "`n";
            $size = 0;
            $position = 0;
            $counter = 0
            foreach($line in $lines)    
            {   
                $counter++;
                #########################################################################################
                ###Text Size Detection
                if(!($script:bullets_and_sizes.Contains($line)))
                {
                    $size = calculate_text_size_new $line
                    #write-host Added Bullet: $line = ($size)
                    $script:bullets_and_sizes.add($line,$size);
                }
                if($script:bullets_and_lines.contains($counter))
                {
                    if($script:bullets_and_lines[$counter] -ne $line)
                    {
                        $script:bullets_and_lines[$counter] = $line
                        #write-host Updated Bullet# $counter  = $script:bullets_and_lines[$counter] 
                    }
                }
                else
                {
                    $script:bullets_and_lines.add($counter,"$line")
                    #write-host Added Bullet# $counter  = $script:bullets_and_lines[$counter] 
                }

                #########################################################################################
                ###Update Feeder & Sidekick Line Location
                
                if($editor.selectionstart -le ($position + $line.length) -and ($editor.selectionstart -ge $position))
                {
                    if($current_bullet -ne $line)
                    {
                        update_feeder #Only Run if bullet user is on changed or moved to a different line
                        $script:current_bullet = $line
                    }
                    
                    $script:current_line = $counter
                    if(($script:location_value -ne "") -and ($script:location_value.text -ne $counter))
                    {
                        #Update Sidekick "Current Line"
                        $script:location_value.text = $counter
                    }
                    #break;
                }
                $position = $position + $line.length + 1;
            }
            update_sizer_box
        }   
        $script:caret_position = $editor.SelectionStart
        if($script:feeder_job -ne "")
        {
            #check status of job
            update_feeder
        }
        Log "Update Feeder End"
        Log "BLANK"
    }

    #########################################################################################
    ###Update Sidekick
    if($script:sidekick_job -ne "")
    {
        Log "Update Sidekick Start"
        #write-host Update sidekick
        update_sidekick
        Log "Update Sidekick End"
        Log "BLANK"
    }

    #########################################################################################
    ###Save History & Run Memory Cleanup
    if((((Get-Date) - $script:save_history_timer).TotalSeconds -gt 10) -or (($script:save_history_job -ne "") -and ([string]$script:save_history_job.state -eq "Completed")))
    {
        Log "Save History Start"
        $script:save_history_timer = Get-Date
        save_history
        Log "Save History End"
        Log "BLANK"

        flush_memory
        var_sizes
    }

    #########################################################################################
    ###Track Text Changes
    if ($Script:CountDown -eq 0)
    {   
        if($editor.Text -cne "$Script:recent_editor_text")
        {    
            Log "Scan Text Start"   
            #write-host "Text Changed"
            $Script:LockInterface = 1;
            $script:caret_position = $editor.SelectionStart;
            $start = $editor.SelectionStart
            $length = $editor.SelectionLength
            [string]$Script:recent_editor_text = $editor.Text

            scan_text

            $editor.rtf = $ghost_editor.Rtf
            $editor.SelectionStart = $script:caret_position
            $editor.SelectionStart = $start;
            $editor.SelectionLength = $length;
            
            while($editor.ZoomFactor -ne $script:zoom) 
            {
                $editor.SelectionStart = $script:caret_position

                #Zoom Changes during RTF replace, but won't change in time... this is a work around.
                $editor.ZoomFactor = $script:zoom
                $ghost_editor.ZoomFactor = $editor.ZoomFactor              
            }
            update_sidekick

            #Force Scroll to Sync
            $editor.CustomVScroll()
            $editor.CustomVScroll()
            $editor.CustomVScroll()
            $editor.CustomVScroll()
            $editor.CustomVScroll()
            $editor.CustomVScroll()

            $Script:LockInterface = 0;
            Log "Scan Text End"
            Log "BLANK"
        }
        $Script:CountDown = 1 
    }
 }
################################################################################
######Initial Checks############################################################
function initial_checks
{
    if(!(test-path -LiteralPath "$dir\Resources"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Required"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Required"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Required\Processing"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Required\Processing"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Bullet Banks"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Bullet Banks"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Acronym Lists"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Acronym Lists"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Packages"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Packages"
    }
    if(!(test-path -LiteralPath "$dir\Resources\Themes"))
    {
        New-Item  -ItemType directory -Path "$dir\Resources\Themes"
    }
    if($PSVersionTable.PSVersion.Major -lt 5)
    {
        $version = $PSVersionTable.PSVersion
        write-host
        write-host
        write-host "WARNING: You are running Powershell version $version. This program is designed to run on Powershell version 5 or greater. Some portions of this script may not function properly."
        write-host "You can download the latest Powershell version from Microsoft here."
        write-host
        write-host "https://www.microsoft.com/en-us/download/details.aspx?id=54616"
        write-host
        write-host "(i.e. Windows Management Framework)"
        write-host
        write-host
    }
    ###############No Dictionary Error
    if(!(test-path -LiteralPath "$dir\Resources\Required\Dictionary.txt"))
    {             
        $error_form = New-Object System.Windows.Forms.Form
        $error_form.FormBorderStyle = 'Fixed3D'
        $error_form.BackColor             = "Black"
        $error_form.Location = new-object System.Drawing.Point(0, 0)
        $error_form.MaximizeBox = $false
        $error_form.SizeGripStyle = "Hide"
        $error_form.Width = 750

        $error_form.Height = 250

        $y_pos = 10;

        $title_label                          = New-Object system.Windows.Forms.Label
        $title_label.text                     = "ERROR: No Dictionary"
        $title_label.ForeColor                = "RED"
        $title_label.Anchor                   = 'top,right'
        $title_label.width                    = ($error_form.width)
        $title_label.height                   = 30
        $title_label.TextAlign = "MiddleCenter"
        $title_label.Font                     = [Drawing.Font]::New("Times New Roman", 17)
        $title_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
        $error_form.controls.Add($title_label);

        $y_pos = 50;

        $message_one                          = New-Object system.Windows.Forms.Label
        $message_one.text                     = "You can download a Dictionary here:"
        $message_one.ForeColor                = "White"
        $message_one.Anchor                   = 'top,right'
        #$message_one.autosize                 = $true
        $message_one.TextAlign                = "Middlecenter"
        $message_one.width                    = $error_form.Width
        $message_one.height                   = 30
        $message_one.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_one.width / 2)),$y_pos);
        $message_one.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $error_form.controls.Add($message_one);

        $y_pos = $y_pos + 40;

        $message_two                          = New-Object system.Windows.Forms.Label
        $message_two.text                     = "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required`n(Click Here)"
        $message_two.ForeColor                = "yellow"
        $message_two.Anchor                   = 'top,right'
        $message_two.autosize                 = $true
        $message_two.TextAlign                = "Middlecenter"
        $message_two.width                    = $error_form.Width
        $message_two.height                   = 60
        $message_two.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_two.width / 2)),$y_pos);
        $message_two.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $message_two.add_click({
            Start-Process "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required"
        })
        $error_form.controls.Add($message_two);
        

        $y_pos = $y_pos + 45;

        $message_three                          = New-Object system.Windows.Forms.Label
        $message_three.text                     = "After download, put file in local directory: \Resources\Required"
        $message_three.ForeColor                = "white"
        $message_three.Anchor                   = 'top,right'
        #$message_three.autosize                 = $true
        $message_three.TextAlign                = "Middlecenter"
        $message_three.width                    = $error_form.Width
        $message_three.height                   = 60
        $message_three.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_three.width / 2)),$y_pos);
        $message_three.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $message_three.add_click({
            Start-Process "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required"
        })
        $error_form.controls.Add($message_three);
        $error_form.ShowDialog();

        exit;
    }

    if(!(test-path -LiteralPath "$dir\Resources\Required\Thesaurus.csv"))
    {
              
        $error_form = New-Object System.Windows.Forms.Form
        $error_form.FormBorderStyle = 'Fixed3D'
        $error_form.BackColor             = "Black"
        $error_form.Location = new-object System.Drawing.Point(0, 0)
        $error_form.MaximizeBox = $false
        $error_form.SizeGripStyle = "Hide"
        $error_form.Width = 750

        $error_form.Height = 250

        $y_pos = 10;

        $title_label                          = New-Object system.Windows.Forms.Label
        $title_label.text                     = "ERROR: No Thesaurus"
        $title_label.ForeColor                = "RED"
        $title_label.Anchor                   = 'top,right'
        $title_label.width                    = ($error_form.width)
        $title_label.height                   = 30
        $title_label.TextAlign = "MiddleCenter"
        $title_label.Font                     = [Drawing.Font]::New("Times New Roman", 17)
        $title_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
        $error_form.controls.Add($title_label);

        $y_pos = 50;

        $message_one                          = New-Object system.Windows.Forms.Label
        $message_one.text                     = "You can download a Thesaurus here:"
        $message_one.ForeColor                = "White"
        $message_one.Anchor                   = 'top,right'
        #$message_one.autosize                 = $true
        $message_one.TextAlign                = "Middlecenter"
        $message_one.width                    = $error_form.Width
        $message_one.height                   = 30
        $message_one.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_one.width / 2)),$y_pos);
        $message_one.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $error_form.controls.Add($message_one);

        $y_pos = $y_pos + 40;

        $message_two                          = New-Object system.Windows.Forms.Label
        $message_two.text                     = "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required`n(Click Here)"
        $message_two.ForeColor                = "yellow"
        $message_two.Anchor                   = 'top,right'
        $message_two.autosize                 = $true
        $message_two.TextAlign                = "Middlecenter"
        $message_two.width                    = $error_form.Width
        $message_two.height                   = 60
        $message_two.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_two.width / 2)),$y_pos);
        $message_two.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $message_two.add_click({
            Start-Process "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required"
        })
        $error_form.controls.Add($message_two);
        

        $y_pos = $y_pos + 45;

        $message_three                          = New-Object system.Windows.Forms.Label
        $message_three.text                     = "After download, put file in local directory: \Resources\Required"
        $message_three.ForeColor                = "white"
        $message_three.Anchor                   = 'top,right'
        #$message_three.autosize                 = $true
        $message_three.TextAlign                = "Middlecenter"
        $message_three.width                    = $error_form.Width
        $message_three.height                   = 60
        $message_three.location                 = New-Object System.Drawing.Point((($error_form.width / 2) - ($message_three.width / 2)),$y_pos);
        $message_three.Font                     = [Drawing.Font]::New("Times New Roman", 12)
        $message_three.add_click({
            Start-Process "https://github.com/Jukari2003/Bullet-Blender/tree/main/Resources/Required"
        })
        $error_form.controls.Add($message_three);
        $error_form.ShowDialog();

        exit;
    }
    ###########################################################################################
    ###Bullet Banks
    if(!(Test-Path -LiteralPath "$dir\Resources\Required\Bullet_lists.txt"))
    {
        $bullet_bank_list = Get-ChildItem "$dir\Resources\Bullet Banks"  -File -filter *.txt
        foreach($bank in $bullet_bank_list)
        {
            $line = $bank.Name + "::1" 
            Add-Content "$dir\Resources\Required\Bullet_lists.txt" $line
            $Script:bullet_banks.Add($bank.Name,1);
        }
    }
    else
    {
        ##Get Current Banks Status
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Bullet_lists.txt")
        while($null -ne ($line = $reader.ReadLine()))
        {
            ($file,$status) = $line -split "::";
            if(!($Script:bullet_banks.Contains($file)) -and (Test-Path -LiteralPath "$dir\Resources\Bullet Banks\$file"))
            {
                $Script:bullet_banks.Add($file,$status);
                Add-Content "$dir\Resources\Required\Bullet_lists_temp.txt" $line
            }
        }
        $reader.close();      
        ##Verify Bank Integrity
        $bullet_bank_list = Get-ChildItem "$dir\Resources\Bullet Banks"  -File -filter *.txt
        foreach($bank in $bullet_bank_list)
        {
            if(!($Script:bullet_banks.contains($bank.name)))
            {
                $line = $bank.Name + "::1" 
                $Script:bullet_banks.Add($bank.name,1);
                Add-Content "$dir\Resources\Required\Bullet_lists_temp.txt" $line
            } 
        }
        if(Test-Path -LiteralPath "$dir\Resources\Required\Bullet_lists_temp.txt")
        {
            if(Test-Path -LiteralPath "$dir\Resources\Required\Bullet_lists.txt")
            {
                Remove-Item -LiteralPath "$dir\Resources\Required\Bullet_lists.txt"
            }
            Rename-Item -LiteralPath "$dir\Resources\Required\Bullet_lists_temp.txt" "$dir\Resources\Required\Bullet_lists.txt"
        }
    }
    ###################################################################################
    ###Acronym Lists
    if(!(Test-Path -LiteralPath "$dir\Resources\Required\Acronym_lists.txt"))
    {
        $acronym_list = Get-ChildItem "$dir\Resources\Acronym Lists"  -File -filter *.csv
        foreach($list in $acronym_list)
        {
            $line = $list.Name + "::1" 
            Add-Content "$dir\Resources\Required\Acronym_lists.txt" $line
            $Script:acronym_lists.Add($list.Name,1);
        }
    }
    else
    {
        ##Get Current Banks Status
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Acronym_lists.txt")
        while($null -ne ($line = $reader.ReadLine()))
        {
            ($file,$status) = $line -split "::";
            if(!($Script:acronym_lists.Contains($file)) -and (Test-Path -LiteralPath "$dir\Resources\Acronym Lists\$file"))
            {
                $Script:acronym_lists.Add($file,$status);
                Add-Content "$dir\Resources\Required\Acronym_lists_temp.txt" $line
            }
        }
        $reader.close();      
        ##Verify Bank Integrity
        $acronym_list = Get-ChildItem "$dir\Resources\Acronym Lists"  -File -filter *.csv
        foreach($list in $acronym_list)
        {
            if(!($Script:acronym_lists.contains($list.name)))
            {
                $line = $list.Name + "::1" 
                $Script:acronym_lists.Add($list.name,1);
                Add-Content "$dir\Resources\Required\Acronym_lists_temp.txt" $line
            } 
        }
        if(Test-Path -LiteralPath "$dir\Resources\Required\Acronym_lists_temp.txt")
        {
            if(Test-Path -LiteralPath "$dir\Resources\Required\Acronym_lists.txt")
            {
                Remove-Item -LiteralPath "$dir\Resources\Required\Acronym_lists.txt"
            }
            Rename-Item -LiteralPath "$dir\Resources\Required\Acronym_lists_temp.txt" "$dir\Resources\Required\Acronym_lists.txt"
        }
    }
    ###################################################################################
    ###Package List
    if(!(Test-Path -LiteralPath "$dir\Resources\Required\Package_list.txt"))
    {
        $package_lists = Get-ChildItem "$dir\Resources\Packages"  -Directory
        foreach($package in $package_lists)
        {
            $line = $package.Name + "::1" 
            Add-Content "$dir\Resources\Required\Package_list.txt" $line
            $Script:package_list.Add($package.Name,1);
        }
    }
    else
    {
        ##Get Current Package Status
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Package_list.txt")
        while($null -ne ($line = $reader.ReadLine()))
        {
            ($file,$status) = $line -split "::";
            if(!($Script:package_list.Contains($file)) -and (Test-Path -LiteralPath "$dir\Resources\Packages\$file"))
            {
                $Script:package_list.Add($file,$status);
                Add-Content "$dir\Resources\Required\Package_list_temp.txt" $line
            }
        }
        $reader.close();      
        ##Verify Bank Integrity
        $package_lists = Get-ChildItem "$dir\Resources\Packages"  -Directory
        foreach($package in $package_lists)
        {
            if(!($Script:package_list.contains($package.name)))
            {
                $line = $package.Name + "::1" 
                $Script:package_list.Add($package.name,1);
                Add-Content "$dir\Resources\Required\Package_list_temp.txt" $line
            } 
        }
        if(Test-Path -LiteralPath "$dir\Resources\Required\Package_list_temp.txt")
        {
            if(Test-Path -LiteralPath "$dir\Resources\Required\Package_list.txt")
            {
                Remove-Item -LiteralPath "$dir\Resources\Required\Package_list.txt"
            }
            Rename-Item -LiteralPath "$dir\Resources\Required\Package_list_temp.txt" "$dir\Resources\Required\Package_list.txt"
        }
    }

    ###################################################################################
    if(!(test-path -LiteralPath "$dir\Resources\Required\Settings.csv"))
    {
        $settings_writer = new-object system.IO.StreamWriter("$dir\Resources\Required\Settings.csv",$true)
        $settings_writer.write("PROPERTY,VALUE`r`n");
        $settings_writer.write("THEME,Dark Castle`r`n");
        $settings_writer.write("PACKAGE,Current`r`n");
        $settings_writer.write("LOAD_PACKAGES_AS_BULLETS,1`r`n");
        $settings_writer.write("TEXT_COMPRESSION,4`r`n");
        $settings_writer.write("CLOCK_SPEED,500`r`n");
        $settings_writer.write("SAVE_HISTORY_THRESHOLD,200`r`n");
        $settings_writer.write("SIZER_BOX_INVERTED,1`r`n");
        $settings_writer.write("MEMORY_FLUSHING,3`r`n");
        $settings_writer.close();
    }

    ##################################################################################
    #Build Default Themes
    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Blue Falcon.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Blue Falcon.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n");
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,Gray`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,#c0c0c0`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT,Lucida Sans Unicode`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,Gray`r`n")
        $theme_writer.write("DIALOG_FONT_SIZE,14.5`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,Blue`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,LightSteelBlue`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,Blue`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,White`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,#00e1e1`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,Blue`r`n")
        $theme_writer.write("EDITOR_FONT,Times New Roman`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,13`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,Red`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,Lime`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,Yellow`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,#00e1e1`r`n")
        $theme_writer.write("FEEDER_FONT,Times New Roman`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,13`r`n")
        $theme_writer.write("INTERFACE_FONT,Calibri`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,12`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,Blue`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,LightSteelBlue`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,White`r`n")
        $theme_writer.close();
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Dark Castle.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Dark Castle.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n")
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,#585858`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,#434343`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,#c0c0c0`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT,Lucida Sans Unicode`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,#c4c4c4`r`n")
        $theme_writer.write("DIALOG_FONT_SIZE,14.5`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,White`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,White`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,#434343`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,#c4c4c4`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,White`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,Blue`r`n")
        $theme_writer.write("EDITOR_FONT,Times New Roman`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,13`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,Yellow`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,Red`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,Lime`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,White`r`n")
        $theme_writer.write("FEEDER_FONT,Times New Roman`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,13`r`n")
        $theme_writer.write("INTERFACE_FONT,Copperplate Gothic Bold`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,8`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,#676767`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,#c4c4c4`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,#282828`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,#434343`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,#c4c4c4`r`n")
        $theme_writer.close();
    }
    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Alien.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Alien.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n")
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,Silver`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,Olive`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT,Lucida Sans Unicode`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,#808040`r`n")
        $theme_writer.write("DIALOG_FONT_SIZE,14.5`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,Olive`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,Lime`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,#00a200`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,#9d9dff`r`n")
        $theme_writer.write("EDITOR_FONT,Segoe UI`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,#00b700`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,12`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,#8080ff`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,Red`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,#8080ff`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT,Segoe UI`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,12`r`n")
        $theme_writer.write("INTERFACE_FONT,Segoe UI`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,11`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,Lime`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,#80ff00`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,#122400`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,White`r`n")
        $theme_writer.close();
    }

    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Aquarius.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Aquarius.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n")
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,Silver`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,#0079bf`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,#777586`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,#44d1ac`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,White`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,#474555`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,#ffffff`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,#0079bf`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,#aba9bb`r`n")
        $theme_writer.write("EDITOR_FONT,Calibri`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,#c8fcea`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,12`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,#009d00`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,#820015`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,#80ff00`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,#0079bf`r`n")
        $theme_writer.write("FEEDER_FONT,Calibri`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,#c8fcea`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,11`r`n")
        $theme_writer.write("INTERFACE_FONT,Calibri`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,11`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,#002164`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,White`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,#0079bf`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,#002164`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,#c8fcea`r`n")
        $theme_writer.close();
    }

    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Dark Mode.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Dark Mode.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n")
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,#444444`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT,Lucida Sans Unicode`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,Gray`r`n")
        $theme_writer.write("DIALOG_FONT_SIZE,14.5`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,Silver`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,#5a5a5a`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,Silver`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,#6262ff`r`n")
        $theme_writer.write("EDITOR_FONT,Segoe UI`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,White`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,12`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,Yellow`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,Red`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,#00ff40`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT,Segoe UI`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,White`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,12`r`n")
        $theme_writer.write("INTERFACE_FONT,Segoe UI`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,11`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,White`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,Black`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,White`r`n")
        $theme_writer.close();
    }

    if(!(Test-Path -LiteralPath "$dir\Resources\Themes\Justine.csv"))
    {
        $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\Justine.csv",$true)
        $theme_writer.write("PROPERTY,VALUE`r`n")
        $theme_writer.write("ADJUSTMENT_BAR_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_BACKGROUND_COLOR,#000071`r`n")
        $theme_writer.write("DIALOG_BUTTON_BACKGROUND_COLOR,#d2d2d2`r`n")
        $theme_writer.write("DIALOG_BUTTON_TEXT_COLOR,Black`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_BACKGROUND_COLOR,blue`r`n")
        $theme_writer.write("DIALOG_DROPDOWN_TEXT_COLOR,white`r`n")
        $theme_writer.write("DIALOG_FONT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_INPUT_BACKGROUND_COLOR,#0606ff`r`n")
        $theme_writer.write("DIALOG_INPUT_TEXT_COLOR,White`r`n")
        $theme_writer.write("DIALOG_SUB_HEADER_COLOR,White`r`n")
        $theme_writer.write("DIALOG_TITLE_BANNER_COLOR,White`r`n")
        $theme_writer.write("DIALOG_TITLE_FONT_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_BACKGROUND_COLOR,White`r`n")
        $theme_writer.write("EDITOR_EXTEND_ACRONYM_FONT_COLOR,Blue`r`n")
        $theme_writer.write("EDITOR_FONT,Times New Roman`r`n")
        $theme_writer.write("EDITOR_FONT_COLOR,Black`r`n")
        $theme_writer.write("EDITOR_FONT_SIZE,12`r`n")
        $theme_writer.write("EDITOR_HIGHLIGHT_COLOR,Yellow`r`n")
        $theme_writer.write("EDITOR_MISSPELLED_FONT_COLOR,#ca0000`r`n")
        $theme_writer.write("EDITOR_SHORTEN_ACRONYM_FONT_COLOR,#ff64b1`r`n")
        $theme_writer.write("FEEDER_BACKGROUND_COLOR,#8affd0`r`n")
        $theme_writer.write("FEEDER_FONT,Times New Roman`r`n")
        $theme_writer.write("FEEDER_FONT_COLOR,Black`r`n")
        $theme_writer.write("FEEDER_FONT_SIZE,12`r`n")
        $theme_writer.write("INTERFACE_FONT,Times New Roman`r`n")
        $theme_writer.write("INTERFACE_FONT_SIZE,11`r`n")
        $theme_writer.write("MAIN_BACKGROUND_COLOR,Silver`r`n")
        $theme_writer.write("MENU_BACKGROUND_COLOR,Silver`r`n")
        $theme_writer.write("MENU_TEXT_COLOR,Black`r`n")
        $theme_writer.write("SIDEKICK_BACKGROUND_COLOR,#c9c9c9`r`n")
        $theme_writer.write("TEXT_CALCULATOR_BACKGROUND_COLOR,#99ccff`r`n")
        $theme_writer.write("TEXT_CALCULATOR_OVER_COLOR,Red`r`n")
        $theme_writer.write("TEXT_CALCULATOR_UNDER_COLOR,Black`r`n")
        $theme_writer.close();
    }
}
################################################################################
######Load Package###############################################################
function load_package
{
    $script:lock_history = 1;
    $script:history = @{};
    $script:old_text = "";
    $sizer_box.text = "";
    [int]$script:history_system_location = 0
    [int]$script:history_user_location = 0
    [int]$script:save_history_tracker = 0
    $script:sidekick_results = "";
    $package = $script:settings['PACKAGE']
    if(Test-Path -literalpath "$dir\Resources\Packages\$package")
    {
        #write-host Loading Package: $package
        #####Load History
        if(Test-Path -LiteralPath "$dir\Resources\Packages\$package\History.txt")
        {
            $reader = New-Object IO.StreamReader "$dir\Resources\Packages\$package\History.txt"
            [int]$place = 0;
            $mode = ""
            [int]$start = 0;
            [int]$end = 0;
            $entry = "" 

            $history_file = Get-content -tail $script:settings['SAVE_HISTORY_THRESHOLD'] "$dir\Resources\Packages\$package\History.txt"; #Updated history method (Ver 1.3 Update)
            foreach($line in $history_file)
            {
                ([int]$place,$mode,[int]$start,[int]$end,$entry) = $line -split '::'
                $entry = $entry -replace "<RETURN>","`n"
                if((!($history.contains($place))) -and ($place -match "\d"))
                {
                    $script:history.add($place,"$mode::$start::$end::$entry")
                }
            }
            [int]$script:history_system_location = $place
            [int]$script:history_user_location = $place
            [int]$script:save_history_tracker = $place
            $reader.close();
        }
        ###########Load Snapshot
        if(Test-Path -LiteralPath "$dir\Resources\Packages\$package\Snapshot.txt")
        {
            $slurp = Get-Content -LiteralPath "$dir\Resources\Packages\$package\Snapshot.txt" -raw
            $slurp = $slurp.substring(0,$slurp.Length - 2)
            $editor.text = $slurp
            $script:old_text = $editor.text
            $editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
            $editor.selectionstart = $editor.text.length
            
        }
        else
        {
            #write-host "ERROR: Package Snapshot missing or empty"
            $editor.text = ""
            $script:old_text = $editor.text
            $editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        }
        $script:Form.Text = "$script:program_title ($package)"
        while($editor.ZoomFactor -ne $script:zoom) 
        {
            #Zoom Changes during RTF replace, but won't change in time... this is a work around.
            $editor.ZoomFactor = $script:zoom             
        }
    }
    else
    {
        ##Package Missing
        $package = "Current"
        $script:settings['PACKAGE'] = "Current"
        update_settings
        if(Test-Path -literalpath "$dir\Resources\Packages\$package")
        {   
            load_package #Load the Default Package
        }
        else
        {
            ##Create a Default Package 
            New-Item  -ItemType directory -Path "$dir\Resources\Packages\$package"
            #save_package_history
        }
        $script:Form.Text = "$script:program_title (Current)"
        $editor.text = "";
        #$editor.Font = [Drawing.Font]::New('Times New Roman', 14)
        $editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        while($editor.ZoomFactor -ne $script:zoom) 
        {
            #Zoom Changes during RTF replace, but won't change in time... this is a work around.
            $editor.ZoomFactor = $script:zoom          
        }
        
    }
    $Script:recent_editor_text = "Changed Loaded Package";
    $script:lock_history = 0;
    $Script:Timer.Start()
  
}
################################################################################
#####Update Settings############################################################
function update_settings
{
    if($script:settings.count -ne 0)
    {
        if(Test-Path "$dir\Resources\Required\Buffer_Settings.csv")
        {
            Remove-Item "$dir\Resources\Required\Buffer_Settings.csv"
        }
        $buffer_settings = new-object system.IO.StreamWriter("$dir\Resources\Required\Buffer_Settings.csv",$true)
        $buffer_settings.write("PROPERTY,VALUE`r`n");
        foreach($setting in $script:settings.getEnumerator() | Sort key)                  #Loop through Input Entries
        {
                $setting_key = $setting.Key                                               
                $setting_value = $setting.Value
                $buffer_settings.write("$setting_key,$setting_value`r`n");
        }
        $buffer_settings.close();
        if(test-path -LiteralPath "$dir\Resources\Required\Buffer_Settings.csv")
        {
            if(Test-Path -LiteralPath "$dir\Resources\Required\Settings.csv")
            {
                Remove-Item -LiteralPath "$dir\Resources\Required\Settings.csv"
            }
            Rename-Item -LiteralPath "$dir\Resources\Required\Buffer_Settings.csv" "$dir\Resources\Required\Settings.csv"
        }
    } 

}
################################################################################
######Load Settings##############################################################
function load_settings
{
    if(Test-Path "$dir\Resources\Required\Settings.csv")
    {
        ################################################################################
        ######Load All Settings#########################################################
        $line_count = 0;
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Settings.csv")
        while($null -ne ($line = $reader.ReadLine()))
        {
            $line_count++;
            if($line_count -ne 1)
            {
                ($key,$value) = $line -split ',',2
                if(!($script:settings.containskey($key)))
                {
                    $script:settings.Add($key,$value);
                }
                #write-host $key
                #write-host $value
            } 
        }
        $reader.close();

        ################################################################################
        ######Verify All Settings Loaded################################################
        $changes = 0;
        if($script:settings['SAVE_HISTORY_THRESHOLD'] -eq $null)  {$changes = 1; $script:settings['SAVE_HISTORY_THRESHOLD'] = 200}
        if($script:settings['MEMORY_FLUSHING'] -eq $null)         {$changes = 1; $script:settings['MEMORY_FLUSHING'] = 3}
        if($script:settings['THEME'] -eq $null)                   {$changes = 1; $script:settings['THEME'] = "Dark Castle"}
        if($script:settings['PACKAGE'] -eq $null)                 {$changes = 1; $script:settings['PACKAGE'] = "Current"}
        if($script:settings['LOAD_PACKAGES_AS_BULLETS'] -eq $null){$changes = 1; $script:settings['LOAD_PACKAGES_AS_BULLETS'] = 1}
        if($script:settings['TEXT_COMPRESSION'] -eq $null)        {$changes = 1; $script:settings['TEXT_COMPRESSION'] = 4}
        if($script:settings['CLOCK_SPEED'] -eq $null)             {$changes = 1; $script:settings['CLOCK_SPEED'] = 500}
        if($script:settings['SIZER_BOX_INVERTED'] -eq $null)      {$changes = 1; $script:settings['SIZER_BOX_INVERTED'] = 1}
        if($changes -eq 1)
        {
            update_settings
        }
    }
    $Script:Timer.Interval = $script:settings['CLOCK_SPEED'];
    #write-host ClocK Speed: $script:settings['CLOCK_SPEED']
    $Script:Timer.Start()
    $Script:Timer.Add_Tick({Idle_Timer})
}
################################################################################
######CSV Line to Array#########################################################
function csv_line_to_array ($line)
{
    if($line -match "^,")
    {
        $line = ",$line"; 
    }
    Select-String '(?:^|,)(?=[^"]|(")?)"?((?(1)[^"]*|[^,"]*))"?(?=,|$)' -input $line -AllMatches | Foreach { [System.Collections.ArrayList]$line_split = $_.matches -replace '^,|"',''}
    return $line_split
}
################################################################################
######CSV Write Line #########################################################
function csv_write_line ($write_line,$data)
{
    ##################################################
    #Function checks to see if there is a comma in the data about to be written
    $return = "";
    if($data -match ',')
    {
        $data = '"' + "$data" + '"'
    }
    if($write_line -eq "")
    {
        $return = "$data"
    }
    else
    {
        $return = "$write_line," + "$data"
    }
    return $return
}
################################################################################
######Load Dictionary###########################################################
function load_dictionary
{
    if(Test-Path "$dir\Resources\Required\Dictionary.txt")
    { 
        $reader = [System.IO.File]::OpenText("$dir\Resources\Required\Dictionary.txt")
        while($null -ne ($line = $reader.ReadLine()))
        {
            if(!($script:dictionary.ContainsKey($line)))
            {
                        $script:dictionary.Add($line,"");
            }
        }
        $reader.Close();
    }
}
################################################################################
######Move Caret on Right Click#################################################
#Rich Text Box won't move the Caret on a right click this is a complex work around.
$cSource = @'
    using System;
    using System.Drawing;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    public class Clicker
    {
        const int MOUSEEVENTF_LEFTDOWN   = 0x0002 ;
        const int MOUSEEVENTF_LEFTUP     = 0x0004 ;

        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        { 
            public int        type;
            public MOUSEINPUT mi;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEINPUT
        {
            public int    dx ;
            public int    dy ;
            public int    mouseData ;
            public int    dwFlags;
            public int    time;
            public IntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        extern static uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);
        public static void LeftClickAtPoint(int x, int y)
        {
            INPUT[] input = new INPUT[3];
            input[1].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
            input[2].mi.dwFlags = MOUSEEVENTF_LEFTUP;
            SendInput(3, input, Marshal.SizeOf(input[0])); 
        }
    }
'@
if (!("Clicker" -as [type])) 
{
    Add-Type -TypeDefinition $cSource -ReferencedAssemblies System.Windows.Forms,System.Drawing
}
################################################################################
######Sync Scrolling############################################################
$typeDef = @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public enum ScrollBarType : uint {
    SbHorz = 0,
    SbVert = 1,
    SbCtl  = 2,
    SbBoth = 3
}
public enum ScrollBarTypes : uint {
    SbHorz = 0,
    SbVert = 1,
    SbCtl  = 2,
    SbBoth = 3
}

public enum Message : uint {
    WmVScroll = 0x0115
}

public enum ScrollBarCommands : uint {
    ThumbPosition = 4,
    ThumbTrack    = 5
}

[Flags()]
public enum ScrollBarInfo : uint {
    Range           = 0x0001,
    Page            = 0x0002,
    Pos             = 0x0004,
    DisableNoScroll = 0x0008,
    TrackPos        = 0x0010,

    All = ( Range | Page | Pos | TrackPos )
}

public class CustomRichTextBox : RichTextBox {
    public Control Buddy { get; set; }

    public bool ThumbTrack = false;

    [StructLayout( LayoutKind.Sequential )]
    public struct ScrollInfo {
        public uint cbSize;
        public uint fMask;
        public int nMin;
        public int nMax;
        public uint nPage;
        public int nPos;
        public int nTrackPos;
    };

    [DllImport( "User32.dll" )]
    public extern static int SendMessage( IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam );
    [DllImport( "User32.dll" )]
    public extern static int GetScrollInfo( IntPtr hWnd, int fnBar, ref ScrollInfo lpsi );

    public void CustomVScroll() {
        int nPos;

        ScrollInfo scrollInfo = new ScrollInfo();
        scrollInfo.cbSize = (uint)Marshal.SizeOf( scrollInfo );

        if (ThumbTrack) {
            scrollInfo.fMask = (uint)ScrollBarInfo.TrackPos;
            GetScrollInfo( this.Handle, (int)ScrollBarType.SbVert, ref scrollInfo );
            nPos = scrollInfo.nTrackPos;
        } else {
            scrollInfo.fMask = (uint)ScrollBarInfo.Pos;
            GetScrollInfo( this.Handle, (int)ScrollBarType.SbVert, ref scrollInfo );
            nPos = scrollInfo.nPos;
        }

        nPos <<= 16;
        uint wParam = (uint)ScrollBarCommands.ThumbPosition | (uint)nPos;
        SendMessage( Buddy.Handle, (int)Message.WmVScroll, new IntPtr( wParam ), new IntPtr( 0 ) );
    }


    public void CustomVScrolls() {
        int nPos;

        ScrollInfo scrollInfo = new ScrollInfo();
        scrollInfo.cbSize = (uint)Marshal.SizeOf( scrollInfo );

        scrollInfo.fMask = (uint)ScrollBarInfo.TrackPos;
        GetScrollInfo( this.Handle, (int)ScrollBarTypes.SbVert, ref scrollInfo );
        nPos = scrollInfo.nTrackPos;


        nPos <<= 16;
        uint wParam = (uint)ScrollBarCommands.ThumbPosition | (uint)nPos;
        SendMessage( Buddy.Handle, (int)Message.WmVScroll, new IntPtr( wParam ), new IntPtr( 0 ) );
    }



    protected override void WndProc( ref System.Windows.Forms.Message m ) {
        if ( m.Msg == (int)Message.WmVScroll ) {
            if ( ( m.WParam.ToInt32() & 0xFF ) == (int)ScrollBarCommands.ThumbTrack ) {
                ThumbTrack = true;
            } else {
                ThumbTrack = false;
            }
        }

        base.WndProc( ref m );
    }
}
"@
$assemblies = ("System.Windows.Forms", "System.Runtime.InteropServices")
Add-Type -ReferencedAssemblies $assemblies -TypeDefinition $typeDef -Language CSharp
################################################################################
######Levenshtein Distance######################################################
function levenshtein($first,$second)
{
    $ignoreCase = $true

    $len1 = $first.length
    $len2 = $second.length

    # If either string has length of zero, the # of edits/distance between them
    # is simply the length of the other string
    #######################################
    if($len1 -eq 0)
    { return $len2 }

    if($len2 -eq 0)
    { return $len1 }

    # make everything lowercase if ignoreCase flag is set
    if($ignoreCase -eq $true)
    {
      $first = $first.tolowerinvariant()
      $second = $second.tolowerinvariant()
    }
    # create 2d Array to store the "distances"
    $dist = new-object -type 'int[,]' -arg ($len1+1),($len2+1)

    # initialize the first row and first column which represent the 2
    # strings we're comparing
    for($i = 0; $i -le $len1; $i++) 
    {  $dist[$i,0] = $i }
    for($j = 0; $j -le $len2; $j++) 
    {  $dist[0,$j] = $j }

    $cost = 0

    for($i = 1; $i -le $len1;$i++)
    {
      for($j = 1; $j -le $len2;$j++)
      {
        if($second[$j-1] -ceq $first[$i-1])
        {
          $cost = 0
        }
        else   
        {
          $cost = 1
        }
        $tempmin = [System.Math]::Min(([int]$dist[($i-1),$j]+1) , ([int]$dist[$i,($j-1)]+1))
        $dist[$i,$j] = [System.Math]::Min($tempmin, ([int]$dist[($i-1),($j-1)] + $cost))
      }
    }
    # the actual distance is stored in the bottom right cell
    return $dist[$len1, $len2];
}
################################################################################
######Character Blocks##########################################################
function load_character_blocks
{
    $character_blocks.add("A",41.0680228862047)
    $character_blocks.add("B",37.9749827637152)
    $character_blocks.add("C",37.9749827637152)
    $character_blocks.add("D",41.0680228862047)
    $character_blocks.add("E",34.7498655190963)
    $character_blocks.add("F",31.6279069767442)
    $character_blocks.add("G",41.0680228862047)
    $character_blocks.add("H",41.0680228862047)
    $character_blocks.add("I",18.9544721013252)
    $character_blocks.add("J",22.1391723031067)
    $character_blocks.add("K",41.0680228862047)
    $character_blocks.add("L",34.7498655190963)
    $character_blocks.add("M",50.6029819237366)
    $character_blocks.add("N",41.0680228862047)
    $character_blocks.add("O",41.0680228862047)
    $character_blocks.add("P",31.6279069767442)
    $character_blocks.add("Q",41.0680228862047)
    $character_blocks.add("R",37.9749827637152)
    $character_blocks.add("S",31.6279069767442)
    $character_blocks.add("T",34.7498655190963)
    $character_blocks.add("U",41.0680228862047)
    $character_blocks.add("V",41.0680228862047)
    $character_blocks.add("W",53.7342657342657)
    $character_blocks.add("X",41.0680228862047)
    $character_blocks.add("Y",41.0680228862047)
    $character_blocks.add("Z",34.7498655190963)
    $character_blocks.add("a",25.242794588589)
    $character_blocks.add("b",28.4313581155686)
    $character_blocks.add("c",25.242794588589)
    $character_blocks.add("d",28.4313581155686)
    $character_blocks.add("e",25.242794588589)
    $character_blocks.add("f",18.9544721013252)
    $character_blocks.add("g",28.4313581155686)
    $character_blocks.add("h",28.4313581155686)
    $character_blocks.add("i",15.8139534883721)
    $character_blocks.add("j",15.8139534883721)
    $character_blocks.add("k",28.4313581155686)
    $character_blocks.add("l",15.8139534883721)
    $character_blocks.add("m",44.2783446062135)
    $character_blocks.add("n",28.4313581155686)
    $character_blocks.add("o",28.4313581155686)
    $character_blocks.add("p",28.4313581155686)
    $character_blocks.add("q",28.4313581155686)
    $character_blocks.add("r",18.9544721013252)
    $character_blocks.add("s",22.1391723031067)
    $character_blocks.add("t",15.8139534883721)
    $character_blocks.add("u",28.4313581155686)
    $character_blocks.add("v",28.4313581155686)
    $character_blocks.add("w",41.0680228862047)
    $character_blocks.add("x",28.4313581155686)
    $character_blocks.add("y",28.4313581155686)
    $character_blocks.add("z",25.242794588589)
    $character_blocks.add("0",28.4313581155686)
    $character_blocks.add("1",28.4313581155686)
    $character_blocks.add("2",28.4313581155686)
    $character_blocks.add("3",28.4313581155686)
    $character_blocks.add("4",28.4313581155686)
    $character_blocks.add("5",28.4313581155686)
    $character_blocks.add("6",28.4313581155686)
    $character_blocks.add("7",28.4313581155686)
    $character_blocks.add("8",28.4313581155686)
    $character_blocks.add("9",28.4313581155686)
    $character_blocks.add("``",18.9544721013252)
    $character_blocks.add("~",30.8010171646535)
    $character_blocks.add("!",18.9544721013252)
    $character_blocks.add("@",52.4009324009324)
    $character_blocks.add("#",28.4313581155686)
    $character_blocks.add("$",28.4313581155686)
    $character_blocks.add("%",47.3855968592811)
    $character_blocks.add("^",26.6952849131067)
    $character_blocks.add("&",44.2783446062135)
    $character_blocks.add("*",28.4313581155686)
    $character_blocks.add("(",18.9544721013252)
    $character_blocks.add(")",18.9544721013252)
    $character_blocks.add("-",18.9544721013252)
    $character_blocks.add("_",28.4313581155686)
    $character_blocks.add("+",32.0979020979021)
    $character_blocks.add("=",32.0979020979021)
    $character_blocks.add("{",27.330649148831)
    $character_blocks.add("}",27.330649148831)
    $character_blocks.add("[",18.9544721013252)
    $character_blocks.add("]",18.9544721013252)
    $character_blocks.add("\",15.8139534883721)
    $character_blocks.add(";",15.8139534883721)
    $character_blocks.add(":",15.8139534883721)
    $character_blocks.add("`"",23.2478632478632)
    $character_blocks.add("'",10.2462066235651)
    $character_blocks.add("<",32.0979020979021)
    $character_blocks.add(">",32.0979020979021)
    $character_blocks.add(",",14.2159411269359)
    $character_blocks.add(".",14.2159411269359)
    $character_blocks.add("/",15.8139534883721)
    $character_blocks.add("?",25.242794588589)    
    $character_blocks.add("|",11.3886113886114)
    $character_blocks.add(" ",18.9877255611521)
    $character_blocks.add(" ",14.2159411269359) #Regular Space
    $character_blocks.add(" ",14.2159411269359) #Puncuation Space
    $character_blocks.add(" ",11.3886113886114) #2009
    $character_blocks.add(" ",9.47735191637631) #2006 Hair Space
    $character_blocks.add(" ",4.75524475524475) #200A Thin Space

    ##
    $character_blocks.add("€",28.4313581155686)
    $character_blocks.add("…",56.9617616426127)
    $character_blocks.add("°",22.7772227772228)    
    $character_blocks.add("‘",18.9544721013252)
    $character_blocks.add("é",25.242794588589)
    #$character_blocks.add('—',56.9617616426127)
    #$character_blocks.add("”",25.242794588589)
    #$character_blocks.add("“",25.242794588589)
}
################################################################################
######Write History##############################################################
function write_history
{  
    if($editor.text -cne $script:old_text)
    {
        #####Delete Selected Text
        if($script:history_replace_text -ne "")
        {
            $script:history_system_location++
            $script:history.add($script:history_system_location,"D::$script:history_replace_text_start::$script:history_replace_text_end::$script:history_replace_text")
            $script:old_text = $script:old_text.substring(0,$script:history_replace_text_start) + $script:old_text.substring(($script:history_replace_text_start + $script:history_replace_text_end),($script:old_text.length - ($script:history_replace_text_start + $script:history_replace_text_end)))       
            $script:history_user_location = $script:history_system_location
            #write-host Added "D::$script:history_replace_text_start::$script:history_replace_text_end::$script:history_replace_text"
            $script:history_replace_text = ""
        }

        #####Append/Delete Text
        ($mode,$start,$end,$changes) = string_compare $editor.text $script:old_text
        #write-host ------------------------------
        #write-host S1 = $editor.text
        #write-host S2 = $old_text
        #write-host M $mode
        #write-host S $start
        #write-host E $end
        #write-host C -$changes-


        $script:old_text = $editor.text    
        $script:history_system_location++
        $script:history.add($script:history_system_location,"$mode::$start::$end::$changes")
    
        if($script:text_lock -eq -1)
        {
            #write-host USER LOCK
            $script:history_user_location = $script:history_system_location
        }
        #write-host Added "$mode::$start::$end::$changes"
    }
}
################################################################################
######Calculate String##########################################################
function calc_string($string1,$string2)
{
    $start = 0;
    While(($start -le ($string1.Length -1)) -and ($start -le ($string2.Length -1)))
    { 
        if(!($string1.substring($start,1) -ceq $string2.substring($start,1)))
        {
            break;
        }
        $start++
    }
    return($start)
}
################################################################################
######String Compare############################################################
function string_compare($string1,$string2)
{
    #write-host S1= $string1
    #write-host S2= $string2
    $start = 0
    $end = 0;
    $mode = "F"
    #F = Failed
    #M = Strings The Same
    #A = Appended
    #D = String Deleted
    ###############################################
    if($string1.Length -ge $string2.Length)
    {
        if($string1 -ceq $string2)
        {
            return('M','0','0',"")
        }
        $mode = "Append"
        $start = calc_string $string1 $string2
        $end = $string1.Length - $string2.Length
        $changes = $string1.Substring($start,$end)
        return('A',$start,$end,$changes)
    }
    else
    {
        
        $mode = "Delete"
        $start = calc_string $string2 $string1
        $end = $string2.Length - $string1.length
        $changes = $string2.Substring($start,$end)
        return('D',$start,$end,$changes)
    }   
    ###############################################
}
################################################################################
######Undo History##############################################################
function undo_history
{
    #write-host UNDO $script:history[$script:history_user_location]
    if(($script:history_user_location -le $script:history_system_location) -and ($script:history_user_location -gt 0))
    {
        if($script:history.contains($script:history_user_location))
        {
            ###Try to get from memory
            ($mode,$start,$end,$change) = $script:history[$script:history_user_location] -split "::"
            #write-host "Got it from Memory"
        }
        else
        {
            if($script:history_user_location -le 1)
            {
                #$script:history_user_location = 2;
            }
            ###Try to get from file
            $package = $script:settings['PACKAGE']
            $data = Get-Content -literalpath "$dir\Resources\Packages\$package\History.txt" | Select -Index ($script:history_user_location - 1)
            ($place,$mode,$start,$end,$change) = $data -split "::"
            $change = $change -replace "<RETURN>","`n"
            if($place -ne $script:history_user_location)
            {
                write-host "ERROR: History File Misaligned $place = $script:history_user_location"
            }
            
            #write-host "Got it from File $data"
        }
        if($mode -ne "")
        {
            
            if($mode -eq "A")
            {
                #write-host Deleting $change
                $editor.SelectionStart = $start
                $editor.SelectionLength = $end
                $editor.SelectedText = ""

            }
            elseif($mode -eq "D")
            {
                #write-host Appending $change
                $editor.SelectionStart = $start
                $editor.SelectionLength = 0
                $editor.SelectedText = "$change"  
            }
            $script:history_user_location--
        }
        else
        {
            write-host ERROR: Failed History at $script:history_user_location
        }
    }
    else
    {
        #write-host "Undo Nope"
    }
}
################################################################################
######Redo History##############################################################
function redo_history
{
    if(($script:history_user_location -lt $script:history_system_location) -and ($script:history_user_location -ge 0) -and ($script:history_user_location -lt $script:text_lock))
    {
        if($script:history.contains($script:history_user_location))
        {
            ###Try to get from memory
            ($mode,$start,$end,$change) = $script:history[$script:history_user_location] -split "::"
            #write-host "Got it from Memory"
        }
        else
        {
            ###Try to get from file
            #if($script:history_user_location -le $script:history_system_location)
            #{
                $package = $script:settings['PACKAGE']
                if($script:history_user_location -le 1) {$script:history_user_location = 1}
                $data = Get-Content -literalpath "$dir\Resources\Packages\$package\History.txt" | Select -Index ($script:history_user_location - 1)
                ($place,$mode,$start,$end,$change) = $data -split "::"
                $change = $change -replace "<RETURN>","`n"
                if($place -ne $script:history_user_location)
                {
                    write-host "History File Misaligned $place = $script:history_user_location"
                }
                #write-host "Got it from File $data"
            #}
        }
        if($mode -ne "")
        {
            $script:history_user_location++
            if($mode -eq "D")
            {
                #write-host Deleting $change
                $editor.SelectionStart = $start
                $editor.SelectionLength = $end
                $editor.SelectedText = ""

            }
            elseif($mode -eq "A")
            {
                #write-host Appending $change
                $editor.SelectionStart = $start
                $editor.SelectionLength = 0
                $editor.SelectedText = $change
            }
        }
        else
        {
            write-host ERROR: Failed History at $script:history_user_location
        }      
    }
    else
    {
        #write-host Redo Nope
    }
}
################################################################################
######Save History##############################################################
function save_history
{
    #write-host Saving History
    if(($script:save_history_job -eq "") -and ($script:save_history_tracker -lt $script:history_system_location) -and ($script:lock_history -ne 1))
    {
        #write-host Running save Job
        #write-host Last Saved = $script:save_history_tracker
        #write-host Sys Location = $script:history_system_location

        $script:save_history_job = Start-Job -ScriptBlock {

            $history = $using:history
            $tracker = $using:save_history_tracker
            $dir = $using:dir
            $package = $using:settings['PACKAGE']
            $editor_text = $using:editor.text

            Set-Content -literalpath "$dir\Resources\Packages\$package\Snapshot.txt" $editor_text -encoding Unicode

            $writer = [IO.StreamWriter]::new("$dir\Resources\Packages\$package\History.txt", $true, [Text.Encoding]::UTF8)
            foreach($item in $history.GetEnumerator() | sort key)
            {
                $item.value = $item.value -replace "`n", "<RETURN>"
                if($item.key -gt $tracker)
                {
                    $tracker = $item.key
                    [string]$line = [string]$item.key + "::" + [string]$item.value
                    $writer.WriteLine($line)
                }
            }
            $writer.close()

            return $tracker
        }
    }
    else
    {
        if(($script:save_history_job -ne "") -and ($script:save_history_job.state -eq "Completed"))
        {
            #write-host "Job Finished"
            $script:save_history_tracker = Receive-Job -Job $script:save_history_job
            $script:save_history_job = "";
            #######################################################Scrub History in Memory (Ver 1.3 Update)
            
            $package = $script:settings['PACKAGE'];
            if(Test-path -LiteralPath "$dir\Resources\Packages\$package\History.txt")
            {
                $last_line = 0;
                (Get-Content "$dir\Resources\Packages\$package\History.txt" -read 100 | % { $last_line += $_.Length })
                $remove_counter = ($last_line  - $script:settings['SAVE_HISTORY_THRESHOLD'])
                #write-host Last Line: $last_line
                #write-host System Location: $script:history_system_location
                #write-host History Tracker: $script:save_history_tracker
                #write-host History Count: $script:history.count
                #write-host Remove Count:$remove_counter
                while($remove_counter -gt 0)
                {
                   
                       if($script:history.Contains($remove_counter))
                       {
                            #write-host $remove_counter
                            $script:history.remove($remove_counter)
                       }
                       if($script:history.count -le $script:settings['SAVE_HISTORY_THRESHOLD'])
                       {
                            #write-host Stopped $remove_counter
                            break;
                       }
                       $remove_counter--;
                }
                #write-host $script:history.count
            }
            #######################################################
            
        }
    }
}
################################################################################
######Save Package Dialog#######################################################
function save_package_dialog
{
    $script:lock_history = 1;
    $script:return = 0;
    $current_package = "";
    if($script:settings['PACKAGE'] -eq "Current")
    {
        $current_package = ""
    }
    else
    {
        $current_package = $script:settings['PACKAGE']
    }
    $save_pacakge_form = New-Object System.Windows.Forms.Form
    $save_pacakge_form.FormBorderStyle = 'Fixed3D'
    $save_pacakge_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $save_pacakge_form.Location = new-object System.Drawing.Point(0, 0)
    $save_pacakge_form.Size = new-object System.Drawing.Size(440, 120)
    $save_pacakge_form.MaximizeBox = $false
    $save_pacakge_form.SizeGripStyle = "Hide"
    $save_pacakge_form.Text = "Save Package"
    #$save_pacakge_form.TopMost = $True
    $save_pacakge_form.TabIndex = 0
    $save_pacakge_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $package_name_label                          = New-Object system.Windows.Forms.Label
    $package_name_label.text                     = "Package Name:";
    $package_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $package_name_label.Anchor                   = 'top,right'
    $package_name_label.width                    = 160
    $package_name_label.height                   = 30
    $package_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $package_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))

    $save_pacakge_form.controls.Add($package_name_label);

    $package_name_input                         = New-Object system.Windows.Forms.TextBox                       
    $package_name_input.AutoSize                 = $true
    $package_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $package_name_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $package_name_input.Anchor                   = 'top,left'
    $package_name_input.width                    = 250
    $package_name_input.height                   = 30
    $package_name_input.location                 = New-Object System.Drawing.Point(($package_name_label.Location.x + $package_name_label.Width + 5) ,12)
    $package_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $package_name_input.text                     = $current_package
    $package_name_input.name                     = $current_package
    $package_name_input.Add_TextChanged({
        $caret = $package_name_input.SelectionStart;
        $package_name_input.text = $package_name_input.text -replace '[^0-9A-Za-z ,-]', ''
        $package_name_input.text = $package_name_input.text.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '

        #$package_name_input.text = (Get-Culture).TextInfo.ToTitleCase($package_name_input.text)
        $package_name_input.SelectionStart = $caret
    });
    $save_pacakge_form.controls.Add($package_name_input);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($save_pacakge_form.width / 2) - ($submit_button.width)),45);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Save"
    $submit_button.Name = ""
    $submit_button.Add_Click({ 
        [array]$errors = "";
        $og_package = $script:settings['PACKAGE']
        $new_package = $package_name_input.text

        if($new_package -eq "")
        {
            $errors += "You must provide a name."
        }
        if($new_package -eq "Current")
        {
            $errors += "`"Current`" is a system file, and cannot be used."
        }
        if($errors.count -eq 1)
        {
            if(Test-Path -literalpath "$dir\Resources\Packages\$new_package")
            {
                $message = "`"$new_package`" already exists. Overwrite?`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Overwrite?", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                      if(!($og_package -eq $new_package))
                      {
                        #Packages are different
       
                        $value = $script:package_list[$og_package]
                        $script:package_list.remove($og_package);
                        $script:package_list.remove($new_package);
                        $script:package_list.add($new_package,$value);
                        

                        Remove-item "$dir\Resources\Packages\$new_package" -Recurse
                        Copy-item "$dir\Resources\Packages\$og_package" "$dir\Resources\Packages\$new_package" -Recurse
                        $script:settings['PACKAGE'] = $new_package
                        save_history
                      }
                      elseif(Test-path -literalpath "$dir\Resources\Packages\$new_package")
                      {
                        #Package is the same
                        $script:settings['PACKAGE'] = $new_package
                        save_history
                      }
                      if(($og_package -eq "Current") -and (Test-path -literalpath "$dir\Resources\Packages\$new_package") -and (test-path -literalpath "$dir\Resources\Packages\Current"))
                      {
                            Remove-Item "$dir\Resources\Packages\Current" -Recurse
                            $script:package_list.remove("Current");
                      }
                      $script:Form.Text = "$script:program_title ($new_package)"
                      update_settings
                      save_package_tracker
                      $script:return = 1;
                      $script:recent_editor_text = "Changed"; 
                      $save_pacakge_form.close();
                      
                }
            }
            else
            {
                $message = "Are you sure you want to save package as `"$new_package`"`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Save?", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    Copy-item "$dir\Resources\Packages\$og_package" "$dir\Resources\Packages\$new_package" -Recurse
                    $script:settings['PACKAGE'] = $new_package

                    $script:package_list.add($new_package,$script:package_list[$og_package]);
                    $script:package_list.remove($og_package);


                    save_history
                    if(($og_package -eq "Current") -and (Test-path -literalpath "$dir\Resources\Packages\$new_package") -and (test-path -literalpath "$dir\Resources\Packages\Current"))
                    {
                        Remove-Item "$dir\Resources\Packages\Current" -Recurse
                        $script:package_list.remove("Current");
                    }
                    
                    $script:Form.Text = "$script:program_title ($new_package)"
                    update_settings
                    save_package_tracker
                    $script:return = 1;
                    $script:recent_editor_text = "Changed"; 
                    $save_pacakge_form.close();
                    
                }
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }



    });
    $save_pacakge_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($save_pacakge_form.width / 2)),45);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
        $script:return = 0;
        $save_pacakge_form.close();
            
    });
    $save_pacakge_form.controls.Add($cancel_button) 


    $null = $save_pacakge_form.ShowDialog()
    $script:lock_history = 0;  
}
################################################################################
######Manage Package Dialog#####################################################
function manage_package_dialog
{
    save_history
    #write-host $dir\Resources\Packages

    $packages = Get-ChildItem -LiteralPath "$dir\Resources\Packages" -Directory -Force -ErrorAction SilentlyContinue #| Select-Object FullName

    #write-host ($packages | Measure-Object).Count
     
    $spacer = 0;
    $manage_package_form = New-Object System.Windows.Forms.Form
    $manage_package_form.FormBorderStyle = 'Fixed3D'
    $manage_package_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $manage_package_form.Location = new-object System.Drawing.Point(0, 0)
    $manage_package_form.MaximizeBox = $false
    $manage_package_form.SizeGripStyle = "Hide"
    $manage_package_form.Width = 800
    if(($packages | Measure-Object).Count -eq 0)
    {
        $manage_package_form.Height = 200;
    }
    elseif(((($packages | Measure-Object).Count * 65) + 140) -ge 600)
    {
        $manage_package_form.Height = 600;
        $manage_package_form.Autoscroll = $true
        $spacer = 20
    }
    else
    {
        $manage_package_form.Height = ((($packages | Measure-Object).Count * 65) + 140)
    }
    $manage_package_form.Text = "Manage Packages"
    $manage_package_form.TabIndex = 0
    $manage_package_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    ################################################################################################
    $y_pos = 10;


    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "Manage Packages";
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($manage_package_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point((($manage_package_form.width / 2) - ($title_label.width / 2)),$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $manage_package_form.controls.Add($title_label);

    $y_pos = $y_pos + 40;

    $separator_bar                             = New-Object system.Windows.Forms.Label
    $separator_bar.text                        = ""
    $separator_bar.AutoSize                    = $false
    $separator_bar.BorderStyle                 = "fixed3d"
    #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
    $separator_bar.Anchor                      = 'top,left'
    $separator_bar.width                       = (($manage_package_form.width - 50) - $spacer)
    $separator_bar.height                      = 1
    $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
    $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $separator_bar.TextAlign                   = 'MiddleLeft'
    $manage_package_form.controls.Add($separator_bar);

    $y_pos = $y_pos + 5;

    if(($packages | Measure-Object).Count -ne 0)
    {
        #####################################################################################
        foreach($package in $packages | sort name)
        {

            
            #write-host $package.fullname
            #write-host $package.name
            #exit;


            $package_file = $package.fullname + "\Snapshot.txt"
            $package_name = $package.name
            $package_value =  $script:package_list[$package.name]


            $package_name_label                          = New-Object system.Windows.Forms.Label
            $package_name_label.text                     = "$package_name";
            $package_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $package_name_label.Anchor                   = 'top,right'
            $package_name_label.width                    = (($manage_package_form.width - 50) - $spacer)
            $package_name_label.height                   = 30
            $package_name_label.location                 = New-Object System.Drawing.Point((20 + $spacer),$y_pos)
            $package_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $manage_package_form.controls.Add($package_name_label);

            $y_pos = $y_pos + 30;
                    
            $load_button           = New-Object System.Windows.Forms.Button
            $load_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $load_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $load_button.Width     = 90
            $load_button.height     = 25
            $load_button.Location  = New-Object System.Drawing.Point(20,$y_pos);
            $load_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $load_button.Text      = "Load"
            $load_button.Name      = $package_name
            $load_button.Add_Click({
                $script:settings['PACKAGE'] = $this.name
                load_package
                update_settings
                $Script:recent_editor_text = "Changed"
            });
            $manage_package_form.controls.Add($load_button) 

            $delete_button           = New-Object System.Windows.Forms.Button
            $delete_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $delete_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $delete_button.Width     = 90
            $delete_button.height     = 25
            $delete_button.Location  = New-Object System.Drawing.Point(($load_button.Location.x + $load_button.width + 5),$y_pos);
            $delete_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $delete_button.Text      ="Delete"
            $delete_button.Name      = $package_name 
            $delete_button.Add_Click({
                $file = $this.name
                $message = "Are you sure you want to delete the `"$file`" package? You cannot revert this action.`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    if(Test-Path -LiteralPath "$dir\Resources\Packages\$file")
                    {
                        Remove-Item -LiteralPath "$dir\Resources\Packages\$file" -Recurse
                        $script:package_list.Remove($file);
                        save_package_tracker
                        build_bullet_menu     
                        if($script:settings['PACKAGE'] -eq $file)
                        {
                           $script:settings['PACKAGE'] = "Current"
                           load_package
                           update_settings
                        }
                        $Script:recent_editor_text = "Changed"
                        $script:reload_function = "manage_package_dialog" 
                        
                        $manage_package_form.close();
                        
                    }     
                }

            });
            $manage_package_form.controls.Add($delete_button)

            $rename_button           = New-Object System.Windows.Forms.Button
            $rename_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $rename_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $rename_button.Width     = 90
            $rename_button.height     = 25
            $rename_button.Location  = New-Object System.Drawing.Point(($delete_button.Location.x + $delete_button.width + 5),$y_pos);
            $rename_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $rename_button.Text      ="Rename"
            $rename_button.Name      = $package.fullname
            $rename_button.Add_Click({
                $old_name = $this.name
                $new_name = rename_dialog $old_name

                #write-host ON $old_name
                #write-host NN $new_name
                
                if(($new_name -cne $old_name) -and ($new_name -ne ""))
                {
                    $old_key = [System.IO.Path]::GetFileNameWithoutExtension($old_name)
                    $new_key = [System.IO.Path]::GetFileNameWithoutExtension($new_name)
                    $old_value = $script:package_list["$old_key"]
                    $script:package_list.remove("$old_key");
                    $script:package_list.add("$new_key",$old_value);
                    #write-host ON1 $old_key $old_value
                    #write-host NN1 $new_key
                    save_package_tracker
                    build_bullet_menu
                    if($script:settings['PACKAGE'] -eq $old_key)
                    {
                        $script:settings['PACKAGE'] = $new_key
                        load_package
                        update_settings
                    }    
                    $script:reload_function = "manage_package_dialog"
                    $Script:recent_editor_text = "Changed"
                    $manage_package_form.close();
                }
            });
            $manage_package_form.controls.Add($rename_button)
            


            $enable_checkbox = new-object System.Windows.Forms.checkbox
            $enable_checkbox.Location = new-object System.Drawing.Size(($rename_button.Location.x + $rename_button.width + 5),$y_pos);
            $enable_checkbox.Size = new-object System.Drawing.Size(300,30)
            $enable_checkbox.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $enable_checkbox.name = $package_name          
            $enable_checkbox.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            if($package_value -eq "0")
            {
                $enable_checkbox.Checked = $false
                $enable_checkbox.text = "Bullet Feed Disabled"
            }
            else
            {
                $enable_checkbox.Checked = $true
                $enable_checkbox.text = "Bullet Feed Enabled"
            }
            $enable_checkbox.Add_CheckStateChanged({
                if($this.Checked -eq $true)
                {
                    $this.text = "Bullet Feed Enabled"
                    $script:package_list[$this.name] = 1;
                    save_package_tracker
                    build_bullet_menu
                }
                else
                {
                    $this.text = "Bullet Feed Disabled"
                    $script:package_list[$this.name] = 0;
                    save_package_tracker
                    build_bullet_menu
                }
            })
            $manage_package_form.controls.Add($enable_checkbox);


            #######################################################
            $line_count = 0
            #write-host $package_file
            if(Test-Path -LiteralPath "$package_file")
            {
                #write-host $package_file
                $reader = New-Object IO.StreamReader $package_file
                while($null -ne ($line = $reader.ReadLine()))
                {
                    #write-host $line
                    if(($line -match "^- |^- |^- |^- ") -and ($line.length -ge 85))
                    {
                        #write-host $line
                        $line_count++;
                    }
                }
                $reader.Close()
            }
            $item_count_label                          = New-Object system.Windows.Forms.Label
            $item_count_label.text                     = "$line_count Bullets";
            $item_count_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $item_count_label.Anchor                   = 'top,right'
            $item_count_label.TextAlign = "MiddleRight"
            $item_count_label.width                    = 180
            $item_count_label.height                   = 30
            $item_count_label.location                 = New-Object System.Drawing.Point((($manage_package_form.width - 190) - $spacer),$y_pos);
            $item_count_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
            $manage_package_form.controls.Add($item_count_label);

            $y_pos = $y_pos + 30
            $separator_bar                             = New-Object system.Windows.Forms.Label
            $separator_bar.text                        = ""
            $separator_bar.AutoSize                    = $false
            $separator_bar.BorderStyle                 = "fixed3d"
            #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
            $separator_bar.Anchor                      = 'top,left'
            $separator_bar.width                       = (($manage_package_form.width - 50) - $spacer)
            $separator_bar.height                      = 1
            $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
            $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $separator_bar.TextAlign                   = 'MiddleLeft'
            $manage_package_form.controls.Add($separator_bar);
            $y_pos = $y_pos + 5
        }
    
        $manage_package_form.ShowDialog()
    }
    else
    {
        $message = "You have no Bullet Banks to manage.`nYou must create or import a Bullet Bank first."
        #[System.Windows.MessageBox]::Show($message,"No bank",'Ok')

        $error_label                          = New-Object system.Windows.Forms.Label
        $error_label.text                     = "$message";
        $error_label.ForeColor                = "Red"
        $error_label.Anchor                   = 'top,right'
        $error_label.width                    = ($manage_package_form.width - 10)
        $error_label.height                   = 50
        $error_label.TextAlign = "MiddleCenter"
        $error_label.location                 = New-Object System.Drawing.Point(10,$y_pos)
        $error_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $manage_package_form.controls.Add($error_label);
        $manage_package_form.ShowDialog()
    }
    
}
################################################################################
######Color Dialog##############################################################
function interface_dialog
{
    
    $v_spacer = 5
    $spacer = 20
    $color_form                                 = New-Object System.Windows.Forms.Form
    $title_label                                = New-Object system.Windows.Forms.Label
    $theme_combo                                = New-Object System.Windows.Forms.ComboBox
    $manage_theme_button                        = New-Object System.Windows.Forms.Button
    $separator_bar1                             = New-Object system.Windows.Forms.Label
    $save_theme_button                          = New-Object System.Windows.Forms.Button
    $cancel_theme_button                        = New-Object System.Windows.Forms.Button
    $separator_bar2                             = New-Object system.Windows.Forms.Label
    $header_label1                              = New-Object system.Windows.Forms.Label
    $main_background_color_label                = New-Object system.Windows.Forms.Label
    $main_background_color_input                = New-Object system.Windows.Forms.TextBox
    $main_background_color_button               = New-Object System.Windows.Forms.Button
    $menu_text_color_label                      = New-Object system.Windows.Forms.Label
    $menu_text_color_input                      = New-Object system.Windows.Forms.TextBox
    $menu_text_color_button                     = New-Object System.Windows.Forms.Button
    $menu_background_color_label                = New-Object system.Windows.Forms.Label
    $menu_background_color_input                = New-Object system.Windows.Forms.TextBox
    $menu_background_color_button               = New-Object System.Windows.Forms.Button
    $adjustment_bars_color_label                = New-Object system.Windows.Forms.Label
    $adjustment_bars_color_input                = New-Object system.Windows.Forms.TextBox  
    $adjustment_bars_color_button               = New-Object System.Windows.Forms.Button
    $header_label2                              = New-Object system.Windows.Forms.Label
    $editor_background_color_label              = New-Object system.Windows.Forms.Label
    $editor_background_color_input              = New-Object system.Windows.Forms.TextBox   
    $editor_background_color_button             = New-Object System.Windows.Forms.Button
    $editor_font_color_label                    = New-Object system.Windows.Forms.Label
    $editor_font_color_input                    = New-Object system.Windows.Forms.TextBox
    $editor_font_color_button                   = New-Object System.Windows.Forms.Button
    $editor_misspelled_font_color_label         = New-Object system.Windows.Forms.Label
    $editor_misspelled_font_color_input         = New-Object system.Windows.Forms.TextBox  
    $editor_misspelled_font_color_button        = New-Object System.Windows.Forms.Button
    $editor_extend_acronym_font_color_label     = New-Object system.Windows.Forms.Label
    $editor_extend_acronym_font_color_input     = New-Object system.Windows.Forms.TextBox  
    $editor_extend_acronym_font_color_button    = New-Object System.Windows.Forms.Button
    $editor_shorten_acronym_font_color_label    = New-Object system.Windows.Forms.Label
    $editor_shorten_acronym_font_color_input    = New-Object system.Windows.Forms.TextBox  
    $editor_shorten_acronym_font_color_button   = New-Object System.Windows.Forms.Button
    $editor_highlight_color_label               = New-Object system.Windows.Forms.Label
    $editor_highlight_color_input               = New-Object system.Windows.Forms.TextBox  
    $editor_highlight_color_button              = New-Object System.Windows.Forms.Button
    $editor_font_label                          = New-Object system.Windows.Forms.Label
    $editor_font_combo                          = New-Object System.Windows.Forms.ComboBox
    $editor_font_size_combo                     = New-Object System.Windows.Forms.ComboBox
    $header_label3                              = New-Object system.Windows.Forms.Label
    $text_caclulator_background_color_label     = New-Object system.Windows.Forms.Label
    $text_caclulator_background_color_input     = New-Object system.Windows.Forms.TextBox
    $text_caclulator_background_color_button    = New-Object System.Windows.Forms.Button
    $text_caclulator_under_color_label          = New-Object system.Windows.Forms.Label
    $text_caclulator_under_color_input          = New-Object system.Windows.Forms.TextBox  
    $text_caclulator_under_color_button         = New-Object System.Windows.Forms.Button
    $text_caclulator_over_color_label           = New-Object system.Windows.Forms.Label
    $text_caclulator_over_color_input           = New-Object system.Windows.Forms.TextBox 
    $text_caclulator_over_color_button          = New-Object System.Windows.Forms.Button
    $header_label4                              = New-Object system.Windows.Forms.Label
    $feeder_background_color_label              = New-Object system.Windows.Forms.Label
    $feeder_background_color_input              = New-Object system.Windows.Forms.TextBox 
    $feeder_background_color_button             = New-Object System.Windows.Forms.Button
    $feeder_font_color_label                    = New-Object system.Windows.Forms.Label
    $feeder_font_color_input                    = New-Object system.Windows.Forms.TextBox 
    $feeder_font_color_button                   = New-Object System.Windows.Forms.Button 
    $feeder_font_label                          = New-Object system.Windows.Forms.Label
    $feeder_font_combo                          = New-Object System.Windows.Forms.ComboBox
    $feeder_font_size_combo                     = New-Object System.Windows.Forms.ComboBox
    $header_label5                              = New-Object system.Windows.Forms.Label
    $sidekick_background_color_label            = New-Object system.Windows.Forms.Label
    $sidekick_background_color_input            = New-Object system.Windows.Forms.TextBox 
    $sidekick_background_color_button           = New-Object System.Windows.Forms.Button
    $header_label6                              = New-Object system.Windows.Forms.Label
    $dialog_background_color_label              = New-Object system.Windows.Forms.Label
    $dialog_background_color_input              = New-Object system.Windows.Forms.TextBox   
    $dialog_background_color_button             = New-Object System.Windows.Forms.Button
    $dialog_title_font_color_label              = New-Object system.Windows.Forms.Label
    $dialog_title_font_color_input              = New-Object system.Windows.Forms.TextBox 
    $dialog_title_font_color_button             = New-Object System.Windows.Forms.Button
    $dialog_title_banner_color_label            = New-Object system.Windows.Forms.Label
    $dialog_title_banner_color_input            = New-Object system.Windows.Forms.TextBox
    $dialog_title_banner_color_button           = New-Object System.Windows.Forms.Button
    $dialog_sub_header_color_label              = New-Object system.Windows.Forms.Label
    $dialog_sub_header_color_input              = New-Object system.Windows.Forms.TextBox  
    $dialog_sub_header_color_button             = New-Object System.Windows.Forms.Button
    $dialog_input_text_color_label              = New-Object system.Windows.Forms.Label
    $dialog_button_background_color_label       = New-Object system.Windows.Forms.Label
    $dialog_input_background_color_label        = New-Object system.Windows.Forms.Label
    $dialog_button_text_color_label             = New-Object system.Windows.Forms.Label
    $dialog_input_text_color_input              = New-Object system.Windows.Forms.TextBox  
    $dialog_button_text_color_input             = New-Object system.Windows.Forms.TextBox  
    $dialog_button_background_color_input       = New-Object system.Windows.Forms.TextBox
    $dialog_input_background_color_input        = New-Object system.Windows.Forms.TextBox
    $dialog_input_text_color_button             = New-Object System.Windows.Forms.Button    
    $dialog_button_background_color_button      = New-Object System.Windows.Forms.Button
    $dialog_input_background_color_button       = New-Object System.Windows.Forms.Button
    $dialog_button_text_color_button            = New-Object System.Windows.Forms.Button
    $dialog_font_color_label                    = New-Object system.Windows.Forms.Label
    $dialog_font_color_input                    = New-Object system.Windows.Forms.TextBox   
    $dialog_font_color_button                   = New-Object System.Windows.Forms.Button
    $interface_font_combo                       = New-Object System.Windows.Forms.ComboBox
    $interface_font_size_combo                  = New-Object System.Windows.Forms.ComboBox
    $interface_font_label                       = New-Object system.Windows.Forms.Label


    $color_form.FormBorderStyle = 'Fixed3D'
    $color_form.AutoScroll = $true
    $color_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $color_form.Location = new-object System.Drawing.Point(0, 0)
    $color_form.MaximizeBox = $false
    $color_form.SizeGripStyle = "Hide"
    $color_form.Width = 850
    $color_form.Height = 900
    $color_form.text = "Theme Settings";

    $y_pos = 10;

    
    $title_label.text                     = "Theme Settings";
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($color_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $color_form.controls.Add($title_label);

    $y_pos = $y_pos + $title_label.height + $v_spacer
    
    $theme_combo.Items.Clear();
    $theme_combo.width = 180
    $theme_combo.autosize = $false
    $theme_combo.Anchor = 'top,right'
    $theme_combo.Location = New-Object System.Drawing.Point((($color_form.width / 2) - $theme_combo.width),($y_pos + 3))
    $theme_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $theme_combo.DropDownStyle = "DropDownList"
    $theme_combo.AccessibleName = ""; 

    $themes = Get-ChildItem -Path "$dir\Resources\Themes" -File -Force -Filter *.csv
    foreach($theme in $themes)
    {
        $theme = [System.IO.Path]::GetFileNameWithoutExtension($theme)
        $theme_combo.Items.Add("$theme"); 
    }
    $theme_combo.SelectedItem = $settings['THEME'];
    $theme_combo.Add_SelectedValueChanged({
        $found = 1;
        foreach($og in $script:theme_settings.GetEnumerator())
        {
            [string]$og1 = $script:theme_original[$og.key]
            [string]$og2 = $og.value
            if($og1 -ne $og2)
            {
                $found = 0;
            }
        }
        if($found -eq 1)
        { 
            load_theme $this.SelectedItem
            $settings['THEME'] = $this.SelectedItem
            update_settings
        }
        else
        {
            #Changes Found
            $message = "Loading a new theme will clear your changes. Are you sure you want to continue?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                load_theme $this.SelectedItem
                $settings['THEME'] = $this.SelectedItem 
                update_settings  
            }
        }        
    })



    $color_form.controls.Add($theme_combo);

    
    $manage_theme_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $manage_theme_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $manage_theme_button.Width     = 150
    #$manage_theme_button.autosize = $true
    $manage_theme_button.height     = 30
    $manage_theme_button.Location  = New-Object System.Drawing.Point(($theme_combo.location.x + $theme_combo.width + 5),$y_pos);
    $manage_theme_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $manage_theme_button.Text      ="Manage Themes"
    $manage_theme_button.Name = ""
    $manage_theme_button.Add_Click({
        manage_themes
    })
    $color_form.controls.Add($manage_theme_button)

    $y_pos = $y_pos + 35
    
    $separator_bar1.text                        = ""
    $separator_bar1.AutoSize                    = $false
    $separator_bar1.BorderStyle                 = "fixed3d"
    #$separator_bar1.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
    $separator_bar1.Anchor                      = 'top,left'
    $separator_bar1.width                       = (($color_form.width - 50) - $spacer)
    $separator_bar1.height                      = 1
    $separator_bar1.location                    = New-Object System.Drawing.Point(20,$y_pos)
    $separator_bar1.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $separator_bar1.TextAlign                   = 'MiddleLeft'
    $color_form.controls.Add($separator_bar1);

    $y_pos = $y_pos + 5

    
    $header_label1.text                     = "General";
    $header_label1.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label1.Anchor                   = 'top,right'
    $header_label1.width                    = ($color_form.width / 4)
    $header_label1.height                   = 30
    $header_label1.TextAlign = "MiddleCenter"
    $header_label1.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label1.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label1);

    ######################################################################################################################
    #MainBackground

    $y_pos = $y_pos + 30 
    
    $main_background_color_label.text                     = "Main Background Color:";
    $main_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $main_background_color_label.Anchor                   = 'top,right'
    $main_background_color_label.TextAlign = "MiddleRight"
    $main_background_color_label.width                    = 350
    $main_background_color_label.height                   = 30
    $main_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $main_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($main_background_color_label);

    
    $main_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'],([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))                            
    $main_background_color_input.AutoSize                 = $false
    $main_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $main_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $main_background_color_input.Anchor                   = 'top,left'
    $main_background_color_input.width                    = 160
    $main_background_color_input.height                   = 27
    $main_background_color_input.location                 = New-Object System.Drawing.Point(($main_background_color_label.location.x + $main_background_color_label.width + 5),$y_pos)  
    $main_background_color_input.text                     = $script:theme_settings['MAIN_BACKGROUND_COLOR'] -replace '#',''
    $main_background_color_input.name                     = $script:theme_settings['MAIN_BACKGROUND_COLOR'] -replace '#',''

    $main_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['MAIN_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['MAIN_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $script:Form.BackColor = $script:theme_settings['MAIN_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($main_background_color_input);

    
    $main_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $main_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    #$main_background_color_button.autosize = $true
    #$main_background_color_button.AutoSizeMode = 'GrowAndShrink'
    $main_background_color_button.Width     = 100
    $main_background_color_button.height     = 27
    $main_background_color_button.Location  = New-Object System.Drawing.Point(($main_background_color_input.Location.x + $main_background_color_input.width + 5),$y_pos);
    $main_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $main_background_color_button.Text      ="Pick Color"
    $main_background_color_button.Name      = ""
    $main_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['MAIN_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['MAIN_BACKGROUND_COLOR'] = "#$color"
            $main_background_color_input.text = $color
            $main_background_color_input.name = $color
            $script:Form.BackColor = $script:theme_settings['MAIN_BACKGROUND_COLOR']
        }
        else
        {
            $script:theme_settings['MAIN_BACKGROUND_COLOR'] = "$color"
            $main_background_color_input.text = $color
            $main_background_color_input.name = $color
            $script:Form.BackColor = $script:theme_settings['MAIN_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($main_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Menu Text Color
    $y_pos = $y_pos + 30 
    
    $menu_text_color_label.text                     = "Menu Text Color:";
    $menu_text_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $menu_text_color_label.Anchor                   = 'top,right'
    $menu_text_color_label.TextAlign = "MiddleRight"
    $menu_text_color_label.width                    = 350
    $menu_text_color_label.height                   = 30
    $menu_text_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $menu_text_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($menu_text_color_label);

                       
    $menu_text_color_input.AutoSize                 = $false
    $menu_text_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $menu_text_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $menu_text_color_input.Anchor                   = 'top,left'
    $menu_text_color_input.width                    = 160
    $menu_text_color_input.height                   = 27
    $menu_text_color_input.location                 = New-Object System.Drawing.Point(($menu_text_color_label.location.x + $menu_text_color_label.width + 5),$y_pos)
    $menu_text_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $menu_text_color_input.text                     = $script:theme_settings['MENU_TEXT_COLOR'] -replace '#',''
    $menu_text_color_input.name                     = $script:theme_settings['MENU_TEXT_COLOR'] -replace '#',''
    $menu_text_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")

        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['MENU_TEXT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['MENU_TEXT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
               #$script:theme_settings['MENU_TEXT_COLOR'] = $this.text
               $MenuBar.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $FileMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $EditMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $OptionsMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $AboutMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $BulletMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               $script:AcronymMenu.Forecolor = $script:theme_settings['MENU_TEXT_COLOR']
               
        }
    })
    $color_form.controls.Add($menu_text_color_input);

    
    $menu_text_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $menu_text_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $menu_text_color_button.Width     = 100
    $menu_text_color_button.height     = 27
    $menu_text_color_button.Location  = New-Object System.Drawing.Point(($menu_text_color_input.Location.x + $menu_text_color_input.width + 5),$y_pos);
    $menu_text_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $menu_text_color_button.Text      ="Pick Color"
    $menu_text_color_button.Name      = ""
    $menu_text_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['MENU_TEXT_COLOR']
        ($trash,$color) = color_picker
        if($color -like '[a-f0-9]*')
        {
            $script:theme_settings['MENU_TEXT_COLOR'] = "#$color"
            $MenuBar.Forecolor = "#$color"
            $FileMenu.Forecolor = "#$color"
            $EditMenu.Forecolor = "#$color"
            $OptionsMenu.Forecolor = "#$color"
            $AboutMenu.Forecolor = "#$color"
            $BulletMenu.Forecolor = "#$color"
            $script:AcronymMenu.Forecolor = "#$color"
            $menu_text_color_input.text = $color
            $menu_text_color_input.name = $color
        }
        else
        {
            $script:theme_settings['MENU_TEXT_COLOR'] = "$color"
            $MenuBar.Forecolor = "$color"
            $FileMenu.Forecolor = "$color"
            $EditMenu.Forecolor = "$color"
            $OptionsMenu.Forecolor = "$color"
            $AboutMenu.Forecolor = "$color"
            $BulletMenu.Forecolor = "$color"
            $script:AcronymMenu.Forecolor = "$color"
            $menu_text_color_input.text = $color
            $menu_text_color_input.name = $color
        }
    })
    $color_form.controls.Add($menu_text_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Menu Background Color
    $y_pos = $y_pos + 30 
    
    $menu_background_color_label.text                     = "Menu Background Color:";
    $menu_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $menu_background_color_label.Anchor                   = 'top,right'
    $menu_background_color_label.TextAlign = "MiddleRight"
    $menu_background_color_label.width                    = 350
    $menu_background_color_label.height                   = 30
    $menu_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $menu_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($menu_background_color_label);

                           
    $menu_background_color_input.AutoSize                 = $false
    $menu_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $menu_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $menu_background_color_input.Anchor                   = 'top,left'
    $menu_background_color_input.width                    = 160
    $menu_background_color_input.height                   = 27
    $menu_background_color_input.location                 = New-Object System.Drawing.Point(($menu_background_color_label.location.x + $menu_background_color_label.width + 5),$y_pos)
    $menu_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $menu_background_color_input.text                     = $script:theme_settings['MENU_BACKGROUND_COLOR'] -replace '#',''
    $menu_background_color_input.name                     = $script:theme_settings['MENU_BACKGROUND_COLOR'] -replace '#',''
    $menu_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['MENU_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['MENU_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            #$script:theme_settings['MENU_BACKGROUND_COLOR'] = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $MenuBar.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $FileMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $EditMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $OptionsMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $AboutMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $BulletMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
            $script:AcronymMenu.Backcolor = $script:theme_settings['MENU_BACKGROUND_COLOR']
               
        }
    })
    $color_form.controls.Add($menu_background_color_input);

    
    $menu_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $menu_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $menu_background_color_button.Width     = 100
    $menu_background_color_button.height     = 27
    $menu_background_color_button.Location  = New-Object System.Drawing.Point(($menu_background_color_input.Location.x + $menu_background_color_input.width + 5),$y_pos);
    $menu_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $menu_background_color_button.Text      ="Pick Color"
    $menu_background_color_button.Name      = ""
    $menu_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['MENU_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        if($color -match "\d")
        {
            $script:theme_settings['MENU_BACKGROUND_COLOR'] = "#$color"
            $MenuBar.Backcolor = "#$color"
            $FileMenu.Backcolor = "#$color"
            $EditMenu.Backcolor = "#$color"
            $OptionsMenu.Backcolor = "#$color"
            $AboutMenu.Backcolor = "#$color"
            $BulletMenu.Backcolor = "#$color"
            $script:AcronymMenu.Backcolor = "#$color"
            $menu_background_color_input.text = "$color"
            $menu_background_color_input.name = "$color"
        }
        else
        {
            $script:theme_settings['MENU_BACKGROUND_COLOR'] = $color
            $MenuBar.Backcolor = "$color"
            $FileMenu.Backcolor = "$color"
            $EditMenu.Backcolor = "$color"
            $OptionsMenu.Backcolor = "$color"
            $AboutMenu.Backcolor = "$color"
            $BulletMenu.Backcolor = "$color"
            $script:AcronymMenu.Backcolor = "$color"
            $menu_background_color_input.text = $color
            $menu_background_color_input.name = $color
        }
    })
    $color_form.controls.Add($menu_background_color_button);

    ######################################################################################################################
    ######################################################################################################################
    #Adjustment Bars
    $y_pos = $y_pos + 30 
    
    $adjustment_bars_color_label.text                     = "Ajustment Bar Color:";
    $adjustment_bars_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $adjustment_bars_color_label.Anchor                   = 'top,right'
    $adjustment_bars_color_label.TextAlign = "MiddleRight"
    $adjustment_bars_color_label.width                    = 350
    $adjustment_bars_color_label.height                   = 30
    $adjustment_bars_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $adjustment_bars_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($adjustment_bars_color_label);

                         
    $adjustment_bars_color_input.AutoSize                 = $false
    $adjustment_bars_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $adjustment_bars_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $adjustment_bars_color_input.Anchor                   = 'top,left'
    $adjustment_bars_color_input.width                    = 160
    $adjustment_bars_color_input.height                   = 27
    $adjustment_bars_color_input.location                 = New-Object System.Drawing.Point(($adjustment_bars_color_label.location.x + $adjustment_bars_color_label.width + 5),$y_pos)
    $adjustment_bars_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $adjustment_bars_color_input.text                     = $script:theme_settings['ADJUSTMENT_BAR_COLOR'] -replace '#',''
    $adjustment_bars_color_input.name                     = $script:theme_settings['ADJUSTMENT_BAR_COLOR'] -replace '#',''
    $adjustment_bars_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['ADJUSTMENT_BAR_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['ADJUSTMENT_BAR_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $sidekick_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
            $bullet_feeder_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        }
    })
    $color_form.controls.Add($adjustment_bars_color_input);

    
    $adjustment_bars_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $adjustment_bars_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $adjustment_bars_color_button.Width     = 100
    $adjustment_bars_color_button.height     = 27
    $adjustment_bars_color_button.Location  = New-Object System.Drawing.Point(($adjustment_bars_color_input.Location.x + $adjustment_bars_color_input.width + 5),$y_pos);
    $adjustment_bars_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $adjustment_bars_color_button.Text      ="Pick Color"
    $adjustment_bars_color_button.Name      = ""
    $adjustment_bars_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['ADJUSTMENT_BAR_COLOR'] = "#$color"
            $adjustment_bars_color_input.text = $color
            $adjustment_bars_color_input.name = $color
            $sidekick_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
            $bullet_feeder_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        }
        else
        {
            $script:theme_settings['ADJUSTMENT_BAR_COLOR'] = "$color"
            $adjustment_bars_color_input.text = $color
            $adjustment_bars_color_input.name = $color
            $sidekick_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
            $bullet_feeder_panel.BackColor = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        }
    })
    $color_form.controls.Add($adjustment_bars_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Interface Font Type/Size
    $y_pos = $y_pos + 30
    
    $interface_font_label.text                     = "Interface Font:";
    $interface_font_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $interface_font_label.Anchor                   = 'top,right'
    $interface_font_label.TextAlign                = "MiddleRight"
    $interface_font_label.width                    = 350
    $interface_font_label.autosize                 = $false
    $interface_font_label.height                   = 30
    $interface_font_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $interface_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($interface_font_label);


    
    $interface_font_combo.Items.Clear();
    $interface_font_combo.width = 160
    $interface_font_combo.Anchor = 'top,right'
    $interface_font_combo.Autosize = $false
    $interface_font_combo.location                 = New-Object System.Drawing.Point(($interface_font_label.location.x + $interface_font_label.width + 25), ($y_pos + 3))
    $interface_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $interface_font_combo.DropDownStyle = "DropDownList"
    $interface_font_combo.AccessibleName = ""; 
    $fonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families
    foreach($font in $fonts)
    {
        $interface_font_combo.Items.Add($font.name); 
    }
    $interface_font_combo.SelectedItem = $script:theme_settings['INTERFACE_FONT']
    $interface_font_combo.Add_SelectedValueChanged({
       
        $script:theme_settings['INTERFACE_FONT'] = $this.SelectedItem

        $script:sidekickgui -eq "New"
        sidekick_display
        $title_label.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
        $theme_combo.font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $manage_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar1.Font                           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $save_theme_button.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $cancel_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar2.Font                           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $header_label1.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $main_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $main_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $main_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $menu_text_color_label.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_text_color_input.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_text_color_button.Font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $adjustment_bars_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label2.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $editor_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_misspelled_font_color_label.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_misspelled_font_color_input.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_misspelled_font_color_button.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_extend_acronym_font_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_extend_acronym_font_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_extend_acronym_font_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $editor_shorten_acronym_font_color_label.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_shorten_acronym_font_color_input.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_shorten_acronym_font_color_button.Font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_label.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_input.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_highlight_color_button.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $editor_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $editor_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label3.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $text_caclulator_background_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_background_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_background_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_label.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_input.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_under_color_button.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $text_caclulator_over_color_label.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_over_color_input.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_over_color_button.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label4.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $feeder_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $feeder_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $feeder_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $feeder_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label5.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $sidekick_background_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $sidekick_background_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $sidekick_background_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label6.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))   
        $dialog_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $dialog_title_font_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_font_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_font_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_title_banner_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_banner_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_banner_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_sub_header_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_text_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_text_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_text_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_background_color_label.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_background_color_input.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_background_color_button.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_button_text_color_label.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_text_color_input.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_text_color_button.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_label.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_input.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_background_color_button.Font    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $interface_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $interface_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $interface_font_size_combo.font                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $FileMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $EditMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $BulletMenu.Font                               = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $script:AcronymMenu.Font                       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $OptionsMenu.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $AboutMenu.Font                                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $script:sidekickgui = "New"
        update_sidekick

    })
    $color_form.controls.Add($interface_font_combo)

    
    $interface_font_size_combo.Items.Clear();
    $interface_font_size_combo.width = 80
    $interface_font_size_combo.Anchor = 'top,right'
    $interface_font_size_combo.location                    = New-Object System.Drawing.Point(($interface_font_combo.location.x + $interface_font_combo.width + 5), ($y_pos + 3))
    $interface_font_size_combo.autosize = $false
    $interface_font_size_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $interface_font_size_combo.DropDownStyle = "DropDownList"
    $interface_font_size_combo.AccessibleName = ""; 
    $counter = 2
    While($counter -le 20)
    {
        $counter = $counter + 0.5;
        $interface_font_size_combo.Items.Add($counter);
        if($counter -eq [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        {
           $interface_font_size_combo.SelectedItem = $counter
        }
    }
    $interface_font_size_combo.Add_SelectedValueChanged({
        [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] = $this.SelectedItem

        $title_label.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
        $theme_combo.font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $manage_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar2.Font                           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $save_theme_button.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $cancel_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar1.Font                           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $header_label1.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $main_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $main_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $main_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $menu_text_color_label.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_text_color_input.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_text_color_button.Font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $adjustment_bars_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label2.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $editor_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_misspelled_font_color_label.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_misspelled_font_color_input.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_misspelled_font_color_button.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_extend_acronym_font_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_extend_acronym_font_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_extend_acronym_font_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $editor_shorten_acronym_font_color_label.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_shorten_acronym_font_color_input.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_shorten_acronym_font_color_button.Font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_label.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_input.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_highlight_color_button.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $editor_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label3.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $text_caclulator_background_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_background_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_background_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_label.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_input.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_under_color_button.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $text_caclulator_over_color_label.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_over_color_input.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_over_color_button.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label4.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $feeder_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $feeder_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $feeder_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $feeder_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label5.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $sidekick_background_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $sidekick_background_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $sidekick_background_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label6.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))   
        $dialog_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $dialog_title_font_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_font_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_font_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_title_banner_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_banner_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_banner_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_sub_header_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_text_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_text_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_text_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_background_color_label.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_background_color_input.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_background_color_button.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_button_text_color_label.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_text_color_input.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_text_color_button.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_label.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_input.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_background_color_button.Font    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $interface_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $interface_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $interface_font_size_combo.font                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $FileMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $EditMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $BulletMenu.Font                               = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $script:AcronymMenu.Font                       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $OptionsMenu.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $AboutMenu.Font                                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $script:sidekickgui = "New"
        update_sidekick
    })
    $color_form.controls.Add($interface_font_size_combo)

    ######################################################################################################################
    ######################################################################################################################
    ##Editor Header
    $y_pos = $y_pos + 35
    
    $header_label2.text                     = "Editor";
    $header_label2.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label2.Anchor                   = 'top,right'
    $header_label2.width                    = ($color_form.width / 4)
    $header_label2.height                   = 30
    $header_label2.TextAlign = "MiddleCenter"
    $header_label2.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label2.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label2);

    ######################################################################################################################
    #Editor Background Color
    $y_pos = $y_pos + 30 
    
    $editor_background_color_label.text                     = "Background Color:";
    $editor_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_background_color_label.Anchor                   = 'top,right'
    $editor_background_color_label.TextAlign = "MiddleRight"
    $editor_background_color_label.width                    = 350
    $editor_background_color_label.height                   = 30
    $editor_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_background_color_label);

                        
    $editor_background_color_input.AutoSize                 = $false
    $editor_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_background_color_input.Anchor                   = 'top,left'
    $editor_background_color_input.width                    = 160
    $editor_background_color_input.height                   = 27
    $editor_background_color_input.location                 = New-Object System.Drawing.Point(($editor_background_color_label.location.x + $editor_background_color_label.width + 5),$y_pos)
    $editor_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_background_color_input.text                     = $script:theme_settings['EDITOR_BACKGROUND_COLOR'] -replace '#',''
    $editor_background_color_input.name                     = $script:theme_settings['EDITOR_BACKGROUND_COLOR'] -replace '#',''
    $editor_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $editor.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_background_color_input);

    
    $editor_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_background_color_button.Width     = 100
    $editor_background_color_button.height     = 27
    $editor_background_color_button.Location  = New-Object System.Drawing.Point(($editor_background_color_input.Location.x + $editor_background_color_input.width + 5),$y_pos);
    $editor_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_background_color_button.Text      ="Pick Color"
    $editor_background_color_button.Name      = ""
    $editor_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_BACKGROUND_COLOR'] = "#$color"
            $editor_background_color_input.text = $color
            $editor_background_color_input.name = $color
            $editor.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_BACKGROUND_COLOR'] = "$color"
            $editor_background_color_input.text = $color
            $editor_background_color_input.name = $color
            $editor.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Editor Font Color
    $y_pos = $y_pos + 30 
    
    $editor_font_color_label.text                     = "Font Color:";
    $editor_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_font_color_label.Anchor                   = 'top,right'
    $editor_font_color_label.TextAlign = "MiddleRight"
    $editor_font_color_label.width                    = 350
    $editor_font_color_label.height                   = 30
    $editor_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_font_color_label);

                          
    $editor_font_color_input.AutoSize                 = $false
    $editor_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_font_color_input.Anchor                   = 'top,left'
    $editor_font_color_input.width                    = 160
    $editor_font_color_input.height                   = 27
    $editor_font_color_input.location                 = New-Object System.Drawing.Point(($editor_font_color_label.location.x + $editor_font_color_label.width + 5),$y_pos)
    $editor_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_font_color_input.text                     = $script:theme_settings['EDITOR_FONT_COLOR'] -replace '#',''
    $editor_font_color_input.name                     = $script:theme_settings['EDITOR_FONT_COLOR'] -replace '#',''
    $editor_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_font_color_input);

    
    $editor_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_font_color_button.Width     = 100
    $editor_font_color_button.height     = 27
    $editor_font_color_button.Location  = New-Object System.Drawing.Point(($editor_font_color_input.Location.x + $editor_font_color_input.width + 5),$y_pos);
    $editor_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_font_color_button.Text      ="Pick Color"
    $editor_font_color_button.Name      = ""
    $editor_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_FONT_COLOR'] = "#$color"
            $editor_font_color_input.text = $color
            $editor_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_FONT_COLOR'] = "$color"
            $editor_font_color_input.text = $color
            $editor_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Editor Misspelled Font Color
    $y_pos = $y_pos + 30 
    
    $editor_misspelled_font_color_label.text                     = "Mispelled Font Color:";
    $editor_misspelled_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_misspelled_font_color_label.Anchor                   = 'top,right'
    $editor_misspelled_font_color_label.TextAlign = "MiddleRight"
    $editor_misspelled_font_color_label.width                    = 350
    $editor_misspelled_font_color_label.height                   = 30
    $editor_misspelled_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_misspelled_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_misspelled_font_color_label);

                         
    $editor_misspelled_font_color_input.AutoSize                 = $false
    $editor_misspelled_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_misspelled_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_misspelled_font_color_input.Anchor                   = 'top,left'
    $editor_misspelled_font_color_input.width                    = 160
    $editor_misspelled_font_color_input.height                   = 27
    $editor_misspelled_font_color_input.location                 = New-Object System.Drawing.Point(($editor_misspelled_font_color_label.location.x + $editor_misspelled_font_color_label.width + 5),$y_pos)
    $editor_misspelled_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_misspelled_font_color_input.text                     = $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] -replace '#',''
    $editor_misspelled_font_color_input.name                     = $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] -replace '#',''
    $editor_misspelled_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_misspelled_font_color_input);

    
    $editor_misspelled_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_misspelled_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_misspelled_font_color_button.Width     = 100
    $editor_misspelled_font_color_button.height     = 27
    $editor_misspelled_font_color_button.Location  = New-Object System.Drawing.Point(($editor_misspelled_font_color_input.Location.x + $editor_misspelled_font_color_input.width + 5),$y_pos);
    $editor_misspelled_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_misspelled_font_color_button.Text      ="Pick Color"
    $editor_misspelled_font_color_button.Name      = ""
    $editor_misspelled_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] = "#$color"
            $editor_misspelled_font_color_input.text = $color
            $editor_misspelled_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] = "$color"
            $editor_misspelled_font_color_input.text = $color
            $editor_misspelled_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_misspelled_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Editor Extend Acronym Font Color
    $y_pos = $y_pos + 30 
    
    $editor_extend_acronym_font_color_label.text                     = "Extend Acronym Font Color:";
    $editor_extend_acronym_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_extend_acronym_font_color_label.Anchor                   = 'top,right'
    $editor_extend_acronym_font_color_label.TextAlign = "MiddleRight"
    $editor_extend_acronym_font_color_label.width                    = 350
    $editor_extend_acronym_font_color_label.height                   = 30
    $editor_extend_acronym_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_extend_acronym_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_extend_acronym_font_color_label);

                         
    $editor_extend_acronym_font_color_input.AutoSize                 = $false
    $editor_extend_acronym_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_extend_acronym_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_extend_acronym_font_color_input.Anchor                   = 'top,left'
    $editor_extend_acronym_font_color_input.width                    = 160
    $editor_extend_acronym_font_color_input.height                   = 27
    $editor_extend_acronym_font_color_input.location                 = New-Object System.Drawing.Point(($editor_extend_acronym_font_color_label.location.x + $editor_extend_acronym_font_color_label.width + 5),$y_pos)
    $editor_extend_acronym_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_extend_acronym_font_color_input.text                     = $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] -replace '#',''
    $editor_extend_acronym_font_color_input.name                     = $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] -replace '#',''
    $editor_extend_acronym_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_extend_acronym_font_color_input);

    
    $editor_extend_acronym_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_extend_acronym_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_extend_acronym_font_color_button.Width     = 100
    $editor_extend_acronym_font_color_button.height     = 27
    $editor_extend_acronym_font_color_button.Location  = New-Object System.Drawing.Point(($editor_extend_acronym_font_color_input.Location.x + $editor_extend_acronym_font_color_input.width + 5),$y_pos);
    $editor_extend_acronym_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_extend_acronym_font_color_button.Text      ="Pick Color"
    $editor_extend_acronym_font_color_button.Name      = ""
    $editor_extend_acronym_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] = "#$color"
            $editor_extend_acronym_font_color_input.text = $color
            $editor_extend_acronym_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] = "$color"
            $editor_extend_acronym_font_color_input.text = $color
            $editor_extend_acronym_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_extend_acronym_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Editor Shorten Acronym Font Color
    $y_pos = $y_pos + 30 
    
    $editor_shorten_acronym_font_color_label.text                     = "Shorten Acronym Font Color:";
    $editor_shorten_acronym_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_shorten_acronym_font_color_label.Anchor                   = 'top,right'
    $editor_shorten_acronym_font_color_label.TextAlign = "MiddleRight"
    $editor_shorten_acronym_font_color_label.width                    = 350
    $editor_shorten_acronym_font_color_label.height                   = 30
    $editor_shorten_acronym_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_shorten_acronym_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_shorten_acronym_font_color_label);

                         
    $editor_shorten_acronym_font_color_input.AutoSize                 = $false
    $editor_shorten_acronym_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_shorten_acronym_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_shorten_acronym_font_color_input.Anchor                   = 'top,left'
    $editor_shorten_acronym_font_color_input.width                    = 160
    $editor_shorten_acronym_font_color_input.height                   = 27
    $editor_shorten_acronym_font_color_input.location                 = New-Object System.Drawing.Point(($editor_shorten_acronym_font_color_label.location.x + $editor_shorten_acronym_font_color_label.width + 5),$y_pos)
    $editor_shorten_acronym_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_shorten_acronym_font_color_input.text                     = $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] -replace '#',''
    $editor_shorten_acronym_font_color_input.name                     = $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] -replace '#',''
    $editor_shorten_acronym_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_shorten_acronym_font_color_input);

    
    $editor_shorten_acronym_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_shorten_acronym_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_shorten_acronym_font_color_button.Width     = 100
    $editor_shorten_acronym_font_color_button.height     = 27
    $editor_shorten_acronym_font_color_button.Location  = New-Object System.Drawing.Point(($editor_shorten_acronym_font_color_input.Location.x + $editor_shorten_acronym_font_color_input.width + 5),$y_pos);
    $editor_shorten_acronym_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_shorten_acronym_font_color_button.Text      ="Pick Color"
    $editor_shorten_acronym_font_color_button.Name      = ""
    $editor_shorten_acronym_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] = "#$color"
            $editor_shorten_acronym_font_color_input.text = $color
            $editor_shorten_acronym_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] = "$color"
            $editor_shorten_acronym_font_color_input.text = $color
            $editor_shorten_acronym_font_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_shorten_acronym_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Editor Highlight Color
    $y_pos = $y_pos + 30 
    
    $editor_highlight_color_label.text                     = "Highlight Color:";
    $editor_highlight_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_highlight_color_label.Anchor                   = 'top,right'
    $editor_highlight_color_label.TextAlign = "MiddleRight"
    $editor_highlight_color_label.width                    = 350
    $editor_highlight_color_label.height                   = 30
    $editor_highlight_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_highlight_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_highlight_color_label);

                         
    $editor_highlight_color_input.AutoSize                 = $false
    $editor_highlight_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $editor_highlight_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $editor_highlight_color_input.Anchor                   = 'top,left'
    $editor_highlight_color_input.width                    = 160
    $editor_highlight_color_input.height                   = 27
    $editor_highlight_color_input.location                 = New-Object System.Drawing.Point(($editor_highlight_color_label.location.x + $editor_highlight_color_label.width + 5),$y_pos)
    $editor_highlight_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $editor_highlight_color_input.text                     = $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] -replace '#',''
    $editor_highlight_color_input.name                     = $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] -replace '#',''
    $editor_highlight_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_highlight_color_input);

    
    $editor_highlight_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $editor_highlight_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $editor_highlight_color_button.Width     = 100
    $editor_highlight_color_button.height     = 27
    $editor_highlight_color_button.Location  = New-Object System.Drawing.Point(($editor_highlight_color_input.Location.x + $editor_highlight_color_input.width + 5),$y_pos);
    $editor_highlight_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $editor_highlight_color_button.Text      ="Pick Color"
    $editor_highlight_color_button.Name      = ""
    $editor_highlight_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['EDITOR_HIGHLIGHT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] = "#$color"
            $editor_highlight_color_input.text = $color
            $editor_highlight_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] = "$color"
            $editor_highlight_color_input.text = $color
            $editor_highlight_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($editor_highlight_color_button);
    ######################################################################################################################
    ######################################################################################################################
    ##Editor Font

    $y_pos = $y_pos + 30 
    
    $editor_font_label.text                     = "Editor Font:";
    $editor_font_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $editor_font_label.Anchor                   = 'top,right'
    $editor_font_label.TextAlign = "MiddleRight"
    $editor_font_label.width                    = 350
    $editor_font_label.height                   = 30
    $editor_font_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $editor_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($editor_font_label);
  
    $editor_font_combo.Items.Clear();
    $editor_font_combo.width = 180
    $editor_font_combo.Anchor = 'top,right'
    $editor_font_combo.autosize = $false
    $editor_font_combo.location                 = New-Object System.Drawing.Point(($editor_font_label.location.x + $editor_font_label.width + 25), ($y_pos + 3))
    $editor_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $editor_font_combo.DropDownStyle = "DropDownList"
    $editor_font_combo.AccessibleName = ""; 
    $fonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families
    foreach($font in $fonts)
    {
        $editor_font_combo.Items.Add($font.name); 
    }
    $editor_font_combo.SelectedItem = $script:theme_settings['EDITOR_FONT']
    $editor_font_combo.Add_SelectedValueChanged({
       
        $script:theme_settings['EDITOR_FONT'] = $this.SelectedItem
        $editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        $sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])    
        scan_text
        $script:Form.refresh();
        $sizer_art.refresh();
        $Script:recent_editor_text = "Changed"

    })
    $color_form.controls.Add($editor_font_combo)

    
    $editor_font_size_combo.Items.Clear();
    $editor_font_size_combo.width = 80
    $editor_font_size_combo.autosize = $false
    $editor_font_size_combo.Anchor = 'top,right'
    $editor_font_size_combo.location   = New-Object System.Drawing.Point(($editor_font_combo.location.x + $editor_font_combo.width + 5), ($y_pos + 3))
    $editor_font_size_combo.font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $editor_font_size_combo.DropDownStyle = "DropDownList"
    $editor_font_size_combo.AccessibleName = ""; 
    $counter = 8
    While($counter -le 20)
    {
        $counter = $counter + 0.5;
        $editor_font_size_combo.Items.Add($counter);
        if($counter -eq [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        {
           $editor_font_size_combo.SelectedItem = $counter
        }
    }
    $editor_font_size_combo.Add_SelectedValueChanged({
        [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'] = $this.SelectedItem
        $editor.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        $sizer_box.Font = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])    
        scan_text
        $script:Form.refresh();
        $sizer_art.refresh();
        $Script:recent_editor_text = "Changed"
    })
    $color_form.controls.Add($editor_font_size_combo)

    ######################################################################################################################
    ######################################################################################################################
    ##Text Size Calculator Header
    $y_pos = $y_pos + 35
    
    $header_label3.text                     = "Text Size Calculator";
    $header_label3.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label3.Anchor                   = 'top,right'
    $header_label3.width                    = ($color_form.width / 4)
    $header_label3.height                   = 30
    $header_label3.TextAlign = "MiddleCenter"
    $header_label3.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label3.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label3);

    ######################################################################################################################
    #Text Size Calculator Background Color
    $y_pos = $y_pos + 30 
    
    $text_caclulator_background_color_label.text                     = "Background Color:";
    $text_caclulator_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $text_caclulator_background_color_label.Anchor                   = 'top,right'
    $text_caclulator_background_color_label.TextAlign = "MiddleRight"
    $text_caclulator_background_color_label.width                    = 350
    $text_caclulator_background_color_label.height                   = 30
    $text_caclulator_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $text_caclulator_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($text_caclulator_background_color_label);

                         
    $text_caclulator_background_color_input.AutoSize                 = $false
    $text_caclulator_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $text_caclulator_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $text_caclulator_background_color_input.Anchor                   = 'top,left'
    $text_caclulator_background_color_input.width                    = 160
    $text_caclulator_background_color_input.height                   = 27
    $text_caclulator_background_color_input.location                 = New-Object System.Drawing.Point(($text_caclulator_background_color_label.location.x + $text_caclulator_background_color_label.width + 5),$y_pos)
    $text_caclulator_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $text_caclulator_background_color_input.text                     = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] -replace '#',''
    $text_caclulator_background_color_input.name                     = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] -replace '#',''
    $text_caclulator_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $sizer_box.backcolor = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
            $sizer_art.refresh();
            $script:Form.Refresh();
        }
    })
    $color_form.controls.Add($text_caclulator_background_color_input);

    
    $text_caclulator_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $text_caclulator_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $text_caclulator_background_color_button.Width     = 100
    $text_caclulator_background_color_button.height     = 27
    $text_caclulator_background_color_button.Location  = New-Object System.Drawing.Point(($text_caclulator_background_color_input.Location.x + $text_caclulator_background_color_input.width + 5),$y_pos);
    $text_caclulator_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $text_caclulator_background_color_button.Text      ="Pick Color"
    $text_caclulator_background_color_button.Name      = ""
    $text_caclulator_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] = "#$color"
            $text_caclulator_background_color_input.text = $color
            $text_caclulator_background_color_input.name = $color
            $sizer_box.backcolor = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
            $sizer_art.refresh();
            $script:Form.Refresh();

        }
        else
        {
            $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] = "$color"
            $text_caclulator_background_color_input.text = $color
            $text_caclulator_background_color_input.name = $color
            $sizer_box.backcolor = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
            $sizer_art.refresh();
            $script:Form.Refresh();
        }
    })
    $color_form.controls.Add($text_caclulator_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Text Size Calculator Under Limit Font Color
    $y_pos = $y_pos + 30 
    
    $text_caclulator_under_color_label.text                     = "Under Limit Font Color:";
    $text_caclulator_under_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $text_caclulator_under_color_label.Anchor                   = 'top,right'
    $text_caclulator_under_color_label.TextAlign = "MiddleRight"
    $text_caclulator_under_color_label.width                    = 350
    $text_caclulator_under_color_label.height                   = 30
    $text_caclulator_under_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $text_caclulator_under_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($text_caclulator_under_color_label);

                         
    $text_caclulator_under_color_input.AutoSize                 = $false
    $text_caclulator_under_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $text_caclulator_under_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $text_caclulator_under_color_input.Anchor                   = 'top,left'
    $text_caclulator_under_color_input.width                    = 160
    $text_caclulator_under_color_input.height                   = 27
    $text_caclulator_under_color_input.location                 = New-Object System.Drawing.Point(($text_caclulator_under_color_label.location.x + $text_caclulator_under_color_label.width + 5),$y_pos)
    $text_caclulator_under_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $text_caclulator_under_color_input.text                     = $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] -replace '#',''
    $text_caclulator_under_color_input.name                     = $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] -replace '#',''
    $text_caclulator_under_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        $this.text
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($text_caclulator_under_color_input);

    
    $text_caclulator_under_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $text_caclulator_under_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $text_caclulator_under_color_button.Width     = 100
    $text_caclulator_under_color_button.height     = 27
    $text_caclulator_under_color_button.Location  = New-Object System.Drawing.Point(($text_caclulator_under_color_input.Location.x + $text_caclulator_under_color_input.width + 5),$y_pos);
    $text_caclulator_under_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $text_caclulator_under_color_button.Text      ="Pick Color"
    $text_caclulator_under_color_button.Name      = ""
    $text_caclulator_under_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] = "#$color"
            $text_caclulator_under_color_input.text = $color
            $text_caclulator_under_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] = "$color"
            $text_caclulator_under_color_input.text = $color
            $text_caclulator_under_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($text_caclulator_under_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Text Size Calculator Over Limit Color
    $y_pos = $y_pos + 30 
    
    $text_caclulator_over_color_label.text                     = "Over Limit Font Color:";
    $text_caclulator_over_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $text_caclulator_over_color_label.Anchor                   = 'top,right'
    $text_caclulator_over_color_label.TextAlign = "MiddleRight"
    $text_caclulator_over_color_label.width                    = 350
    $text_caclulator_over_color_label.height                   = 30
    $text_caclulator_over_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $text_caclulator_over_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($text_caclulator_over_color_label);

                          
    $text_caclulator_over_color_input.AutoSize                 = $false
    $text_caclulator_over_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $text_caclulator_over_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $text_caclulator_over_color_input.Anchor                   = 'top,left'
    $text_caclulator_over_color_input.width                    = 160
    $text_caclulator_over_color_input.height                   = 27
    $text_caclulator_over_color_input.location                 = New-Object System.Drawing.Point(($text_caclulator_over_color_label.location.x + $text_caclulator_over_color_label.width + 5),$y_pos)
    $text_caclulator_over_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $text_caclulator_over_color_input.text                     = $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] -replace '#',''
    $text_caclulator_over_color_input.name                     = $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] -replace '#',''
    $text_caclulator_over_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {   
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($text_caclulator_over_color_input);

    
    $text_caclulator_over_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $text_caclulator_over_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $text_caclulator_over_color_button.Width     = 100
    $text_caclulator_over_color_button.height     = 27
    $text_caclulator_over_color_button.Location  = New-Object System.Drawing.Point(($text_caclulator_over_color_input.Location.x + $text_caclulator_over_color_input.width + 5),$y_pos);
    $text_caclulator_over_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $text_caclulator_over_color_button.Text      ="Pick Color"
    $text_caclulator_over_color_button.Name      = ""
    $text_caclulator_over_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] = "#$color"
            $text_caclulator_over_color_input.text = $color
            $text_caclulator_over_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
        else
        {
            $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] = "$color"
            $text_caclulator_over_color_input.text = $color
            $text_caclulator_over_color_input.name = $color
            $Script:recent_editor_text = "Changed"
        }
    })
    $color_form.controls.Add($text_caclulator_over_color_button);
    ######################################################################################################################

    $y_pos = $y_pos + 35
    
    $header_label4.text                     = "Bullet Feed";
    $header_label4.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label4.Anchor                   = 'top,right'
    $header_label4.width                    = ($color_form.width / 4)
    $header_label4.height                   = 30
    $header_label4.TextAlign = "MiddleCenter"
    $header_label4.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label4.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label4);

    ######################################################################################################################
    #Feeder Background Color
    $y_pos = $y_pos + 30 
    
    $feeder_background_color_label.text                     = "Background Color:";
    $feeder_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $feeder_background_color_label.Anchor                   = 'top,right'
    $feeder_background_color_label.TextAlign = "MiddleRight"
    $feeder_background_color_label.width                    = 350
    $feeder_background_color_label.height                   = 30
    $feeder_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $feeder_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($feeder_background_color_label);

                          
    $feeder_background_color_input.AutoSize                 = $false
    $feeder_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $feeder_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $feeder_background_color_input.Anchor                   = 'top,left'
    $feeder_background_color_input.width                    = 160
    $feeder_background_color_input.height                   = 27
    $feeder_background_color_input.location                 = New-Object System.Drawing.Point(($feeder_background_color_label.location.x + $feeder_background_color_label.width + 5),$y_pos)
    $feeder_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $feeder_background_color_input.text                     = $script:theme_settings['FEEDER_BACKGROUND_COLOR'] -replace '#',''
    $feeder_background_color_input.name                     = $script:theme_settings['FEEDER_BACKGROUND_COLOR'] -replace '#',''
    $feeder_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['FEEDER_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['FEEDER_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $feeder_box.BackColor = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($feeder_background_color_input);

    
    $feeder_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $feeder_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $feeder_background_color_button.Width     = 100
    $feeder_background_color_button.height     = 27
    $feeder_background_color_button.Location  = New-Object System.Drawing.Point(($feeder_background_color_input.Location.x + $feeder_background_color_input.width + 5),$y_pos);
    $feeder_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $feeder_background_color_button.Text      ="Pick Color"
    $feeder_background_color_button.Name      = ""
    $feeder_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['FEEDER_BACKGROUND_COLOR'] = "#$color"
            $feeder_background_color_input.text = $color
            $feeder_background_color_input.name = $color
            $feeder_box.BackColor = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
        }
        else
        {
            $script:theme_settings['FEEDER_BACKGROUND_COLOR'] = "$color"
            $feeder_background_color_input.text = $color
            $feeder_background_color_input.name = $color
            $feeder_box.BackColor = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($feeder_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Bullet Feed Font Color
    $y_pos = $y_pos + 30 
    
    $feeder_font_color_label.text                     = "Font Color:";
    $feeder_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $feeder_font_color_label.Anchor                   = 'top,right'
    $feeder_font_color_label.TextAlign = "MiddleRight"
    $feeder_font_color_label.width                    = 350
    $feeder_font_color_label.height                   = 30
    $feeder_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $feeder_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($feeder_font_color_label);

                         
    $feeder_font_color_input.AutoSize                 = $false
    $feeder_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $feeder_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $feeder_font_color_input.Anchor                   = 'top,left'
    $feeder_font_color_input.width                    = 160
    $feeder_font_color_input.height                   = 27
    $feeder_font_color_input.location                 = New-Object System.Drawing.Point(($feeder_font_color_label.location.x + $feeder_font_color_label.width + 5),$y_pos)
    $feeder_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $feeder_font_color_input.text                     = $script:theme_settings['FEEDER_FONT_COLOR'] -replace '#',''
    $feeder_font_color_input.name                     = $script:theme_settings['FEEDER_FONT_COLOR'] -replace '#',''
    $feeder_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['FEEDER_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['FEEDER_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $feeder_box.ForeColor = $script:theme_settings['FEEDER_FONT_COLOR']
        }
    })
    $color_form.controls.Add($feeder_font_color_input);

    
    $feeder_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $feeder_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $feeder_font_color_button.Width     = 100
    $feeder_font_color_button.height     = 27
    $feeder_font_color_button.Location  = New-Object System.Drawing.Point(($feeder_font_color_input.Location.x + $feeder_font_color_input.width + 5),$y_pos);
    $feeder_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $feeder_font_color_button.Text      ="Pick Color"
    $feeder_font_color_button.Name      = ""
    $feeder_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['FEEDER_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['FEEDER_FONT_COLOR'] = "#$color"
            $feeder_font_color_input.text = $color
            $feeder_font_color_input.name = $color
            $feeder_box.ForeColor = $script:theme_settings['FEEDER_FONT_COLOR']
        }
        else
        {
            $script:theme_settings['FEEDER_FONT_COLOR'] = "$color"
            $feeder_font_color_input.text = $color
            $feeder_font_color_input.name = $color
            $feeder_box.ForeColor = $script:theme_settings['FEEDER_FONT_COLOR']
        }
    })
    $color_form.controls.Add($feeder_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    ##Feeder Font

    $y_pos = $y_pos + 30 
    
    $feeder_font_label.text                     = "Editor Font:";
    $feeder_font_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $feeder_font_label.Anchor                   = 'top,right'
    $feeder_font_label.TextAlign = "MiddleRight"
    $feeder_font_label.width                    = 350
    $feeder_font_label.height                   = 30
    $feeder_font_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $feeder_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($feeder_font_label);


    
    $feeder_font_combo.Items.Clear();
    $feeder_font_combo.width = 180
    $feeder_font_combo.autosize = $false
    $feeder_font_combo.Anchor = 'top,right'
    $feeder_font_combo.location                 = New-Object System.Drawing.Point(($feeder_font_label.location.x + $feeder_font_label.width + 25), ($y_pos + 3))
    $feeder_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $feeder_font_combo.DropDownStyle = "DropDownList"
    $feeder_font_combo.AccessibleName = ""; 
    $fonts = (New-Object System.Drawing.Text.InstalledFontCollection).Families
    foreach($font in $fonts)
    {
        $feeder_font_combo.Items.Add($font.name); 
    }
    $feeder_font_combo.SelectedItem = $script:theme_settings['FEEDER_FONT']
    $feeder_font_combo.Add_SelectedValueChanged({
       
        $script:theme_settings['FEEDER_FONT'] = $this.SelectedItem
        $feeder_box.text = "";
        $feeder_box.Font = [Drawing.Font]::New($script:theme_settings['FEEDER_FONT'], [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'])
        update_feeder
        $script:Form.refresh();
    })
    $color_form.controls.Add($feeder_font_combo)

    
    $feeder_font_size_combo.Items.Clear();
    $feeder_font_size_combo.autosize = $false
    $feeder_font_size_combo.width = 80
    $feeder_font_size_combo.Anchor = 'top,right'
    $feeder_font_size_combo.location   = New-Object System.Drawing.Point(($feeder_font_combo.location.x + $feeder_font_combo.width + 5), ($y_pos + 3))
    $feeder_font_size_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
    $feeder_font_size_combo.DropDownStyle = "DropDownList"
    $feeder_font_size_combo.AccessibleName = ""; 
    $counter = 8
    While($counter -le 20)
    {
        $counter = $counter + 0.5;
        $feeder_font_size_combo.Items.Add($counter);
        if($counter -eq [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'])
        {
           $feeder_font_size_combo.SelectedItem = $counter
        }
    }
    $feeder_font_size_combo.Add_SelectedValueChanged({
        [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'] = $this.SelectedItem
        $feeder_box.text = "";
        $feeder_box.Font = [Drawing.Font]::New($script:theme_settings['FEEDER_FONT'], [Decimal]$script:theme_settings['FEEDER_FONT_SIZE'])
        update_feeder
        $script:Form.refresh();
    })
    $color_form.controls.Add($feeder_font_size_combo)




    ######################################################################################################################

    $y_pos = $y_pos + 35
    
    $header_label5.text                     = "Sidekick Panel";
    $header_label5.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label5.Anchor                   = 'top,right'
    $header_label5.width                    = ($color_form.width / 4)
    $header_label5.height                   = 30
    $header_label5.TextAlign = "MiddleCenter"
    $header_label5.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label5.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label5);

    ######################################################################################################################
    #Sidekick Background Color
    $y_pos = $y_pos + 30 
    
    $sidekick_background_color_label.text                     = "Background Color:";
    $sidekick_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $sidekick_background_color_label.Anchor                   = 'top,right'
    $sidekick_background_color_label.TextAlign = "MiddleRight"
    $sidekick_background_color_label.width                    = 350
    $sidekick_background_color_label.height                   = 30
    $sidekick_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $sidekick_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($sidekick_background_color_label);

                          
    $sidekick_background_color_input.AutoSize                 = $false
    $sidekick_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $sidekick_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $sidekick_background_color_input.Anchor                   = 'top,left'
    $sidekick_background_color_input.width                    = 160
    $sidekick_background_color_input.height                   = 27
    $sidekick_background_color_input.location                 = New-Object System.Drawing.Point(($sidekick_background_color_label.location.x + $sidekick_background_color_label.width + 5),$y_pos)
    $sidekick_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $sidekick_background_color_input.text                     = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] -replace '#',''
    $sidekick_background_color_input.name                     = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] -replace '#',''
    $sidekick_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $left_panel.BackColor                = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($sidekick_background_color_input);

    
    $sidekick_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $sidekick_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $sidekick_background_color_button.Width     = 100
    $sidekick_background_color_button.height     = 27
    $sidekick_background_color_button.Location  = New-Object System.Drawing.Point(($sidekick_background_color_input.Location.x + $sidekick_background_color_input.width + 5),$y_pos);
    $sidekick_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $sidekick_background_color_button.Text      ="Pick Color"
    $sidekick_background_color_button.Name      = ""
    $sidekick_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] = "#$color"
            $sidekick_background_color_input.text = $color
            $sidekick_background_color_input.name = $color
            $left_panel.BackColor                = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
        }
        else
        {
            $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] = "$color"
            $sidekick_background_color_input.text = $color
            $sidekick_background_color_input.name = $color
            $left_panel.BackColor                = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($sidekick_background_color_button);
    ######################################################################################################################

    $y_pos = $y_pos + 35
    
    $header_label6.text                     = "Dialog Boxes";
    $header_label6.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $header_label6.Anchor                   = 'top,right'
    $header_label6.width                    = ($color_form.width / 4)
    $header_label6.height                   = 30
    $header_label6.TextAlign = "MiddleCenter"
    $header_label6.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $header_label6.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
    $color_form.controls.Add($header_label6);

    ######################################################################################################################
    #Dialog Background Color
    $y_pos = $y_pos + 30 
    
    $dialog_background_color_label.text                     = "Background Color:";
    $dialog_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_background_color_label.Anchor                   = 'top,right'
    $dialog_background_color_label.TextAlign = "MiddleRight"
    $dialog_background_color_label.width                    = 350
    $dialog_background_color_label.height                   = 30
    $dialog_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_background_color_label);

                        
    $dialog_background_color_input.AutoSize                 = $false
    $dialog_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_background_color_input.Anchor                   = 'top,left'
    $dialog_background_color_input.width                    = 160
    $dialog_background_color_input.height                   = 27
    $dialog_background_color_input.location                 = New-Object System.Drawing.Point(($dialog_background_color_label.location.x + $dialog_background_color_label.width + 5),$y_pos)
    $dialog_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_background_color_input.text                     = $script:theme_settings['DIALOG_BACKGROUND_COLOR'] -replace '#',''
    $dialog_background_color_input.name                     = $script:theme_settings['DIALOG_BACKGROUND_COLOR'] -replace '#',''
    $dialog_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $color_form.Backcolor = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($dialog_background_color_input);

    
    $dialog_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_background_color_button.Width     = 100
    $dialog_background_color_button.height     = 27
    $dialog_background_color_button.Location  = New-Object System.Drawing.Point(($dialog_background_color_input.Location.x + $dialog_background_color_input.width + 5),$y_pos);
    $dialog_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_background_color_button.Text      ="Pick Color"
    $dialog_background_color_button.Name      = ""
    $dialog_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_BACKGROUND_COLOR'] = "#$color"
            $dialog_background_color_input.text = $color
            $dialog_background_color_input.name = $color
            $color_form.Backcolor = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
        }
        else
        {
            $script:theme_settings['DIALOG_BACKGROUND_COLOR'] = "$color"
            $dialog_background_color_input.text = $color
            $dialog_background_color_input.name = $color
            $color_form.Backcolor = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
        }
    })
    $color_form.controls.Add($dialog_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Title Font Color
    $y_pos = $y_pos + 30 
    
    $dialog_title_font_color_label.text                     = "Title Font Color:";
    $dialog_title_font_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_title_font_color_label.Anchor                   = 'top,right'
    $dialog_title_font_color_label.TextAlign = "MiddleRight"
    $dialog_title_font_color_label.width                    = 350
    $dialog_title_font_color_label.height                   = 30
    $dialog_title_font_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_title_font_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_title_font_color_label);

                          
    $dialog_title_font_color_input.AutoSize                 = $false
    $dialog_title_font_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_title_font_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_title_font_color_input.Anchor                   = 'top,left'
    $dialog_title_font_color_input.width                    = 160
    $dialog_title_font_color_input.height                   = 27
    $dialog_title_font_color_input.location                 = New-Object System.Drawing.Point(($dialog_title_font_color_label.location.x + $dialog_title_font_color_label.width + 5),$y_pos)
    $dialog_title_font_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_title_font_color_input.text                     = $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] -replace '#',''
    $dialog_title_font_color_input.name                     = $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] -replace '#',''
    $dialog_title_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_title_font_color_input);

    
    $dialog_title_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_title_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_title_font_color_button.Width     = 100
    $dialog_title_font_color_button.height     = 27
    $dialog_title_font_color_button.Location  = New-Object System.Drawing.Point(($dialog_title_font_color_input.Location.x + $dialog_title_font_color_input.width + 5),$y_pos);
    $dialog_title_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_title_font_color_button.Text      ="Pick Color"
    $dialog_title_font_color_button.Name      = ""
    $dialog_title_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] = "#$color"
            $dialog_title_font_color_input.text = $color
            $dialog_title_font_color_input.name = $color
            $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] = "$color"
            $dialog_title_font_color_input.text = $color
            $dialog_title_font_color_input.name = $color
            $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_title_font_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Title Banner Color
    $y_pos = $y_pos + 30 
    
    $dialog_title_banner_color_label.text                     = "Title Banner Color:";
    $dialog_title_banner_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_title_banner_color_label.Anchor                   = 'top,right'
    $dialog_title_banner_color_label.TextAlign = "MiddleRight"
    $dialog_title_banner_color_label.width                    = 350
    $dialog_title_banner_color_label.height                   = 30
    $dialog_title_banner_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_title_banner_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_title_banner_color_label);

                         
    $dialog_title_banner_color_input.AutoSize                 = $false
    $dialog_title_banner_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_title_banner_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_title_banner_color_input.Anchor                   = 'top,left'
    $dialog_title_banner_color_input.width                    = 160
    $dialog_title_banner_color_input.height                   = 27
    $dialog_title_banner_color_input.location                 = New-Object System.Drawing.Point(($dialog_title_banner_color_label.location.x + $dialog_title_banner_color_label.width + 5),$y_pos)
    $dialog_title_banner_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_title_banner_color_input.text                     = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] -replace '#',''
    $dialog_title_banner_color_input.name                     = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] -replace '#',''
    $dialog_title_banner_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $title_label.Backcolor = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_title_banner_color_input);

    
    $dialog_title_banner_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_title_banner_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_title_banner_color_button.Width     = 100
    $dialog_title_banner_color_button.height     = 27
    $dialog_title_banner_color_button.Location  = New-Object System.Drawing.Point(($dialog_title_banner_color_input.Location.x + $dialog_title_banner_color_input.width + 5),$y_pos);
    $dialog_title_banner_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_title_banner_color_button.Text      ="Pick Color"
    $dialog_title_banner_color_button.Name      = ""
    $dialog_title_banner_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] = "#$color"
            $dialog_title_banner_color_input.text = $color
            $dialog_title_banner_color_input.name = $color
            $title_label.Backcolor = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] = "$color"
            $dialog_title_banner_color_input.text = $color
            $dialog_title_banner_color_input.name = $color
            $title_label.Backcolor = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_title_banner_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Sub Header Font Color
    $y_pos = $y_pos + 30 
    
    $dialog_sub_header_color_label.text                     = "Sub Header Font Color:";
    $dialog_sub_header_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_sub_header_color_label.Anchor                   = 'top,right'
    $dialog_sub_header_color_label.TextAlign = "MiddleRight"
    $dialog_sub_header_color_label.width                    = 350
    $dialog_sub_header_color_label.height                   = 30
    $dialog_sub_header_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_sub_header_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_sub_header_color_label);

                         
    $dialog_sub_header_color_input.AutoSize                 = $false
    $dialog_sub_header_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_sub_header_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_sub_header_color_input.Anchor                   = 'top,left'
    $dialog_sub_header_color_input.width                    = 160
    $dialog_sub_header_color_input.height                   = 27
    $dialog_sub_header_color_input.location                 = New-Object System.Drawing.Point(($dialog_sub_header_color_label.location.x + $dialog_sub_header_color_label.width + 5),$y_pos)
    $dialog_sub_header_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_sub_header_color_input.text                     = $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] -replace '#',''
    $dialog_sub_header_color_input.name                     = $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] -replace '#',''
    $dialog_sub_header_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] = "#" + $this.text
            
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $header_label1.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label2.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label3.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label4.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label5.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label6.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_sub_header_color_input);

    
    $dialog_sub_header_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_sub_header_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_sub_header_color_button.Width     = 100
    $dialog_sub_header_color_button.height     = 27
    $dialog_sub_header_color_button.Location  = New-Object System.Drawing.Point(($dialog_sub_header_color_input.Location.x + $dialog_sub_header_color_input.width + 5),$y_pos);
    $dialog_sub_header_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_sub_header_color_button.Text      ="Pick Color"
    $dialog_sub_header_color_button.Name      = ""
    $dialog_sub_header_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] = "#$color"
            $dialog_sub_header_color_input.text = $color
            $dialog_sub_header_color_input.name = $color
            $header_label1.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label2.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label3.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label4.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label5.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label6.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick

        }
        else
        {
            $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] = "$color"
            $dialog_sub_header_color_input.text = $color
            $dialog_sub_header_color_input.name = $color
            $header_label1.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label2.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label3.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label4.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label5.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $header_label6.ForeColor = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_sub_header_color_button);
    ######################################################################################################################
    
    
    
    ######################################################################################################################
    #Dialog Font Color
    $y_pos = $y_pos + 30 
    
    $dialog_font_color_label.text                       = "Font Color:";
    $dialog_font_color_label.ForeColor                  = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_font_color_label.Anchor                     = 'top,right'
    $dialog_font_color_label.TextAlign                  = "MiddleRight"
    $dialog_font_color_label.width                      = 350
    $dialog_font_color_label.height                     = 30
    $dialog_font_color_label.location                   = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_font_color_label.Font                       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_font_color_label);

                        
    $dialog_font_color_input.AutoSize                   = $false
    $dialog_font_color_input.ForeColor                  = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_font_color_input.BackColor                  = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_font_color_input.Anchor                     = 'top,left'
    $dialog_font_color_input.width                      = 160
    $dialog_font_color_input.height                     = 27
    $dialog_font_color_input.location                   = New-Object System.Drawing.Point(($dialog_font_color_label.location.x + $dialog_font_color_label.width + 5),$y_pos)
    $dialog_font_color_input.Font                       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_font_color_input.text                       = $script:theme_settings['DIALOG_FONT_COLOR'] -replace '#',''
    $dialog_font_color_input.name                       = $script:theme_settings['DIALOG_FONT_COLOR'] -replace '#',''
    $dialog_font_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_FONT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_FONT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $editor_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $main_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_text_color_label.ForeColor                      = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $adjustment_bars_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $interface_font_label.ForeColor                       = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_misspelled_font_color_label.ForeColor         = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_extend_acronym_font_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_shorten_acronym_font_color_label.ForeColor    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_highlight_color_label.ForeColor               = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_background_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_under_color_label.ForeColor          = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_over_color_label.ForeColor           = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $sidekick_background_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_font_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_banner_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_sub_header_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_text_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_background_color_label.ForeColor       = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_background_color_label.ForeColor        = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_text_color_label.ForeColor             = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_font_color_input);

    
    $dialog_font_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_font_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_font_color_button.Width     = 100
    $dialog_font_color_button.height     = 27
    $dialog_font_color_button.Location  = New-Object System.Drawing.Point(($dialog_font_color_input.Location.x + $dialog_font_color_input.width + 5),$y_pos);
    $dialog_font_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_font_color_button.Text      ="Pick Color"
    $dialog_font_color_button.Name      = ""
    $dialog_font_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_FONT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_FONT_COLOR'] = "#$color"
            $dialog_font_color_input.text = $color
            $dialog_font_color_input.name = $color
            $editor_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $main_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_text_color_label.ForeColor                      = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $adjustment_bars_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_misspelled_font_color_label.ForeColor         = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_extend_acronym_font_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_shorten_acronym_font_color_label.ForeColor    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_highlight_color_label.ForeColor               = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_background_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_under_color_label.ForeColor          = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_over_color_label.ForeColor           = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $sidekick_background_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_font_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_banner_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_sub_header_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_text_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_background_color_label.ForeColor       = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_background_color_label.ForeColor        = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_text_color_label.ForeColor             = $script:theme_settings['DIALOG_FONT_COLOR']
            $interface_font_label.ForeColor                       = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_FONT_COLOR'] = "$color"
            $dialog_font_color_input.text = $color
            $dialog_font_color_input.name = $color
            $editor_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_label.ForeColor                          = $script:theme_settings['DIALOG_FONT_COLOR']
            $main_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_text_color_label.ForeColor                      = $script:theme_settings['DIALOG_FONT_COLOR']
            $menu_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $adjustment_bars_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_misspelled_font_color_label.ForeColor         = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_extend_acronym_font_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_shorten_acronym_font_color_label.ForeColor    = $script:theme_settings['DIALOG_FONT_COLOR']
            $editor_highlight_color_label.ForeColor               = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_background_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_under_color_label.ForeColor          = $script:theme_settings['DIALOG_FONT_COLOR']
            $text_caclulator_over_color_label.ForeColor           = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $feeder_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $sidekick_background_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_font_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_title_banner_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_sub_header_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_text_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_background_color_label.ForeColor       = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_input_background_color_label.ForeColor        = $script:theme_settings['DIALOG_FONT_COLOR']
            $dialog_button_text_color_label.ForeColor             = $script:theme_settings['DIALOG_FONT_COLOR']
            $interface_font_label.ForeColor                       = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_font_color_button);
    ######################################################################################################################

    ######################################################################################################################
    #Dialog Input Text Color
    $y_pos = $y_pos + 30 
    
    $dialog_input_text_color_label.text                     = "Text Input Font Color:";
    $dialog_input_text_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_input_text_color_label.Anchor                   = 'top,right'
    $dialog_input_text_color_label.TextAlign = "MiddleRight"
    $dialog_input_text_color_label.width                    = 350
    $dialog_input_text_color_label.height                   = 30
    $dialog_input_text_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_input_text_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_input_text_color_label);

                         
    $dialog_input_text_color_input.AutoSize                 = $false
    $dialog_input_text_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_input_text_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_input_text_color_input.Anchor                   = 'top,left'
    $dialog_input_text_color_input.width                    = 160
    $dialog_input_text_color_input.height                   = 27
    $dialog_input_text_color_input.location                 = New-Object System.Drawing.Point(($dialog_input_text_color_label.location.x + $dialog_input_text_color_label.width + 5),$y_pos)
    $dialog_input_text_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_input_text_color_input.text                     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] -replace '#',''
    $dialog_input_text_color_input.name                     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] -replace '#',''
    $dialog_input_text_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $main_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_text_color_input.ForeColor                      = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $adjustment_bars_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_misspelled_font_color_input.ForeColor         = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_extend_acronym_font_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_shorten_acronym_font_color_input.ForeColor    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_highlight_color_input.ForeColor               = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_background_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_under_color_input.ForeColor          = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_over_color_input.ForeColor           = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $sidekick_background_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_font_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_banner_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_sub_header_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] 
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_text_color_input.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_background_color_input.ForeColor       = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_button.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_input_text_color_input);

    
    $dialog_input_text_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_input_text_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_input_text_color_button.Width     = 100
    $dialog_input_text_color_button.height     = 27
    $dialog_input_text_color_button.Location  = New-Object System.Drawing.Point(($dialog_input_text_color_input.Location.x + $dialog_input_text_color_input.width + 5),$y_pos);
    $dialog_input_text_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_input_text_color_button.Text      ="Pick Color"
    $dialog_input_text_color_button.Name      = ""
    $dialog_input_text_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] = "#$color"
            $dialog_input_text_color_input.text = $color
            $dialog_input_text_color_input.name = $color
            $main_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_text_color_input.ForeColor                      = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $adjustment_bars_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_misspelled_font_color_input.ForeColor         = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_extend_acronym_font_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_shorten_acronym_font_color_input.ForeColor    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_highlight_color_input.ForeColor               = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_background_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_under_color_input.ForeColor          = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_over_color_input.ForeColor           = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $sidekick_background_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_font_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_banner_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_sub_header_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_text_color_input.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_background_color_input.ForeColor       = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] = "$color"
            $dialog_input_text_color_input.text = $color
            $dialog_input_text_color_input.name = $color
            $main_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_text_color_input.ForeColor                      = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $menu_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $adjustment_bars_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_misspelled_font_color_input.ForeColor         = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_extend_acronym_font_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_shorten_acronym_font_color_input.ForeColor    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $editor_highlight_color_input.ForeColor               = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_background_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_under_color_input.ForeColor          = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $text_caclulator_over_color_input.ForeColor           = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $feeder_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $sidekick_background_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_font_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_title_banner_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_sub_header_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_text_color_input.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_button_background_color_input.ForeColor       = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_input_text_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Text Input Background Color
    $y_pos = $y_pos + 30 
    
    $dialog_input_background_color_label.text                     = "Text Input Background Color:";
    $dialog_input_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_input_background_color_label.Anchor                   = 'top,right'
    $dialog_input_background_color_label.TextAlign = "MiddleRight"
    $dialog_input_background_color_label.width                    = 350
    $dialog_input_background_color_label.height                   = 30
    $dialog_input_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_input_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_input_background_color_label);

                          
    $dialog_input_background_color_input.AutoSize                 = $false
    $dialog_input_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_input_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_input_background_color_input.Anchor                   = 'top,left'
    $dialog_input_background_color_input.width                    = 160
    $dialog_input_background_color_input.height                   = 27
    $dialog_input_background_color_input.location                 = New-Object System.Drawing.Point(($dialog_input_background_color_label.location.x + $dialog_input_background_color_label.width + 5),$y_pos)
    $dialog_input_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_input_background_color_input.text                     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] -replace '#',''
    $dialog_input_background_color_input.name                     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] -replace '#',''
    $dialog_input_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            
            $main_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_text_color_input.BackColor                      = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $adjustment_bars_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_misspelled_font_color_input.BackColor         = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_input.BackColor    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_highlight_color_input.BackColor               = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] 
            $text_caclulator_background_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_under_color_input.BackColor          = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_over_color_input.BackColor           = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $sidekick_background_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_font_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_banner_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_sub_header_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_background_color_input.BackColor        = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_text_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_text_color_input.backcolor             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_background_color_input.backcolor       = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick

        }
    })
    $color_form.controls.Add($dialog_input_background_color_input);

    
    $dialog_input_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_input_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_input_background_color_button.Width     = 100
    $dialog_input_background_color_button.height     = 27
    $dialog_input_background_color_button.Location  = New-Object System.Drawing.Point(($dialog_input_background_color_input.Location.x + $dialog_input_background_color_input.width + 5),$y_pos);
    $dialog_input_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_input_background_color_button.Text      ="Pick Color"
    $dialog_input_background_color_button.Name      = ""
    $dialog_input_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] = "#$color"
            $dialog_input_background_color_input.text = $color
            $dialog_input_background_color_input.name = $color
            $main_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_text_color_input.BackColor                      = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $adjustment_bars_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_misspelled_font_color_input.BackColor         = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_input.BackColor    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_highlight_color_input.BackColor               = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] 
            $text_caclulator_background_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_under_color_input.BackColor          = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_over_color_input.BackColor           = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $sidekick_background_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_font_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_banner_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_sub_header_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_background_color_input.BackColor        = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_text_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_text_color_input.backcolor             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_background_color_input.backcolor       = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick

        }
        else
        {
            $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] = "$color"
            $dialog_input_background_color_input.text = $color
            $dialog_input_background_color_input.name = $color
            $main_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_text_color_input.BackColor                      = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $menu_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $adjustment_bars_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_misspelled_font_color_input.BackColor         = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_input.BackColor    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $editor_highlight_color_input.BackColor               = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] 
            $text_caclulator_background_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_under_color_input.BackColor          = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $text_caclulator_over_color_input.BackColor           = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $feeder_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $sidekick_background_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_font_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_title_banner_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_sub_header_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_background_color_input.BackColor        = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_input_text_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_text_color_input.backcolor             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $dialog_button_background_color_input.backcolor       = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick

        }
    })
    $color_form.controls.Add($dialog_input_background_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Button Text Color
    $y_pos = $y_pos + 30 
    
    $dialog_button_text_color_label.text                     = "Button Text Color:";
    $dialog_button_text_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_button_text_color_label.Anchor                   = 'top,right'
    $dialog_button_text_color_label.TextAlign = "MiddleRight"
    $dialog_button_text_color_label.width                    = 350
    $dialog_button_text_color_label.height                   = 30
    $dialog_button_text_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_button_text_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_button_text_color_label);

                          
    $dialog_button_text_color_input.AutoSize                 = $false
    $dialog_button_text_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_button_text_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_button_text_color_input.Anchor                   = 'top,left'
    $dialog_button_text_color_input.width                    = 160
    $dialog_button_text_color_input.height                   = 27
    $dialog_button_text_color_input.location                 = New-Object System.Drawing.Point(($dialog_button_text_color_label.location.x + $dialog_button_text_color_label.width + 5),$y_pos)
    $dialog_button_text_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_button_text_color_input.text                     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] -replace '#',''
    $dialog_button_text_color_input.name                     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] -replace '#',''
    $dialog_button_text_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $main_background_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_text_color_button.ForeColor                      = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_background_color_button.ForeColor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $adjustment_bars_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_misspelled_font_color_button.forecolor         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_extend_acronym_font_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_shorten_acronym_font_color_button.forecolor    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_background_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_under_color_button.forecolor          = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_over_color_button.forecolor           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $sidekick_background_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_font_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_banner_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_sub_header_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_text_color_button.forecolor             = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_background_color_button.ForeColor        = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_text_color_button.ForeColor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_font_color_button.ForeColor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $manage_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $save_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $cancel_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_background_color_button.ForeColor       = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_button_text_color_input);

    
    $dialog_button_text_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_button_text_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_button_text_color_button.Width     = 100
    $dialog_button_text_color_button.height     = 27
    $dialog_button_text_color_button.Location  = New-Object System.Drawing.Point(($dialog_button_text_color_input.Location.x + $dialog_button_text_color_input.width + 5),$y_pos);
    $dialog_button_text_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_button_text_color_button.Text      ="Pick Color"
    $dialog_button_text_color_button.Name      = ""
    $dialog_button_text_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] = "#$color"
            $dialog_button_text_color_input.text = $color
            $dialog_button_text_color_input.name = $color
            $main_background_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_text_color_button.ForeColor                      = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_background_color_button.ForeColor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $adjustment_bars_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_misspelled_font_color_button.forecolor         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_extend_acronym_font_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_shorten_acronym_font_color_button.forecolor    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_background_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_under_color_button.forecolor          = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_over_color_button.forecolor           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $sidekick_background_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_font_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_banner_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_sub_header_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_text_color_button.forecolor             = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_background_color_button.ForeColor        = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_text_color_button.ForeColor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_font_color_button.ForeColor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $manage_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $save_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $cancel_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_background_color_button.ForeColor       = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] = "$color"
            $dialog_button_text_color_input.text = $color
            $dialog_button_text_color_input.name = $color
            $main_background_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_text_color_button.ForeColor                      = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $menu_background_color_button.ForeColor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $adjustment_bars_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_misspelled_font_color_button.forecolor         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_extend_acronym_font_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_shorten_acronym_font_color_button.forecolor    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_background_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_under_color_button.forecolor          = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $text_caclulator_over_color_button.forecolor           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $feeder_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $sidekick_background_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_font_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_title_banner_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_sub_header_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_text_color_button.forecolor             = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_background_color_button.ForeColor        = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_input_text_color_button.ForeColor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_font_color_button.ForeColor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $manage_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $save_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $cancel_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $dialog_button_background_color_button.ForeColor       = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_button_text_color_button);
    ######################################################################################################################
    ######################################################################################################################
    #Dialog Button Background Color
    $y_pos = $y_pos + 30 
    
    $dialog_button_background_color_label.text                     = "Button Background Color:";
    $dialog_button_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $dialog_button_background_color_label.Anchor                   = 'top,right'
    $dialog_button_background_color_label.TextAlign = "MiddleRight"
    $dialog_button_background_color_label.width                    = 350
    $dialog_button_background_color_label.height                   = 30
    $dialog_button_background_color_label.location                 = New-Object System.Drawing.Point(20,$y_pos);
    $dialog_button_background_color_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $color_form.controls.Add($dialog_button_background_color_label);

                        
    $dialog_button_background_color_input.AutoSize                 = $false
    $dialog_button_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $dialog_button_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $dialog_button_background_color_input.Anchor                   = 'top,left'
    $dialog_button_background_color_input.width                    = 160
    $dialog_button_background_color_input.height                   = 27
    $dialog_button_background_color_input.location                 = New-Object System.Drawing.Point(($dialog_button_background_color_label.location.x + $dialog_button_background_color_label.width + 5),$y_pos)
    $dialog_button_background_color_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $dialog_button_background_color_input.text                     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] -replace '#',''
    $dialog_button_background_color_input.name                     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] -replace '#',''
    $dialog_button_background_color_input.add_lostfocus({
        
        $this.text = $this.text -replace '#',''
        $this.text = $this.text -replace '\W',''
        $valid = 0;
        [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
        if($colors -match (" " + $this.text + " "))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] = $this.text
        }
        if(($this.text.length -eq 6) -and ($this.text -like '[a-f0-9]*') -and (!($this.text -match '[g-z]')))
        {
            $valid = 1;
            $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] = "#" + $this.text
        }
        if($valid -eq 0)
        {  
            $this.text = $this.name
            $message = "Invalid Color"
            [System.Windows.MessageBox]::Show($message,"Error",'Ok')     
        }
        else
        {
            $main_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_text_color_button.BackColor                      = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $adjustment_bars_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_misspelled_font_color_button.BackColor         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_button.BackColor    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_background_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_under_color_button.BackColor          = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_over_color_button.BackColor           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $sidekick_background_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_font_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_banner_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_sub_header_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_background_color_button.BackColor        = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_text_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $manage_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $save_theme_button.BackColor                           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $cancel_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_background_color_button.backcolor       = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_text_color_button.BackColor             = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_button_background_color_input);

    
    $dialog_button_background_color_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $dialog_button_background_color_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $dialog_button_background_color_button.Width     = 100
    $dialog_button_background_color_button.height     = 27
    $dialog_button_background_color_button.Location  = New-Object System.Drawing.Point(($dialog_button_background_color_input.Location.x + $dialog_button_background_color_input.width + 5),$y_pos);
    $dialog_button_background_color_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $dialog_button_background_color_button.Text      ="Pick Color"
    $dialog_button_background_color_button.Name      = ""
    $dialog_button_background_color_button.Add_Click({
        $script:color_picker = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        ($trash,$color) = color_picker
        #write-host $color = $script:color_picker
        if($color -match "\d")
        {
            $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] = "#$color"
            $dialog_button_background_color_input.text = $color
            $dialog_button_background_color_input.name = $color
            $main_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_text_color_button.BackColor                      = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $adjustment_bars_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_misspelled_font_color_button.BackColor         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_button.BackColor    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_background_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_under_color_button.BackColor          = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_over_color_button.BackColor           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $sidekick_background_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_font_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_banner_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_sub_header_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_background_color_button.BackColor        = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_text_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $manage_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $save_theme_button.BackColor                           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $cancel_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_background_color_button.backcolor       = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_text_color_button.BackColor             = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
        else
        {
            $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] = "$color"
            $dialog_button_background_color_input.text = $color
            $dialog_button_background_color_input.name = $color
            $main_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_text_color_button.BackColor                      = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $menu_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $adjustment_bars_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_misspelled_font_color_button.BackColor         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_extend_acronym_font_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_shorten_acronym_font_color_button.BackColor    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_background_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_under_color_button.BackColor          = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $text_caclulator_over_color_button.BackColor           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $feeder_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $sidekick_background_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_font_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_title_banner_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_sub_header_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_background_color_button.BackColor        = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_input_text_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $manage_theme_button.BackColor                           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $save_theme_button.BackColor                           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $cancel_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_background_color_button.backcolor       = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $dialog_button_text_color_button.BackColor             = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $script:sidekickgui = "New"
            update_sidekick
        }
    })
    $color_form.controls.Add($dialog_button_background_color_button);

    $y_pos = $y_pos + 35 
    ##############################################################################################################################
    $separator_bar2.text                        = ""
    $separator_bar2.AutoSize                    = $false
    $separator_bar2.BorderStyle                 = "fixed3d"
    #$separator_bar2.ForeColor                  = $script:theme_settings['DIALOG_BOX_BACKGROUND_COLOR']
    $separator_bar2.Anchor                      = 'top,left'
    $separator_bar2.width                       = (($color_form.width - 50) - $spacer)
    $separator_bar2.height                      = 1
    $separator_bar2.location                    = New-Object System.Drawing.Point(20,$y_pos)
    $separator_bar2.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $separator_bar2.TextAlign                   = 'MiddleLeft'
    $color_form.controls.Add($separator_bar2);

    $y_pos = $y_pos + 15 

    #$save_theme_button.autosize = $true
    #$save_theme_button.AutoSizeMode = 'GrowAndShrink'
    $save_theme_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $save_theme_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $save_theme_button.Width     = 125
    $save_theme_button.height     = 27
    $save_theme_button.Location  = New-Object System.Drawing.Point((($color_form.width / 2) - ($save_theme_button.width - 5)),($y_pos));
    $save_theme_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $save_theme_button.Text      ="Save Theme"
    $save_theme_button.Name = ""
    $save_theme_button.Add_Click({
        save_theme_dialog
        
    })
    $color_form.controls.Add($save_theme_button)

    $color_form.add_FormClosing({param($sender,$e)
        $found = 0;
        foreach($og in $script:theme_settings.GetEnumerator())
        {
            [string]$og1 = $script:theme_original[$og.key]
            [string]$og2 = $og.value
            if($og1 -ne $og2)
            {
                $found = 1;
            }
        }
        if($found -eq 0)
        {
            #No Changes          
            #$sender.close();     #Irrelevant in closing form, could be the culprit in crashes.
        }
        else
        {
            #Changes Found
            $message = "You haven't saved your settings, are you sure you want to leave?`n`n"
            $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
            if($yesno -eq "Yes")
            {
                $script:Form.BackColor                                = $script:theme_original['MAIN_BACKGROUND_COLOR']   
                $MenuBar.BackColor                                    = $script:theme_original['MENU_BACKGROUND_COLOR']
                $MenuBar.ForeColor                                    = $script:theme_original['MENU_TEXT_COLOR']
                $FileMenu.BackColor                                   = $script:theme_original['MENU_BACKGROUND_COLOR'] 
                $FileMenu.ForeColor                                   = $script:theme_original['MENU_TEXT_COLOR']
                $EditMenu.BackColor                                   = $script:theme_original['MENU_BACKGROUND_COLOR']
                $EditMenu.ForeColor                                   = $script:theme_original['MENU_TEXT_COLOR']
                $OptionsMenu.BackColor                                = $script:theme_original['MENU_BACKGROUND_COLOR']
                $OptionsMenu.ForeColor                                = $script:theme_original['MENU_TEXT_COLOR']
                $AboutMenu.BackColor                                  = $script:theme_original['MENU_BACKGROUND_COLOR']
                $AboutMenu.ForeColor                                  = $script:theme_original['MENU_TEXT_COLOR']
                $BulletMenu.BackColor                                 = $script:theme_original['MENU_BACKGROUND_COLOR']
                $BulletMenu.ForeColor                                 = $script:theme_original['MENU_TEXT_COLOR']
                $script:AcronymMenu.BackColor                         = $script:theme_original['MENU_BACKGROUND_COLOR']
                $script:AcronymMenu.ForeColor                         = $script:theme_original['MENU_TEXT_COLOR']
                $sidekick_panel.BackColor                      = $script:theme_original['ADJUSTMENT_BAR_COLOR']
                $bullet_feeder_panel.BackColor                 = $script:theme_original['ADJUSTMENT_BAR_COLOR']
                $editor.BackColor                                     = $script:theme_original['EDITOR_BACKGROUND_COLOR']
                $sizer_box.backcolor                                  = $script:theme_original['TEXT_CALCULATOR_BACKGROUND_COLOR']
                $feeder_box.BackColor                                 = $script:theme_original['FEEDER_BACKGROUND_COLOR']
                $feeder_box.ForeColor                                 = $script:theme_original['FEEDER_FONT_COLOR']
                $left_panel.BackColor                          = $script:theme_original['SIDEKICK_BACKGROUND_COLOR']
                load_theme $script:settings['THEME']
                build_file_menu
                build_options_menu
                build_about_menu
                build_bullet_menu
                build_acronym_menu
                $Script:recent_editor_text = "Changed"
                $script:sidekickgui = "New"
                update_sidekick
                #$color_form.close();       
            }
            else
            {
                $e.cancel = $true
            }
        }
    });



    
    $cancel_theme_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_theme_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_theme_button.Width     = 125
    #$cancel_theme_button.autosize = $true
    #$cancel_theme_button.AutoSizeMode = 'GrowAndShrink'
    $cancel_theme_button.height     = 27
    $cancel_theme_button.Location  = New-Object System.Drawing.Point(($save_theme_button.Width + $save_theme_button.location.x  + 5),($y_pos));
    $cancel_theme_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_theme_button.Text      ="Cancel"
    $cancel_theme_button.Add_Click({
           $color_form.close();
    });
    $color_form.controls.Add($cancel_theme_button)




    $color_form.refresh();
    $color_form.ShowDialog()
    
}
################################################################################
######Color Picker##############################################################
function color_picker
{
    $colorDialog = new-object System.Windows.Forms.ColorDialog
    $colorDialog.AllowFullOpen = $true
    $colorDialog.FullOpen = $true
    $colorDialog.color = $script:color_picker
    $colorDialog.ShowDialog()
    [string]$colors = (" " + ([System.Drawing.Color] | gm -Static -MemberType Properties).name + " ")
    if($colors -match $colorDialog.Color.Name)
    {
        return $colorDialog.Color.Name
    }
    else
    {
        $red = [System.Convert]::ToString($colordialog.color.R,16)
        $green = [System.Convert]::ToString($colordialog.color.G,16)
        $blue = [System.Convert]::ToString($colordialog.color.B,16)
        if($red.length -eq 1){$red = "0" + "$red"}
        if($green.length -eq 1){$green = "0" + "$green"}
        if($blue.length -eq 1){$blue = "0" + "$blue"}
        $hex = "$red" + "$green" + "$blue"
        return $hex
    }     
}
################################################################################
######Load Theme################################################################
function load_theme($theme)
{
    
    $script:theme_settings = @{};
    $loader = 0;
    if($theme -eq $null)
    {
        $theme = $settings['THEME']
    }
    else
    {
        $loader = 1;
    }
    #write-host Loading Theme: $theme
    ###########################################################################
    if(!(Test-Path "$dir\Resources\Themes\$theme.csv"))
    {
        $settings['THEME'] = "Blue Falcon"
        $theme = $settings['THEME']
    }
    if(Test-Path "$dir\Resources\Themes\$theme.csv")
    {
        $line_count = 0;
        $reader = [System.IO.File]::OpenText("$dir\Resources\Themes\$theme.csv")
        while($null -ne ($line = $reader.ReadLine()))
        {
            $line_count++;
            if($line_count -ne 1)
            {
                ($key,$value) = $line -split ',',2
                if(!($script:theme_settings.containskey($key)))
                {
                    $script:theme_settings.Add($key,$value);
                }
                #write-host $key
                #write-host $value
            } 
        }
        $reader.close();
        #####################################################
        #Create OG Values
        $script:theme_original = @{};
        foreach($color in $script:theme_settings.GetEnumerator())
        {
            if(!($script:theme_original.Contains($color.key)))
            {
                $script:theme_original.Add($color.key,$color.value);
            }
        }
        ##################################################### 
        update_settings
        
    }
    #############################################################################
    if($loader -eq 1)
    {
        $script:Form.BackColor                                = $script:theme_settings['MAIN_BACKGROUND_COLOR']

        $MenuBar.ForeColor                                    = $script:theme_settings['MENU_TEXT_COLOR']
        $MenuBar.BackColor                                    = $script:theme_settings['MENU_BACKGROUND_COLOR']      
        $FileMenu.ForeColor                                   = $script:theme_settings['MENU_TEXT_COLOR']
        $FileMenu.BackColor                                   = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $EditMenu.ForeColor                                   = $script:theme_settings['MENU_TEXT_COLOR']
        $EditMenu.BackColor                                   = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $OptionsMenu.ForeColor                                = $script:theme_settings['MENU_TEXT_COLOR']
        $OptionsMenu.BackColor                                = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $AboutMenu.ForeColor                                  = $script:theme_settings['MENU_TEXT_COLOR']
        $AboutMenu.BackColor                                  = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $BulletMenu.ForeColor                                 = $script:theme_settings['MENU_TEXT_COLOR']
        $BulletMenu.BackColor                                 = $script:theme_settings['MENU_BACKGROUND_COLOR']
        $script:AcronymMenu.ForeColor                         = $script:theme_settings['MENU_TEXT_COLOR']
        $script:AcronymMenu.BackColor                         = $script:theme_settings['MENU_BACKGROUND_COLOR']

        $sidekick_panel.BackColor                      = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        $bullet_feeder_panel.BackColor                 = $script:theme_settings['ADJUSTMENT_BAR_COLOR']
        $editor.BackColor                                     = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
        $sizer_box.backcolor                                  = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR']
        $feeder_box.BackColor                                 = $script:theme_settings['FEEDER_BACKGROUND_COLOR']
        $feeder_box.ForeColor                                 = $script:theme_settings['FEEDER_FONT_COLOR']
        $left_panel.BackColor                          = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR']
        $color_form.Backcolor                                 = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
        $title_label.ForeColor                                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
        $title_label.Backcolor                                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
        $header_label1.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $header_label2.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $header_label3.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $header_label4.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $header_label5.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $header_label6.ForeColor                              = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
                  
        $main_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
        $menu_text_color_label.ForeColor                      = $script:theme_settings['DIALOG_FONT_COLOR']
        $menu_background_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
        $adjustment_bars_color_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
        $interface_font_label.ForeColor                       = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_misspelled_font_color_label.ForeColor         = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_extend_acronym_font_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_shorten_acronym_font_color_label.ForeColor    = $script:theme_settings['DIALOG_FONT_COLOR']
        $editor_highlight_color_label.ForeColor               = $script:theme_settings['DIALOG_FONT_COLOR']
        $text_caclulator_background_color_label.ForeColor     = $script:theme_settings['DIALOG_FONT_COLOR']
        $text_caclulator_under_color_label.ForeColor          = $script:theme_settings['DIALOG_FONT_COLOR']
        $text_caclulator_over_color_label.ForeColor           = $script:theme_settings['DIALOG_FONT_COLOR']
        $feeder_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $feeder_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
        $sidekick_background_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_background_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_title_font_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_title_banner_color_label.ForeColor            = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_sub_header_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_font_color_label.ForeColor                    = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_input_text_color_label.ForeColor              = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_button_background_color_label.ForeColor       = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_input_background_color_label.ForeColor        = $script:theme_settings['DIALOG_FONT_COLOR']
        $dialog_button_text_color_label.ForeColor             = $script:theme_settings['DIALOG_FONT_COLOR']

        $main_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $menu_text_color_input.ForeColor                      = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $menu_background_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $adjustment_bars_color_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_misspelled_font_color_input.ForeColor         = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_extend_acronym_font_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_shorten_acronym_font_color_input.ForeColor    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $editor_highlight_color_input.ForeColor               = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $text_caclulator_background_color_input.ForeColor     = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $text_caclulator_under_color_input.ForeColor          = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $text_caclulator_over_color_input.ForeColor           = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $feeder_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $feeder_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $sidekick_background_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_background_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_title_font_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_title_banner_color_input.ForeColor            = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_font_color_input.ForeColor                    = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_sub_header_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_input_text_color_input.ForeColor              = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_button_text_color_input.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_button_background_color_input.ForeColor       = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_input_background_color_input.ForeColor        = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
        $dialog_input_text_color_button.ForeColor             = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']

        $main_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $menu_text_color_input.BackColor                      = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $menu_background_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $adjustment_bars_color_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_misspelled_font_color_input.BackColor         = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_extend_acronym_font_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_shorten_acronym_font_color_input.BackColor    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $editor_highlight_color_input.BackColor               = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] 
        $text_caclulator_background_color_input.BackColor     = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $text_caclulator_under_color_input.BackColor          = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $text_caclulator_over_color_input.BackColor           = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $feeder_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $feeder_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $sidekick_background_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_background_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_title_font_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_title_banner_color_input.BackColor            = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_sub_header_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_input_background_color_input.BackColor        = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_input_text_color_input.BackColor              = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_font_color_input.BackColor                    = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_button_text_color_input.backcolor             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
        $dialog_button_background_color_input.backcolor       = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']

        $main_background_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $menu_text_color_button.ForeColor                      = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $menu_background_color_button.ForeColor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $adjustment_bars_color_button.forecolor                = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_misspelled_font_color_button.forecolor         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_extend_acronym_font_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_shorten_acronym_font_color_button.forecolor    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $editor_highlight_color_button.forecolor               = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $text_caclulator_background_color_button.forecolor     = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $text_caclulator_under_color_button.forecolor          = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $text_caclulator_over_color_button.forecolor           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $feeder_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $feeder_font_color_button.forecolor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $sidekick_background_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_background_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_title_font_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_title_banner_color_button.forecolor            = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_sub_header_color_button.forecolor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_button_text_color_button.forecolor             = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_input_background_color_button.ForeColor        = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_input_text_color_button.ForeColor              = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_font_color_button.ForeColor                    = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $manage_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $save_theme_button.ForeColor                           = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $cancel_theme_button.ForeColor                         = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $dialog_button_background_color_button.ForeColor       = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']

        $main_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $menu_text_color_button.BackColor                      = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $menu_background_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $adjustment_bars_color_button.BackColor                = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_misspelled_font_color_button.BackColor         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_extend_acronym_font_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_shorten_acronym_font_color_button.BackColor    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $editor_highlight_color_button.BackColor               = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $text_caclulator_background_color_button.BackColor     = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $text_caclulator_under_color_button.BackColor          = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $text_caclulator_over_color_button.BackColor           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $feeder_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $feeder_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $sidekick_background_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_background_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_title_font_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_title_banner_color_button.BackColor            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_sub_header_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_input_background_color_button.BackColor        = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_input_text_color_button.BackColor              = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_font_color_button.BackColor                    = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $manage_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $save_theme_button.BackColor                           = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $cancel_theme_button.BackColor                         = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_button_background_color_button.backcolor       = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $dialog_button_text_color_button.BackColor             = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']


        $main_background_color_input.text                     = $script:theme_settings['MAIN_BACKGROUND_COLOR'] -replace '#',''
        $main_background_color_input.Name                     = $script:theme_settings['MAIN_BACKGROUND_COLOR'] -replace '#',''
        $menu_text_color_input.text                           = $script:theme_settings['MENU_TEXT_COLOR'] -replace '#',''
        $menu_text_color_input.name                           = $script:theme_settings['MENU_TEXT_COLOR'] -replace '#',''
        $menu_background_color_input.text                     = $script:theme_settings['MENU_BACKGROUND_COLOR'] -replace '#',''
        $menu_background_color_input.name                     = $script:theme_settings['MENU_BACKGROUND_COLOR'] -replace '#',''
        $adjustment_bars_color_input.text                     = $script:theme_settings['ADJUSTMENT_BAR_COLOR'] -replace '#',''
        $adjustment_bars_color_input.name                     = $script:theme_settings['ADJUSTMENT_BAR_COLOR'] -replace '#',''
        $editor_background_color_input.text                   = $script:theme_settings['EDITOR_BACKGROUND_COLOR'] -replace '#',''
        $editor_background_color_input.name                   = $script:theme_settings['EDITOR_BACKGROUND_COLOR'] -replace '#',''
        $editor_font_color_input.text                         = $script:theme_settings['EDITOR_FONT_COLOR'] -replace '#',''
        $editor_font_color_input.name                         = $script:theme_settings['EDITOR_FONT_COLOR'] -replace '#',''
        $editor_misspelled_font_color_input.text              = $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] -replace '#',''    
        $editor_misspelled_font_color_input.name              = $script:theme_settings['EDITOR_MISSPELLED_FONT_COLOR'] -replace '#',''  
        $editor_extend_acronym_font_color_input.text          = $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] -replace '#',''              
        $editor_extend_acronym_font_color_input.name          = $script:theme_settings['EDITOR_EXTEND_ACRONYM_FONT_COLOR'] -replace '#','' 
        $editor_shorten_acronym_font_color_input.text         = $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] -replace '#',''
        $editor_shorten_acronym_font_color_input.name         = $script:theme_settings['EDITOR_SHORTEN_ACRONYM_FONT_COLOR'] -replace '#',''
        $editor_highlight_color_input.text                    = $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] -replace '#',''
        $editor_highlight_color_input.name                    = $script:theme_settings['EDITOR_HIGHLIGHT_COLOR'] -replace '#',''
        $text_caclulator_background_color_input.text          = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] -replace '#',''
        $text_caclulator_background_color_input.name          = $script:theme_settings['TEXT_CALCULATOR_BACKGROUND_COLOR'] -replace '#',''
        $text_caclulator_under_color_input.text               = $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] -replace '#',''
        $text_caclulator_under_color_input.name               = $script:theme_settings['TEXT_CALCULATOR_UNDER_COLOR'] -replace '#',''
        $text_caclulator_over_color_input.text                = $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] -replace '#',''
        $text_caclulator_over_color_input.name                = $script:theme_settings['TEXT_CALCULATOR_OVER_COLOR'] -replace '#',''
        $feeder_background_color_input.text                   = $script:theme_settings['FEEDER_BACKGROUND_COLOR'] -replace '#',''
        $feeder_background_color_input.name                   = $script:theme_settings['FEEDER_BACKGROUND_COLOR'] -replace '#',''
        $feeder_font_color_input.text                         = $script:theme_settings['FEEDER_FONT_COLOR'] -replace '#',''
        $feeder_font_color_input.name                         = $script:theme_settings['FEEDER_FONT_COLOR'] -replace '#',''
        $sidekick_background_color_input.text                 = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] -replace '#',''
        $sidekick_background_color_input.name                 = $script:theme_settings['SIDEKICK_BACKGROUND_COLOR'] -replace '#',''
        $dialog_background_color_input.text                   = $script:theme_settings['DIALOG_BACKGROUND_COLOR'] -replace '#',''
        $dialog_background_color_input.name                   = $script:theme_settings['DIALOG_BACKGROUND_COLOR'] -replace '#',''
        $dialog_title_font_color_input.text                   = $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] -replace '#',''
        $dialog_title_font_color_input.name                   = $script:theme_settings['DIALOG_TITLE_FONT_COLOR'] -replace '#',''
        $dialog_title_banner_color_input.text                 = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] -replace '#',''
        $dialog_title_banner_color_input.name                 = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR'] -replace '#',''
        $dialog_sub_header_color_input.text                   = $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] -replace '#',''
        $dialog_sub_header_color_input.name                   = $script:theme_settings['DIALOG_SUB_HEADER_COLOR'] -replace '#',''
        $dialog_font_color_input.text                         = $script:theme_settings['DIALOG_FONT_COLOR'] -replace '#',''
        $dialog_font_color_input.name                         = $script:theme_settings['DIALOG_FONT_COLOR'] -replace '#',''
        $dialog_input_text_color_input.text                   = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] -replace '#',''
        $dialog_input_text_color_input.name                   = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR'] -replace '#',''
        $dialog_input_background_color_input.text             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] -replace '#',''
        $dialog_input_background_color_input.name             = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR'] -replace '#',''
        $dialog_button_text_color_input.text                  = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] -replace '#',''
        $dialog_button_text_color_input.name                  = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR'] -replace '#',''
        $dialog_button_background_color_input.text            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] -replace '#',''
        $dialog_button_background_color_input.name            = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR'] -replace '#',''

        $interface_font_combo.SelectedItem      = $script:theme_settings['INTERFACE_FONT']
        $interface_font_size_combo.SelectedItem = $script:theme_settings['INTERFACE_FONT_SIZE']    
        $editor_font_combo.SelectedItem         = $script:theme_settings['EDITOR_FONT']  
        $editor_font_size_combo.SelectedItem    = $script:theme_settings['EDITOR_FONT_SIZE']
        $feeder_font_combo.SelectedItem         = $script:theme_settings['FEEDER_FONT']
        $feeder_font_size_combo.SelectedItem    = $script:theme_settings['FEEDER_FONT_SIZE']

        $title_label.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
        $theme_combo.font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $manage_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar1.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $save_theme_button.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $cancel_theme_button.Font                      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $separator_bar2.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $header_label1.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $main_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $main_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $main_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $menu_text_color_label.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_text_color_input.Font                    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_text_color_button.Font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $menu_background_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $menu_background_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_label.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $adjustment_bars_color_input.Font              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $adjustment_bars_color_button.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label2.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $editor_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_misspelled_font_color_label.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_misspelled_font_color_input.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_misspelled_font_color_button.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $editor_extend_acronym_font_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_extend_acronym_font_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_extend_acronym_font_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $editor_shorten_acronym_font_color_label.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_shorten_acronym_font_color_input.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_shorten_acronym_font_color_button.Font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_label.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_highlight_color_input.Font             = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $editor_highlight_color_button.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $editor_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $editor_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label3.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $text_caclulator_background_color_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_background_color_input.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_background_color_button.Font  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_label.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_under_color_input.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_under_color_button.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $text_caclulator_over_color_label.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $text_caclulator_over_color_input.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $text_caclulator_over_color_button.Font        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label4.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $feeder_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $feeder_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $feeder_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $feeder_font_label.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $feeder_font_combo.font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $feeder_font_size_combo.font                   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $header_label5.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))
        $sidekick_background_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $sidekick_background_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $sidekick_background_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $header_label6.Font                            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 1))   
        $dialog_background_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_background_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_background_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])   
        $dialog_title_font_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_font_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_font_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_title_banner_color_label.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_title_banner_color_input.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_title_banner_color_button.Font         = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_sub_header_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_sub_header_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_font_color_label.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_font_color_input.Font                  = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_font_color_button.Font                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_text_color_label.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_text_color_input.Font            = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_text_color_button.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $dialog_input_background_color_label.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_input_background_color_input.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_input_background_color_button.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])  
        $dialog_button_text_color_label.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_text_color_input.Font           = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_text_color_button.Font          = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_label.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $dialog_button_background_color_input.Font     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $dialog_button_background_color_button.Font    = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']) 
        $interface_font_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $interface_font_combo.font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $interface_font_size_combo.font                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
        $FileMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $EditMenu.Font                                 = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $BulletMenu.Font                               = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $script:AcronymMenu.Font                       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $OptionsMenu.Font                              = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
        $AboutMenu.Font                                = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
                                           
        $script:Form.refresh();
        $sizer_art.refresh();
        
        #build_file_menu
        #build_bullet_menu
        #build_acronym_menu
        $Script:recent_editor_text = "Changed"
        $script:sidekickgui = "New"
        sidekick_display
    }
}
################################################################################
######Save Theme Dialog#########################################################
function save_theme_dialog
{
    $save_theme_form = New-Object System.Windows.Forms.Form
    $save_theme_form.FormBorderStyle = 'Fixed3D'
    $save_theme_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $save_theme_form.Location = new-object System.Drawing.Point(0, 0)
    $save_theme_form.Size = new-object System.Drawing.Size(440, 120)
    $save_theme_form.MaximizeBox = $false
    $save_theme_form.SizeGripStyle = "Hide"
    $save_theme_form.Text = "Save Theme"
    #$save_theme_form.TopMost = $True
    $save_theme_form.TabIndex = 0
    $save_theme_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    $theme_name_label                          = New-Object system.Windows.Forms.Label
    $theme_name_label.text                     = "Theme Name:";
    $theme_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $theme_name_label.Anchor                   = 'top,right'
    $theme_name_label.width                    = 160
    $theme_name_label.height                   = 30
    $theme_name_label.location                 = New-Object System.Drawing.Point(10,10)
    $theme_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))

    $save_theme_form.controls.Add($theme_name_label);

    $theme_name_input                         = New-Object system.Windows.Forms.TextBox                       
    $theme_name_input.AutoSize                 = $false
    $theme_name_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
    $theme_name_input.Backcolor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
    $theme_name_input.Anchor                   = 'top,left'
    $theme_name_input.width                    = 250
    $theme_name_input.height                   = 30
    $theme_name_input.location                 = New-Object System.Drawing.Point(($theme_name_label.Location.x + $theme_name_label.Width + 5) ,12)
    $theme_name_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
    $theme_name_input.text                     = $script:settings['THEME']
    $theme_name_input.name                     = $script:settings['THEME']
    $theme_name_input.Add_TextChanged({
        $caret = $theme_name_input.SelectionStart;
        $theme_name_input.text = $theme_name_input.text -replace '[^0-9A-Za-z ,-]', ''
        $theme_name_input.text = $theme_name_input.text.Split([IO.Path]::GetInvalidFileNameChars()) -join ' '

        #$theme_name_input.text = (Get-Culture).TextInfo.ToTitleCase($theme_name_input.text)
        $theme_name_input.SelectionStart = $caret
    });
    $save_theme_form.controls.Add($theme_name_input);

    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($save_theme_form.width / 2) - ($submit_button.width)),45);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Save"
    $submit_button.Name = ""
    $submit_button.Add_Click({ 
        [array]$errors = "";
        $og_theme = $script:settings['THEME']
        $new_theme = $theme_name_input.text

        if($new_theme -eq "")
        {
            $errors += "You must provide a name."
        }
        if($errors.count -eq 1)
        {
            if(Test-Path -literalpath "$dir\Resources\Themes\$new_theme.csv")
            {
                $message = "`"$new_theme`" already exists. Overwrite?`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Overwrite?", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {         
                    Remove-item "$dir\Resources\Themes\$new_theme.csv" -Recurse
                    $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\$new_theme.csv",$true)
                    $theme_writer.write("PROPERTY,VALUE`r`n");
                    foreach($color in $script:theme_settings.GetEnumerator() | sort key)
                    {
                        $line = $color.key
                        $line = csv_write_line $line $color.value
                        $theme_writer.write("$line`r`n");
                    }
                    $theme_writer.close();
                    
                    

                    $script:theme_original = @{};
                    foreach($color in $script:theme_settings.GetEnumerator())
                    {
                        if(!($script:theme_original.Contains($color.key)))
                        {
                            $script:theme_original.Add($color.key,$color.value);
                        }
                    }

                    $script:settings['THEME'] = $new_theme
                    $theme_combo.SelectedItem = "$new_theme"
                    #write-host 1 - $editor.zoomfactor
                    update_settings
                    build_file_menu
                    build_options_menu
                    build_about_menu
                    #build_bullet_menu
                    #build_acronym_menu
                    #write-host 2 - $editor.zoomfactor
                    $save_theme_form.close();
                    $color_form.close();
                }
            }
            else
            {
                $message = "Are you sure you want to save Theme as `"$new_theme`"`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","Save?", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {         
                    $theme_writer = new-object system.IO.StreamWriter("$dir\Resources\Themes\$new_theme.csv",$true)
                    $theme_writer.write("PROPERTY,VALUE`r`n");
                    foreach($color in $script:theme_settings.GetEnumerator() | sort key)
                    {
                        $line = $color.key
                        $line = csv_write_line $line $color.value
                        $theme_writer.write("$line`r`n");
                    }
                    $theme_writer.close();
                    
                    $script:theme_original = @{};
                    foreach($color in $script:theme_settings.GetEnumerator())
                    {
                        if(!($script:theme_original.Contains($color.key)))
                        {
                            $script:theme_original.Add($color.key,$color.value);
                        }
                    }
                                  
                    $script:settings['THEME'] = $new_theme
                    $theme_combo.Items.Add("$new_theme");
                    $theme_combo.SelectedItem = "$new_theme"
                    update_settings
                    build_file_menu
                    build_options_menu
                    build_about_menu
                    build_bullet_menu
                    build_acronym_menu
                    $save_theme_form.close();
                    $color_form.close();
                    
                }
            }
        }
        else
        {
            $message = "Please fix the following errors:`n`n"
            $counter = 0;
            foreach($error in $errors)
            {
                if($error -ne "")
                {
                    $counter++;
                    $message = $message + "$counter - $error`n"
                } 
            }
            [System.Windows.MessageBox]::Show($message,"Error",'Ok','Error')
        }



    });
    $save_theme_form.controls.Add($submit_button)

    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($save_theme_form.width / 2)),45);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
        $script:return = 0;
        $save_theme_form.close();
            
    });
    $save_theme_form.controls.Add($cancel_button) 

    $null = $save_theme_form.ShowDialog()
}
################################################################################
######Manage Themes#############################################################
function manage_themes
{

    $themes = Get-ChildItem -Path "$dir\Resources\Themes" -File -Force -Filter *.csv

    $spacer = 0;
    $manage_themes_form = New-Object System.Windows.Forms.Form
    $manage_themes_form.FormBorderStyle = 'Fixed3D'
    $manage_themes_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $manage_themes_form.StartPosition = "CenterScreen"
    $manage_themes_form.MaximizeBox = $false
    $manage_themes_form.SizeGripStyle = "Hide"
    $manage_themes_form.Width = 600
    if($themes.get_count() -eq 0)
    {
        $manage_themes_form.Height = 200;
    }
    elseif((($themes.get_count() * 65) + 140) -ge 600)
    {
        $manage_themes_form.Height = 600;
        $manage_themes_form.Autoscroll = $true
        $spacer = 20
    }
    else
    {
        $manage_themes_form.Height = (($themes.get_count() * 65) + 140)
    }
    $manage_themes_form.Text = "Manage Themes"
    #$manage_themes_form.TopMost = $True
    $manage_themes_form.TabIndex = 0
    #$manage_themes_form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen

    #################################################################################################
    $y_pos = 10;

    $title_label1                          = New-Object system.Windows.Forms.Label
    $title_label1.text                     = "Manage Themes";
    $title_label1.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label1.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label1.Anchor                   = 'top,right'
    $title_label1.width                    = ($manage_themes_form.width)
    $title_label1.height                   = 30
    $title_label1.TextAlign = "MiddleCenter"
    $title_label1.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $title_label1.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $manage_themes_form.controls.Add($title_label1);

    $y_pos = $y_pos + 30;
    $message_box                          = New-Object system.Windows.Forms.Label
    $message_box.text                     = "";
    $message_box.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $message_box.Anchor                   = 'top,right'
    $message_box.width                    = ($manage_themes_form.width)
    $message_box.height                   = 30
    $message_box.TextAlign = "MiddleCenter"
    $message_box.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $message_box.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])
    $manage_themes_form.controls.Add($message_box);
    $y_pos = $y_pos + 30;
    
    foreach($theme in $themes)
    {
        $theme = [System.IO.Path]::GetFileNameWithoutExtension($theme)
        

        $them_name_label                          = New-Object system.Windows.Forms.Label
        $them_name_label.text                     = "$theme";
        $them_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
        $them_name_label.Anchor                   = 'top,right'
        $them_name_label.width                    = (($manage_themes_form.width - 50) - $spacer)
        $them_name_label.height                   = 30
        $them_name_label.location                 = New-Object System.Drawing.Point((20 + $spacer),$y_pos)
        $them_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $manage_themes_form.controls.Add($them_name_label);

        $y_pos = $y_pos + 30
        $load_button           = New-Object System.Windows.Forms.Button
        $load_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $load_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $load_button.Width     = 120
        $load_button.height     = 25
        $load_button.Location  = New-Object System.Drawing.Point(20,$y_pos);
        $load_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
        $load_button.Text      = "Load"
        $load_button.Name      = $theme
        $load_button.Add_Click({
            $found = 1;
            foreach($og in $script:theme_settings.GetEnumerator())
            {
                [string]$og1 = $script:theme_original[$og.key]
                [string]$og2 = $og.value
                if($og1 -ne $og2)
                {
                    $found = 0;
                }
            }
            if($found -eq 1)
            { 
                $settings['THEME'] = $this.name
                load_theme $this.name
                $manage_themes_form.close();
                $script:reload_function = "manage_themes"
            }
            else
            {
                #Changes Found
                $message = "Loading a new theme will clear your changes. Are you sure you want to continue?`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    $settings['THEME'] = $this.name
                    load_theme $this.name
                    $manage_themes_form.close();
                    $script:reload_function = "manage_themes"  
                }
            }       
        });
        $manage_themes_form.controls.Add($load_button)

        $rename_button           = New-Object System.Windows.Forms.Button
        $rename_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $rename_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $rename_button.Width     = 120
        $rename_button.height     = 25
        $rename_button.Location  = New-Object System.Drawing.Point((20 + $load_button.Width),$y_pos);
        $rename_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
        $rename_button.Text      = "Rename"
        $rename_button.Name      = $theme
        $rename_button.Add_Click({
            #write-host Rename $this.name
            $old_name = "$dir\Resources\Themes\" + $this.name + ".csv"
            $new_name = rename_dialog $old_name

            $old_key = [System.IO.Path]::GetFileNameWithoutExtension($old_name)
            $new_key = [System.IO.Path]::GetFileNameWithoutExtension($new_name)
            if($settings['THEME'] -eq $old_key)
            {
                $settings['THEME'] = $new_key
                update_settings
            }
            $manage_themes_form.close();
            $script:reload_function = "manage_themes"
        });
        $manage_themes_form.controls.Add($rename_button)

        $delete_button           = New-Object System.Windows.Forms.Button
        $delete_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
        $delete_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
        $delete_button.Width     = 120
        $delete_button.height     = 25
        $delete_button.Location  = New-Object System.Drawing.Point((20 + $load_button.Width + $rename_button.Width),$y_pos);
        $delete_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
        $delete_button.Text      = "Delete"
        $delete_button.Name      = $theme
        $delete_button.Add_Click({
            #write-host Delete $this.name
            $theme = $this.name
            if($settings['THEME'] -eq $this.name)
            {
               $message = "You can not delete an Active theme.`n"
               [System.Windows.MessageBox]::Show($message,"!!!ERROR!!!",'Ok') 
            }
            else
            {

                $message = "Are you sure you want to delete `"$theme`" theme? You cannot revert this action.`n`n"
                $yesno = [System.Windows.Forms.MessageBox]::Show("$message","!!!WARNING!!!", "YesNo" , "Information" , "Button1")
                if($yesno -eq "Yes")
                {
                    $delete_name = "$dir\Resources\Themes\" + $this.name + ".csv"
                    if(Test-Path -literalpath $delete_name)
                    {
                        Remove-Item -literalpath $delete_name -Force
                    }
                    $manage_themes_form.close();
                    $script:reload_function = "manage_themes"
                }
            }
        });
        $manage_themes_form.controls.Add($delete_button)

        if($script:settings['THEME'] -eq $theme)
        {
            $theme_active_label                          = New-Object system.Windows.Forms.Label
            $theme_active_label.text                     = "Active";
            $theme_active_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $theme_active_label.Anchor                   = 'top,right'
            $theme_active_label.width                    = 100
            $theme_active_label.height                   = 30
            $theme_active_label.location                 = New-Object System.Drawing.Point(($manage_themes_form.width - ($spacer + $theme_active_label.width)),$y_pos)
            $theme_active_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $manage_themes_form.controls.Add($theme_active_label);

        }
      
        $y_pos = $y_pos + 30
        $separator_bar                             = New-Object system.Windows.Forms.Label
        $separator_bar.text                        = ""
        $separator_bar.AutoSize                    = $false
        $separator_bar.BorderStyle                 = "fixed3d"
        #$separator_bar.ForeColor                   = $script:settings['DIALOG_BOX_TEXT_BOLD_COLOR']
        $separator_bar.Anchor                      = 'top,left'
        $separator_bar.width                       = (($manage_themes_form.width - 50) - $spacer)
        $separator_bar.height                      = 1
        $separator_bar.location                    = New-Object System.Drawing.Point(20,$y_pos)
        $separator_bar.Font                        = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
        $separator_bar.TextAlign                   = 'MiddleLeft'
        $manage_themes_form.controls.Add($separator_bar);
        $y_pos = $y_pos + 5

    }
    $manage_themes_form.ShowDialog()
}
################################################################################
######Update Sidekick###########################################################
function update_sidekick
{
    if($sidekick_panel.width -ne 5)
    {
        if($script:sidekick_job -eq "")
        {
            $script:sidekick_job = Start-Job -ScriptBlock {
                
                $acros_found = $using:acro_index
                $editor_text = $using:editor

                [string]$editor_text = $editor_text.text
            
                $unique_acros = @{};
                $real_acro_list = @{};
                $total_acronyms = 0;
                $consistency_errors = @{};
                $metrics = @{};
                $script:Formating_errors = @{};
                $acro_list = @{};

                ##################################################################################
                ##Check Acronyms
                foreach($acro in $acros_found.getEnumerator() | sort key)
                {
                    ($mode,$index,$acronym,$meaning) = $acro.key -split "::"
                
                    ###########################################################################
                    ##Check to see if this is an acro within an acro
                    $acro_within_acro = 0;
                    if(($index -ne 0) -and ($index + $acronym.Length -le $editor_text.Length))
                    {
                        if(!(($editor_text).Substring(($index - 1),1) -match '[^A-Za-z0-9]'))
                        {
                            $acro_within_acro = 1;
                        }
                    }
                    ###########################################################################
                    ##Check consistency
                    $found = 0;
                    if($mode -eq "S")
                    {
                        foreach($acro_check in $acros_found.getEnumerator() | sort key)
                        {
                            ($mode2,$index2,$acronym2,$meaning2) = $acro_check.key -split "::"

                            if($acronym -eq $meaning2)
                            {
                                if(!($consistency_errors.contains($acronym)))
                                {
                                    $consistency_errors.add($acronym,$acronym2);
                                    #write-host $acro_within_acro = $acronym = $meaning   
                                }
                            }
                        }
                    }

                    ##################################################################################
                    ##Make EPR/Award List
                    if(!($real_acro_list.Contains("$mode::$acronym::$meaning")) -and ($acro_within_acro -ne 1))
                    {
                        $real_acro_list.Add("$mode::$acronym::$meaning",1);
                        ###########################################################################
                        ##Build Acro List
                        if($mode -eq "E")
                        {
                            $caps = $acronym -creplace '[^A-Z]'
                            $lowers = $acronym -creplace '[^a-z]'
                            if($caps.length -gt $lowers.length)
                            {
                                $meaning = (Get-Culture).TextInfo.ToTitleCase($meaning)
                                $meaning = $meaning -creplace "Of","of"
                                $meaning = $meaning -creplace "of$","Of"
                                $meaning = $meaning -creplace "The","the"
                                $meaning = $meaning -creplace "And","and"
                                $meaning = $meaning -creplace "  "," "
                                $meaning = $meaning.trim();

                                if((!($acro_list.contains($meaning))) -and ($acro_within_acro -ne 1))
                                {
                                    #write-host $acro_within_acro = $acronym = $meaning
                                    $acro_list.Add($meaning,$acronym);      
                                }
                            }

                        }
                        ###########################################################################
                    }
                    elseif($acro_within_acro -ne 1)
                    {
                        $real_acro_list["$mode::$acronym::$meaning"]++;
                        #write-host $acronym $real_acro_list[$acronym]
                    }
                }
                ##################################################################################
                ##Make Unique List
                $counter = 0 
                foreach($acro in $real_acro_list.getEnumerator() | sort value -descending)
                {
                    ($mode,$acronym,$meaning) = $acro.key -split '::'
                    if($mode -eq "E")
                    {
                        $counter++;
                        #write-host - $acronym = $acro.value
                        if((!($unique_acros.Contains($acronym))))
                        {
                            $unique_acros.add($acronym,$acro.value)
                            
                            $total_acronyms = $total_acronyms + $acro.value
                        }
                    }
                }
                ##################################################################################
                ##Find Metrics
                $pattern = "\d+.\d+|\d+"
                $matches = [regex]::Matches($editor_text, $pattern)
            
                if($matches.Success)
                {  
                    foreach($match in $matches)
                    {
                        #write-host ------------------------------------------------------------- $match.index = $match.value
                        $index = $match.index
                        $value = $match.value
                    

                        $spacer = "";
                        ############################################################
                        ##Append Before Number
                        if(($index -ne 0) -and ($editor_text.substring(($index - 1), 1) -match "\$|<|>|\#"))
                        {
                        
                            $value = ($editor_text.substring(($index - 1), 1)) + $value
                            $index = ($index - 1)
                        }
                        elseif(($index -ne 0) -and (!($editor_text.substring(($index - 1), 1) -match ' | | | |-|/|\.|;')))
                        {
                            #Eliminate Invalids
                            $value = ""
                        }
                        ############################################################
                        ##Append After Number
                        if((($index + $value.length + 2) -le $editor_text.length) -and ($editor_text.substring(($index + $value.length), 2) -match "st|th|k+|m+|t+|b+"))
                        {
                            if((($index + $value.length + 3) -le $editor_text.length) -and ($editor_text.substring(($index + $value.length + 2), 1) -match " | | | |-|`n"))
                            {
                                $value = $value + ($editor_text.substring(($index + $value.length), 2))
                            }
                            elseif(($index + $value.length + 2) -eq ($editor_text.length))
                            {
                                $value = $value + ($editor_text.substring(($index + $value.length), 2))
                            }

                        }
                        elseif((($index + $value.length + 1) -le $editor_text.length) -and ($editor_text.substring(($index + $value.length), 1) -match "%|-|k|m|t|b|x|\+"))
                        {
                            if($editor_text.substring(($index + $value.length + 1), 1) -match " | | | |-|\.|`n")
                            {
                                $value = $value + ($editor_text.substring(($index + $value.length), 1))
                            }
                        }
                        ###################
                        #Get Word After
                        $after = ""
                        if($editor_text.length -gt ($index + $value.length + 20))
                        {
                            $size = ($index + $value.length + 20)
                            $after = $editor_text.substring(($index + $value.length),20)
                        }
                        elseif($editor_text.length -eq $index + $value.length)
                        {
                            $after = "";
                        }
                        else
                        {
                            $after = $editor_text.substring(($index + $value.length),($editor_text.length - ($index + $value.length)));
                        }
                        ###########Split After
                        if($after)
                        {
                            $after_split = $after -split ' | | | |-|f/|w/|/|;|:|\.'
                            foreach($split in $after_split)
                            {
                
                                if($split -match '\n')
                                {
                                    $after = $split.trim();
                                    break;
                                }
                                $after = $split.trim();
                                if($after)
                                {
                                    break;
                                }
                            
                            }
                        }
                        if($value)
                        {
                            $key = "$index::$value"
                            if(!($metrics.contains($key)))
                            {
                                $metrics.add($key,$after);
                                #write-host $key = $after
                            }
                        }
                    }
                }
   
                #########################################################################################################
                #Text Disection
                $line_split = $editor_text -split '\n'
                $line_count = 0;
                $bullet_count = 0;
                $header_count = 0;
                $actions = @{};
                $sections = @{};
                $bullets = @{};
            
                foreach($line in $line_split)
                {
                    #write-host ------------------------------
                    $line_count++;
                    $line = $line.trim();
                    $line = $line -replace " | | "," "
                    $duplicate = 0;
                    if($line -ne "")
                    {
                        
                        if($line -match "^-")
                        {
                            #write-host BULLET
                            $bullet_count++;
                        
                            #########Add Bullets
                            #write-host $line
                            if(!($bullets.contains("$line")))
                            {
                                $bullets.add("$line",$line_count);
                            }
                            else
                            {
                                $duplicate = 1;
                                $script:Formating_errors.Add("Line $line_count (Duplicate Bullet)",$line_count);
                            }


                            if($duplicate -ne 1)
                            {
                                ########Split Bullet A.I.R.
                                $line = $line + "::E";   #Mark "Result" Section
                                $bullet_split = $line -split ";|--|\.\.\."
                                foreach($section in $bullet_split)
                                {

                                    ############Determine if this section is the end of the bullet
                                    $ending1 = 0;
                                    #write-host $section
                                    if($section -match '::E$')
                                    {
                                        $ending1 = 1;
                                        $section = $section -replace '::E$',''
                                    }
                                    if(!($sections.contains($section)))
                                    {
                                        ###Get First word in section
                                        $firstword1 = ((($section -replace '-','').trim() -split ' ')[0])
                                        $lastword1 = ((($section -replace '-|!','').trim() -split ' ')[-1])

                                        ###############################################################
                                        ##Cross refrence against all other sections
                                        foreach($other_section in $sections.getEnumerator())
                                        {
                                            $ending2 = 0;
                                            if($other_section.key -match '::E$')
                                            {
                                                $ending2 = 1;
                                                $other_section.key = $other_section.key -replace '::E$',''
                                            }

                                        
                                            $firstword2 = ((($other_section.key -replace '-','').trim() -split ' ')[0])
                                            $lastword2 = ((($other_section.key -replace '-|!','').trim() -split ' ')[-1])

                                            #write-host E1 = $ending1 $lastword1
                                            #write-host E2 = $ending2 $lastword2

                                            if($section -and $other_section.key)
                                            {
                                                if(($section -match "^-") -and ($other_section.key -match "^-") -and ($firstword1.tolower() -eq $firstword2.tolower()))
                                                {
                                                    $message = "Line $line_count (Repeated Opening `"$firstword1`")"
                                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                        
                                                }
                                                elseif($firstword1.tolower() -eq $firstword2.tolower())
                                                {
                                                    $message = "Line $line_count (Repeated Usage `"$firstword1`")"
                                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                                }
                                                if(($ending1 -eq 1) -and ($ending2 -eq 1) -and ($lastword1 -eq $lastword2))
                                                {
                                                    $message = "Line $line_count (Repeated Ending `"$lastword1`")"
                                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                                }
                                            }
                                        }
                                        ###Add to Section List
                                        if($ending1 -eq 0)
                                        {
                                            $sections.add($section,$line_count);
                                        }
                                        else
                                        {
                                            if(!($sections.contains("$section::E")))
                                            {
                                                $sections.add("$section::E",$line_count)
                                            }
                                            else
                                            {
                                                    $section = $section.trim();
                                                    $message = "Line $line_count (Repeated Section `"$section`")"
                                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                            }
                                        }
                                    }
                                    else
                                    {
                                        $section = $section.trim();
                                        $message = "Line $line_count (Repeated Section `"$section`")"
                                        if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                    }
                                }
                                ###################################################################################################
                                ###Missing Dash Space Start
                                if(!($line -match "^- "))
                                {
                                    $message = "Line $line_count (Missing Space After `"-`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                ###Missing ;
                                if(!($line -match ";"))
                                {
                                    ##Check for Exclamation instead
                                    $line_split = $line -split '--'
                                    if(!($line_split[0] -match '!'))
                                    {
                                        $script:Formating_errors.Add("Line $line_count (Missing `";`")",$line_count);
                                    }
                                }
                                elseif(!($line -match "; "))
                                {
                                    $message = "Line $line_count (Missing Space After `";`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                ###Multiple ;
                                $colon_count = ([regex]::Matches($line, ";" )).count
                                if($colon_count -gt 1)
                                {
                                    $message = "Line $line_count (Multiple `";`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                ###Missing --
                                if(!($line -match "--"))
                                {
                                    $message = "Line $line_count (Missing `"--`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                if($line -match " --")
                                {
                                    $message = "Line $line_count (Space Before `"--`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                if($line -match "-- ")
                                {
                                    $message = "Line $line_count (Space After `"--`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                ###Multiple --
                                $dash_count = ([regex]::Matches($line, "--" )).count
                                if($dash_count -gt 1)
                                {
                                    $message = "Line $line_count (Multiple `"--`")"
                                    if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                }
                                ##Extra Spaces
                                $matches = [regex]::Matches($line, "  ")
                                foreach($match in $matches)
                                {
                                    if($match.value -ne " ")
                                    {
                                        $message = "Line $line_count (Extra Spaces)"
                                        if(!($script:Formating_errors.contains($message))){$script:Formating_errors.Add($message,$line_count);}
                                    }
                                }
                            }
                        }
                        else
                        {
                            #write-host HEADER
                            $header_count++;
                        }  
                    }
                }
                #########################################################################################################
                #Word Usage
                $word_counter = @{};
                $simplified_text = ($editor_text -replace "[^a-z'-]| | | |--",' ')

                $text_array = $simplified_text -split ' '
                $word_count = 0;
                foreach($word in $text_array)
                {
                    if($word.length -gt 1)
                    {
                        $word_count++
                        if(!($word_counter.contains($word)) -and (!($real_acro_list.keys -match "E::$word")))
                        {
                            $word_counter.add($word,1);
                        }
                        elseif(!($real_acro_list.keys -match "E::$word"))
                        {
                            $word_counter[$word]++;
                        }
                    }
                }
      

                $result = @{"total_acronyms" = $($total_acronyms); "acro_list" = $($acro_list); "consistency_errors" = $($consistency_errors);"unique_acros" = $($unique_acros);"word_counter" = $($word_counter);"acro_counter" = $($real_acro_list); "header_count" = $($header_count); "bullet_count" = $($bullet_count); "word_count" = $($word_count);"formating_errors" = $($script:Formating_errors);"metrics" = $($metrics);}
                return ($result)
            }
        }
        else
        {   
            if($script:sidekick_job.state -eq "Completed")
            {   
                ($script:sidekick_results) = Receive-Job -Job $script:sidekick_job

                if($script:sidekickgui -match "Built|Update Values")
                {
                    $script:sidekickgui = "Update Values"
                }
                sidekick_display
                $script:sidekick_job = "";
            }
        }





    }#Sidekick width
}
#$script:formating_errors = @{};
################################################################################
######Sidekick Display##########################################################
function sidekick_display
{
    if($sidekick_panel.width -ne 5)
    {
        if(($script:sidekick_results -ne "") -and ($script:sidekickgui -eq "New"))
        {
            $script:sidekickgui = "Built"
            $left_panel.Controls.Clear();
            ################################################################################
            ######Sidekick Build Fresh Window###############################################
            $y_pos = 5

            ################################################################################
            ######Basic Info Label##########################################################
            $basic_info_label                   = New-Object system.Windows.Forms.Label
            $basic_info_label.text                     = "System Info";
            $basic_info_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $basic_info_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $basic_info_label.Anchor                   = 'top,right'
            $basic_info_label.width                    = ($sidekick_panel.width)
            $basic_info_label.height                   = 25
            $basic_info_label.TextAlign = "MiddleCenter"
            $basic_info_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
            $basic_info_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
            $left_panel.Controls.Add($basic_info_label)


            ################################################################################
            ######Package Name Label########################################################
            $y_pos = $y_pos + 25;
            $package_name_label                   = New-Object system.Windows.Forms.Label
            $package_name_label.text                     = "Package Name:";
            $package_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $package_name_label.Anchor                   = 'top,right'
            $package_name_label.autosize = $true
            $package_name_label.width                    = 120
            $package_name_label.height                   = 30
            $package_name_label.TextAlign = "MiddleLeft"
            $package_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $package_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($package_name_label);


            ################################################################################
            ######Package Name Value########################################################
            $script:package_name_value                   = New-Object system.Windows.Forms.Label
            $script:package_name_value.text                     = $script:settings['PACKAGE']
            $script:package_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:package_name_value.Anchor                   = 'top,right'
            $script:package_name_value.autosize = $true
            $script:package_name_value.TextAlign = "MiddleLeft"
            $script:package_name_value.width                    = 150
            $script:package_name_value.height                   = 30
            $script:package_name_value.location                 = New-Object System.Drawing.Point((10 + $package_name_label.width),$y_pos);
            $script:package_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:package_name_value);


            ################################################################################
            ######Bullets Loaded Label######################################################
            $y_pos = $y_pos + 25;
            $bullets_loaded_name_label                          = New-Object system.Windows.Forms.Label
            $bullets_loaded_name_label.text                     = "Bullets Loaded:";
            $bullets_loaded_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $bullets_loaded_name_label.Anchor                   = 'top,right'
            $bullets_loaded_name_label.autosize = $true
            $bullets_loaded_name_label.width                    = 120
            $bullets_loaded_name_label.height                   = 30
            $bullets_loaded_name_label.TextAlign = "MiddleLeft"
            $bullets_loaded_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $bullets_loaded_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($bullets_loaded_name_label);


            ################################################################################
            ######Bullets Loaded Value######################################################
            $script:bullets_loaded_name_value                          = New-Object system.Windows.Forms.Label
            $script:bullets_loaded_name_value.text                     = $Script:bullet_bank.Get_Count()
            $script:bullets_loaded_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:bullets_loaded_name_value.Anchor                   = 'top,right'
            $script:bullets_loaded_name_value.autosize = $true
            $script:bullets_loaded_name_value.TextAlign = "MiddleLeft"
            $script:bullets_loaded_name_value.width                    = 150
            $script:bullets_loaded_name_value.height                   = 30
            $script:bullets_loaded_name_value.location                 = New-Object System.Drawing.Point((10 + $bullets_loaded_name_label.width),$y_pos);
            $script:bullets_loaded_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:bullets_loaded_name_value);


            ################################################################################
            ######Acronyms Loaded Label#####################################################
            $y_pos = $y_pos + 25;
            $acronyms_loaded_name_label                   = New-Object system.Windows.Forms.Label
            $acronyms_loaded_name_label.text                     = "Acronyms Loaded:";
            $acronyms_loaded_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $acronyms_loaded_name_label.Anchor                   = 'top,right'
            $acronyms_loaded_name_label.autosize = $true
            $acronyms_loaded_name_label.width                    = 120
            $acronyms_loaded_name_label.height                   = 30
            $acronyms_loaded_name_label.TextAlign = "MiddleLeft"
            $acronyms_loaded_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $acronyms_loaded_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($acronyms_loaded_name_label);


            ################################################################################
            ######Acronyms Loaded Value#####################################################
            $script:acronyms_loaded_name_value                   = New-Object system.Windows.Forms.Label
            $script:acronyms_loaded_name_value.text                     = $script:acronym_list.Get_Count()
            $script:acronyms_loaded_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:acronyms_loaded_name_value.Anchor                   = 'top,right'
            $script:acronyms_loaded_name_value.autosize = $true
            $script:acronyms_loaded_name_value.TextAlign = "MiddleLeft"
            $script:acronyms_loaded_name_value.width                    = 150
            $script:acronyms_loaded_name_value.height                   = 30
            $script:acronyms_loaded_name_value.location                 = New-Object System.Drawing.Point((10 + $acronyms_loaded_name_label.width),$y_pos);
            $script:acronyms_loaded_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:acronyms_loaded_name_value);


            ################################################################################
            ######User Location Label#######################################################
            $y_pos = $y_pos + 25;
            $location_label                          = New-Object system.Windows.Forms.Label
            $location_label.text                     = "Current Line:";
            $location_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $location_label.Anchor                   = 'top,right'
            $location_label.autosize = $true
            $location_label.width                    = 120
            $location_label.height                   = 30
            $location_label.TextAlign = "MiddleLeft"
            $location_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $location_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######User Location Value#######################################################
            $script:location_value                   = New-Object system.Windows.Forms.Label
            $script:location_value.text                     = $script:current_line
            $script:location_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:location_value.Anchor                   = 'top,right'
            $script:location_value.autosize = $true
            $script:location_value.TextAlign = "MiddleLeft"
            $script:location_value.width                    = 150
            $script:location_value.height                   = 30
            $script:location_value.location                 = New-Object System.Drawing.Point((10 + $location_label.width),$y_pos);
            $script:location_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:location_value);
    

            ################################################################################
            ######Package Info Header#######################################################
            $y_pos = $y_pos + 35;
            $package_info_label                          = New-Object system.Windows.Forms.Label
            $package_info_label.text                     = "Package Information";
            $package_info_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $package_info_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $package_info_label.Anchor                   = 'top,right'
            $package_info_label.width                    = ($sidekick_panel.width)
            $package_info_label.height                   = 25
            $package_info_label.TextAlign = "MiddleCenter"
            $package_info_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
            $package_info_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4)) 
            $left_panel.Controls.Add($package_info_label)
            $left_panel.controls.Add($location_label);


            ################################################################################
            ######Header Count Label########################################################
            $y_pos = $y_pos + 25;
            $headers_count_name_label                          = New-Object system.Windows.Forms.Label
            $headers_count_name_label.text                     = "Headers:";
            $headers_count_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $headers_count_name_label.Anchor                   = 'top,right'
            $headers_count_name_label.autosize = $true
            $headers_count_name_label.width                    = 120
            $headers_count_name_label.height                   = 30
            $headers_count_name_label.TextAlign = "MiddleLeft"
            $headers_count_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $headers_count_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($headers_count_name_label);


            ################################################################################
            ######Header Count Value########################################################
            $script:headers_count_name_value                          = New-Object system.Windows.Forms.Label
            $script:headers_count_name_value.text                     = $script:sidekick_results.header_count
            $script:headers_count_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:headers_count_name_value.Anchor                   = 'top,right'
            $script:headers_count_name_value.autosize = $true
            $script:headers_count_name_value.TextAlign = "MiddleLeft"
            $script:headers_count_name_value.width                    = 150
            $script:headers_count_name_value.height                   = 30
            $script:headers_count_name_value.location                 = New-Object System.Drawing.Point((10 + $headers_count_name_label.width),$y_pos);
            $script:headers_count_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:headers_count_name_value);


            ################################################################################
            ######Bullets Count Label#######################################################
            $y_pos = $y_pos + 25;
            $bullets_count_name_label                          = New-Object system.Windows.Forms.Label
            $bullets_count_name_label.text                     = "Bullets:";
            $bullets_count_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $bullets_count_name_label.Anchor                   = 'top,right'
            $bullets_count_name_label.autosize = $true
            $bullets_count_name_label.width                    = 120
            $bullets_count_name_label.height                   = 30
            $bullets_count_name_label.TextAlign = "MiddleLeft"
            $bullets_count_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $bullets_count_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($bullets_count_name_label);


            ################################################################################
            ######Bullets Count Value#######################################################
            $script:bullets_count_name_value                          = New-Object system.Windows.Forms.Label
            $script:bullets_count_name_value.text                     = $script:sidekick_results.bullet_count
            $script:bullets_count_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:bullets_count_name_value.Anchor                   = 'top,right'
            $script:bullets_count_name_value.autosize = $true
            $script:bullets_count_name_value.TextAlign = "MiddleLeft"
            $script:bullets_count_name_value.width                    = 150
            $script:bullets_count_name_value.height                   = 30
            $script:bullets_count_name_value.location                 = New-Object System.Drawing.Point((10 + $bullets_count_name_label.width),$y_pos);
            $script:bullets_count_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:bullets_count_name_value);


            ################################################################################
            ######Word Count Label##########################################################
            $y_pos = $y_pos + 25;
            $word_count_name_label                          = New-Object system.Windows.Forms.Label
            $word_count_name_label.text                     = "Words:";
            $word_count_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $word_count_name_label.Anchor                   = 'top,right'
            $word_count_name_label.autosize = $true
            $word_count_name_label.width                    = 120
            $word_count_name_label.height                   = 30
            $word_count_name_label.TextAlign = "MiddleLeft"
            $word_count_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $word_count_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($word_count_name_label);


            ################################################################################
            ######Word Count Value##########################################################
            $script:word_count_name_value                   = New-Object system.Windows.Forms.Label
            $script:word_count_name_value.text                     = $script:sidekick_results.word_count[0]
            $script:word_count_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:word_count_name_value.Anchor                   = 'top,right'
            $script:word_count_name_value.autosize = $true
            $script:word_count_name_value.TextAlign = "MiddleLeft"
            $script:word_count_name_value.width                    = 150
            $script:word_count_name_value.height                   = 30
            $script:word_count_name_value.location                 = New-Object System.Drawing.Point((10 + $word_count_name_label.width),$y_pos);
            $script:word_count_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:word_count_name_value);


            ################################################################################
            ######Unique Acros Label########################################################
            $y_pos = $y_pos + 25;
            $unique_acro_count_name_label                          = New-Object system.Windows.Forms.Label
            $unique_acro_count_name_label.text                     = "Unique Acronyms:";
            $unique_acro_count_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $unique_acro_count_name_label.Anchor                   = 'top,right'
            $unique_acro_count_name_label.autosize = $true
            $unique_acro_count_name_label.width                    = 120
            $unique_acro_count_name_label.height                   = 30
            $unique_acro_count_name_label.TextAlign = "MiddleLeft"
            $unique_acro_count_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $unique_acro_count_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($unique_acro_count_name_label);


            ################################################################################
            ######Unique Acros Value########################################################
            $script:unique_acro_count_name_value                   = New-Object system.Windows.Forms.Label
            $script:unique_acro_count_name_value.text                     = $script:sidekick_results.unique_acros.get_count();
            $script:unique_acro_count_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:unique_acro_count_name_value.Anchor                   = 'top,right'
            $script:unique_acro_count_name_value.autosize = $true
            $script:unique_acro_count_name_value.TextAlign = "MiddleLeft"
            $script:unique_acro_count_name_value.width                    = 150
            $script:unique_acro_count_name_value.height                   = 30
            $script:unique_acro_count_name_value.location                 = New-Object System.Drawing.Point((10 + $unique_acro_count_name_label.width),$y_pos);
            $script:unique_acro_count_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:unique_acro_count_name_value);


            ################################################################################
            ######Total Acros Label#########################################################
            $y_pos = $y_pos + 25;
            $total_acro_count_name_label                          = New-Object system.Windows.Forms.Label
            $total_acro_count_name_label.text                     = "Total Acronyms:";
            $total_acro_count_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $total_acro_count_name_label.Anchor                   = 'top,right'
            $total_acro_count_name_label.autosize = $true
            $total_acro_count_name_label.width                    = 120
            $total_acro_count_name_label.height                   = 30
            $total_acro_count_name_label.TextAlign = "MiddleLeft"
            $total_acro_count_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $total_acro_count_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($total_acro_count_name_label);


            ################################################################################
            ######Total Acros Value#########################################################
            $script:total_acro_count_name_value                   = New-Object system.Windows.Forms.Label
            $script:total_acro_count_name_value.text                     = $script:sidekick_results.total_acronyms;
            $script:total_acro_count_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
            $script:total_acro_count_name_value.Anchor                   = 'top,right'
            $script:total_acro_count_name_value.autosize = $true
            $script:total_acro_count_name_value.TextAlign = "MiddleLeft"
            $script:total_acro_count_name_value.width                    = 150
            $script:total_acro_count_name_value.height                   = 30
            $script:total_acro_count_name_value.location                 = New-Object System.Drawing.Point((10 + $total_acro_count_name_label.width),$y_pos);
            $script:total_acro_count_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
            $left_panel.controls.Add($script:total_acro_count_name_value);


            ################################################################################
            ######Errors Info Header########################################################
            $y_pos = $y_pos + 35;
            $errors_info_label                          = New-Object system.Windows.Forms.Label
            $errors_info_label.text                     = "Errors";
            $errors_info_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $errors_info_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $errors_info_label.Anchor                   = 'top,right'
            $errors_info_label.width                    = ($sidekick_panel.width)
            $errors_info_label.height                   = 25
            $errors_info_label.TextAlign = "MiddleCenter"
            $errors_info_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
            $errors_info_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
            $left_panel.Controls.Add($errors_info_label)


            ################################################################################
            ######Formating Errors Label####################################################
            $y_pos = $y_pos + 25;
            $formating_errors_label                          = New-Object system.Windows.Forms.Label
            $formating_errors_label.text                     = "Formating Errors:";
            $formating_errors_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $formating_errors_label.Anchor                   = 'top,right'
            $formating_errors_label.autosize = $true
            $formating_errors_label.width                    = 120
            $formating_errors_label.height                   = 20
            $formating_errors_label.TextAlign = "MiddleLeft"
            $formating_errors_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $formating_errors_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######Formating Errors Combo####################################################
            $y_pos = $y_pos + 30;
            $script:formating_errors_combo                   = New-Object System.Windows.Forms.ComboBox	
            $script:formating_errors_combo.Items.Clear();
            $script:formating_errors_combo.width = ($left_panel.width - 10)
            $script:formating_errors_combo.autosize = $false
            $script:formating_errors_combo.Anchor = 'top,right'
            $script:formating_errors_combo.Location = New-Object System.Drawing.Point(5,$y_pos)
            $script:formating_errors_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
            $script:formating_errors_combo.DropDownStyle = "DropDownList"
            $script:formating_errors_combo.AccessibleName = "On";
            $first = "";
            foreach($format_error in $script:sidekick_results.formating_errors.getEnumerator() | sort value)
            {
                if($first -eq "")
                {
                    $first = $format_error.key
                }
                $script:formating_errors_combo.Items.Add($format_error.key); 
            }
            $script:formating_errors_combo.SelectedItem = $first
            $script:formating_errors_combo.Add_SelectedValueChanged({
                if($this.AccessibleName -eq "On")
                {
                    $editor.SelectAll();
                    $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                    $editor.DeselectAll()

                    $pattern = ""
                    $applies = "ALL"
                    if($this.SelectedItem -match "Duplicate"){ $pattern = ""; $applies = "Duplicate"}
                    if($this.SelectedItem -match "Missing Space After `"-`""){ $pattern = "^-[A-Za-z0-9]|`n-[A-Za-z0-9]"}
                    if($this.SelectedItem -match "Missing Space After `";`""){ $pattern = ";[A-Za-z0-9]"}
                    if($this.SelectedItem -match "Multiple `";`""){ $pattern = ";"; $applies = "Line"}
                    if($this.SelectedItem -match "Multiple `"--`""){ $pattern = "--"; $applies = "Line"}
                    if($this.SelectedItem -match "Space Before `"--`""){ $pattern = " --"; $applies = "Spaces"}
                    if($this.SelectedItem -match "Space After `"--`""){ $pattern = "-- "; $applies = "Spaces"} 
                    if($this.SelectedItem -match "Extra Spaces"){ $pattern = "  "; $applies = "Spaces"}
                    if($this.SelectedItem -match "Missing"){ $pattern = ""; $applies = "Missing"}
                    if($this.SelectedItem -match "Repeated Section|Repeated Usage|Repeated Opening|Repeated Ending"){ 
                        $section = $this.SelectedItem|%{$_.split('"')[1]}
                        $section = $section.substring(0,1).tolower() + $section.substring(1) 
                        $section2 = $section.substring(0,1).toupper() + $section.substring(1) 
                        $pattern = "$section|$section2"
                    }

                    $scope_start = 0;
                    $scope_end = $editor.text.length;
                    $matches = "";

                    if($applies -eq "ALL")
                    {
                        $matches = [regex]::Matches($editor.text, $pattern)
                    }
                    else #Applies to a line
                    {
                            
                        ($item_split) = $this.SelectedItem -split ' '; 
                        $line_split = $editor.text -split '\n'
                        $line_count = 0;
                        $found = "";
                            
                        foreach($line in $line_split)
                        {
                            $line_count++;
                            if($line_count -eq [int]$item_split[1])
                            {
                                $found = $line;
                                $scope_end = $scope_start + $line.length
                                #write-host found it
                                break;
                            }
                            #write-host Looking
                            $scope_start = $scope_start + $line.length + 1;
                        }
                        if(($applies -eq "Line") -or ($applies -eq "Spaces"))
                        {
                            $matches = [regex]::Matches($editor.text, $pattern)
                        }
                        elseif($applies -eq "Duplicate")
                        {
                            $scope_start = 0;
                            $scope_end = $editor.text.length;
                            $found = $found.trim();
                            $pattern = "$([regex]::escape($found))"
                            $matches = [regex]::Matches($editor.text,$pattern)
                        }
                        else #Missing
                        {
                            
                            $editor.SelectionStart = [int]$scope_start
                            $editor.SelectionLength = $found.length
                            $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                            $editor.DeselectAll();
                        }
                        #write-host $scope_start = $scope_end - $found
                    }
                        
                    if($matches)
                    {
                        foreach($match in $matches)
                        {
                            
                            #write-host Match Found $match.index - $scope_start - $scope_end
                            if(($match.index -ge $scope_start)  -and (($match.index + $match.value.length) -le $scope_end))
                            {
                                $before = " ";
                                $after = "";
                                if($match.index -ne 0)
                                {
                                    $before = $editor.text.substring(($match.index -1),1);
                                }
                                if(($before -match "-|;|\s|/|\.") -or ($applies -eq "Spaces"))
                                {
                                    if(($match.index + $match.value.length + 1) -lt $editor.text.Length)
                                    {
                                        $after = $editor.text.substring(($match.index + $match.value.length),1);
                                    }
                                    if(!($after -match "\w")  -or ($applies -eq "Spaces"))
                                    {
                                        $editor.SelectionStart = $match.index
                                        $editor.SelectionLength = $match.value.length
                                        $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                                        $editor.DeselectAll();
                                    }
                                }
                            }
                        }
                    }
                }      
            })
            $left_panel.controls.Add($script:formating_errors_combo);
            $left_panel.controls.Add($formating_errors_label);


            ################################################################################
            ######Consistency Errors Label##################################################
            $y_pos = $y_pos + 25;
            $consistency_errors_label                          = New-Object system.Windows.Forms.Label
            $consistency_errors_label.text                     = "Consistency Errors:";
            $consistency_errors_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $consistency_errors_label.Anchor                   = 'top,right'
            $consistency_errors_label.autosize = $true
            $consistency_errors_label.width                    = 120
            $consistency_errors_label.height                   = 30
            $consistency_errors_label.TextAlign = "MiddleLeft"
            $consistency_errors_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $consistency_errors_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######Consistency Errors Combo##################################################
            $y_pos = $y_pos + 30;
            $script:consistency_errors_combo = New-Object System.Windows.Forms.ComboBox	
            $script:consistency_errors_combo.Items.Clear();
            $script:consistency_errors_combo.width = ($left_panel.width - 10)
            $script:consistency_errors_combo.autosize = $false
            $script:consistency_errors_combo.Anchor = 'top,right'
            $script:consistency_errors_combo.Location = New-Object System.Drawing.Point(5,$y_pos)
            $script:consistency_errors_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
            $script:consistency_errors_combo.DropDownStyle = "DropDownList"
            $script:consistency_errors_combo.AccessibleName = "On";
            $first = "";
            foreach($consistency_error in $script:sidekick_results.consistency_errors.getEnumerator())
            {
                $value = "Using both: " + $consistency_error.key + " & " + $consistency_error.value
                if($first -eq "")
                {
                    $first = $value
                }
                $script:consistency_errors_combo.Items.Add($value); 
            }
            $script:consistency_errors_combo.SelectedItem = $first
            $script:consistency_errors_combo.Add_SelectedValueChanged({
                if($this.AccessibleName -eq "On")
                {
                    $value = $this.SelectedItem -replace "Using both: ",""
                    ($word1,$word2) = $value -split ' & '
                    $word1c = $word1.substring(0,1).toupper() + $word1.substring(1)  
                    $word2c = $word2.substring(0,1).toupper() + $word2.substring(1)
                    $word1l = $word1.substring(0,$word1.Length).toupper() 
                    $word2l = $word2.substring(0,$word2.Length).toupper()
                    $culture_word1 = (Get-Culture).TextInfo.ToTitleCase($word1)
                    $culture_word2 = (Get-Culture).TextInfo.ToTitleCase($word2)

                    $simplified_text = $editor.text -replace " | | ",' ' 
                    $pattern = "$([regex]::escape($word1))|$([regex]::escape($word2))|$([regex]::escape($word1c))|$([regex]::escape($word2c))|$([regex]::escape($word1l))|$([regex]::escape($word2l))|$([regex]::escape($culture_word1))|$([regex]::escape($culture_word2))"
                    $matches = [regex]::Matches($simplified_text, $pattern)
                    
                    $editor.SelectAll();
                    $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                    $editor.DeselectAll()

                    if($matches.Success)
                    {  
                        foreach($match in $matches)
                        {
                            if(((($match.index - 1) -ge 0 ) -and (!($editor.text.Substring(($match.index - 1),1) -match "\w"))) -or (($match.index -1) -lt 0))
                            {
                                if(((($match.index + $match.value.length + 1) -le $editor.text.length) -and (!($editor.text.Substring(($match.index + $match.value.length),1) -match "\w"))) -or (($match.index + $match.value.length) -eq $editor.text.length) -or ($editor.text.Substring(($match.index + $match.value.length - 1),1) -match '/'))
                                {
                                    $editor.SelectionStart = $match.index
                                    $editor.SelectionLength = $match.value.length
                                    $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                                    $editor.DeselectAll()
                                }
                            }
                        }
                    }
                }
                     
            })
            $left_panel.controls.Add($script:consistency_errors_combo);
            $left_panel.controls.Add($consistency_errors_label);


            ################################################################################
            ######Analytics Header##########################################################
            $y_pos = $y_pos + 35;
            $analytics_label                          = New-Object system.Windows.Forms.Label
            $analytics_label.text                     = "Analytics";
            $analytics_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $analytics_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $analytics_label.Anchor                   = 'top,right'
            $analytics_label.width                    = ($sidekick_panel.width)
            $analytics_label.height                   = 25
            $analytics_label.TextAlign = "MiddleCenter"
            $analytics_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
            $analytics_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
            $left_panel.Controls.Add($analytics_label)


            ################################################################################
            ######Top Acros Label###########################################################
            $y_pos = $y_pos + 25;
            $top_used_acros_label                          = New-Object system.Windows.Forms.Label
            $top_used_acros_label.text                     = "Top 20 Acronyms:";
            $top_used_acros_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $top_used_acros_label.Anchor                   = 'top,right'
            $top_used_acros_label.autosize = $true
            $top_used_acros_label.width                    = 120
            $top_used_acros_label.height                   = 30
            $top_used_acros_label.TextAlign = "MiddleLeft"
            $top_used_acros_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $top_used_acros_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######Top Acros Combo###########################################################
            $y_pos = $y_pos + 30;
            $script:top_used_acros_combo = New-Object System.Windows.Forms.ComboBox	
            $script:top_used_acros_combo.Items.Clear();
            $script:top_used_acros_combo.width = ($left_panel.width - 10)
            $script:top_used_acros_combo.autosize = $false
            $script:top_used_acros_combo.Anchor = 'top,right'
            $script:top_used_acros_combo.Location = New-Object System.Drawing.Point(5,$y_pos)
            $script:top_used_acros_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
            $script:top_used_acros_combo.DropDownStyle = "DropDownList"
            $script:top_used_acros_combo.AccessibleName = "On";
            $first = "";   
            foreach($acro in $script:sidekick_results.unique_acros.getEnumerator() | sort value -descending | Select-Object -First 20)
            {

                [string]$value = [string]$acro.value + "x - " + $acro.key
                if($first -eq "")
                {
                    $first = $value
                }
                $script:top_used_acros_combo.Items.Add($value); 
            }
            $script:top_used_acros_combo.SelectedItem = $first
            $script:top_used_acros_combo.Add_SelectedValueChanged({
                if($this.AccessibleName -eq "On")
                {
                    $value = $this.SelectedItem -replace "Using both: ",""
                    ($trash,$word1) = $value -split ' - '
                    $word2 = $word1;
                    $word2 = $word1.substring(0,1).toupper() + $word1.substring(1)  


                    $pattern = "$([regex]::escape($word1))|$([regex]::escape($word2))"
                    $matches = [regex]::Matches($editor.text, $pattern)
                    
                    $editor.SelectAll();
                    $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                    $editor.DeselectAll()

                    if($matches.Success)
                    {  
                        foreach($match in $matches)
                        {
                            if(((($match.index - 1) -ge 0 ) -and (!($editor.text.Substring(($match.index - 1),1) -match "[A-Z0-9]"))) -or (($match.index -1) -lt 0))
                            {
                                if(((($match.index + $match.value.length + 1) -le $editor.text.length ) -and (!($editor.text.Substring(($match.index + $match.value.length),1) -match "[A-Z0-9]"))) -or (($match.index + $match.value.length) -eq $editor.text.length) -or ($editor.text.Substring(($match.index + $match.value.length - 1),1) -match '/'))
                                {
                                    #write-host $match.index - $match.value
                                    $editor.SelectionStart = $match.index
                                    $editor.SelectionLength = $match.value.length
                                    $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                                    $editor.DeselectAll()
                                }
                            }
                        }
                    }
                }
                     
            })
            $left_panel.controls.Add($script:top_used_acros_combo);
            $left_panel.controls.Add($top_used_acros_label);
            

            ################################################################################
            ######Top Used Words Label######################################################
            $y_pos = $y_pos + 25;
            $top_used_words_label                          = New-Object system.Windows.Forms.Label
            $top_used_words_label.text                     = "Top 20 Words:";
            $top_used_words_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $top_used_words_label.Anchor                   = 'top,right'
            $top_used_words_label.autosize = $true
            $top_used_words_label.width                    = 120
            $top_used_words_label.height                   = 30
            $top_used_words_label.TextAlign = "MiddleLeft"
            $top_used_words_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $top_used_words_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######Top Used Words Combo######################################################
            $y_pos = $y_pos + 30;
            $script:top_used_words_combo = New-Object System.Windows.Forms.ComboBox	
            $script:top_used_words_combo.Items.Clear();
            $script:top_used_words_combo.width = ($left_panel.width - 10)
            $script:top_used_words_combo.autosize = $false
            $script:top_used_words_combo.Anchor = 'top,right'
            $script:top_used_words_combo.Location = New-Object System.Drawing.Point(5,$y_pos)
            $script:top_used_words_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
            $script:top_used_words_combo.DropDownStyle = "DropDownList"
            $script:top_used_words_combo.AccessibleName = "On";
            $first = "";
            foreach($word in $script:sidekick_results.word_counter.getEnumerator() | sort value -descending | Select-Object -First 20)
            {
                [string]$value = [string]$word.value + "x - " + $word.key
                if($first -eq "")
                {
                    $first = $value
                }
                $script:top_used_words_combo.Items.Add($value); 
            }
            $script:top_used_words_combo.SelectedItem = $first
            $script:top_used_words_combo.Add_SelectedValueChanged({
                if($this.AccessibleName -eq "On")
                {
                    $value = $this.SelectedItem -replace "Using both: ",""
                    ($trash,$word1) = $value -split ' - '
                    $word2 = $word1;
                    $word2 = $word1.substring(0,1).toupper() + $word1.substring(1)
                    $word3 = $word1.substring(0,1).tolower() + $word1.substring(1) 

                    #write-host $word1 = $word2

                    $pattern = "$([regex]::escape($word1))|$([regex]::escape($word2))|$([regex]::escape($word3))"
                    $matches = [regex]::Matches($editor.text, $pattern)
                    
                    $editor.SelectAll();
                    $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                    $editor.DeselectAll()

                    if($matches.Success)
                    {  
                        foreach($match in $matches)
                        {
                            if(((($match.index - 1) -ge 0 ) -and (!($editor.text.Substring(($match.index - 1),1) -match "[A-Z0-9]"))) -or (($match.index -1) -lt 0))
                            {
                                if(((($match.index + $match.value.length + 1) -le $editor.text.length ) -and (!($editor.text.Substring(($match.index + $match.value.length),1) -match "[A-Z0-9]"))) -or (($match.index + $match.value.length) -eq $editor.text.length))
                                {
                                    #write-host $match.index - $match.value
                                    $editor.SelectionStart = $match.index
                                    $editor.SelectionLength = $match.value.length
                                    $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                                    $editor.DeselectAll()
                                }
                            }
                        }
                    }
                }        
            })
            $left_panel.controls.Add($script:top_used_words_combo);
            $left_panel.controls.Add($top_used_words_label);


            ################################################################################
            ######Metrics Used Label########################################################
            $y_pos = $y_pos + 25;
            $metrics_used_label                          = New-Object system.Windows.Forms.Label
            $metrics_used_label.text                     = "Metrics Used:";
            $metrics_used_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            $metrics_used_label.Anchor                   = 'top,right'
            $metrics_used_label.autosize = $true
            $metrics_used_label.width                    = 120
            $metrics_used_label.height                   = 30
            $metrics_used_label.TextAlign = "MiddleLeft"
            $metrics_used_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
            $metrics_used_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            

            ################################################################################
            ######Metrics Used Combo########################################################
            $y_pos = $y_pos + 30;
            $script:metrics_used_combo = New-Object System.Windows.Forms.ComboBox	
            $script:metrics_used_combo.Items.Clear();
            $script:metrics_used_combo.width = ($left_panel.width - 10)
            $script:metrics_used_combo.autosize = $false
            $script:metrics_used_combo.Anchor = 'top,right'
            $script:metrics_used_combo.Location = New-Object System.Drawing.Point(5,$y_pos)
            $script:metrics_used_combo.font = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE']))
            $script:metrics_used_combo.DropDownStyle = "DropDownList"
            $script:metrics_used_combo.AccessibleName = "On";
            $first = "";
            foreach($metric in $script:sidekick_results.metrics.getEnumerator() | sort value)
            {
                $metric_split = $metric.key -split '::'
                $value = [string]$metric_split[1] + " " + [string]$metric.value + "                                                                                                      ::" + $metric_split[0];
                if($first -eq "")
                {
                    $first = $value
                }
                $script:metrics_used_combo.Items.Add($value); 
            }
            $script:metrics_used_combo.SelectedItem = $first
            $script:metrics_used_combo.Add_SelectedValueChanged({
                if($this.AccessibleName -eq "On")
                {
                    ($metric,$location) = $this.SelectedItem -split "::"

                    $metric = $metric.trim();
                    [int]$location = $location.trim();

                    $editor.SelectAll();
                    $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                    $editor.DeselectAll()

                    $editor.SelectionStart = $location
                    $editor.SelectionLength = $metric.length
                    $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                    $editor.DeselectAll()


                }
                     
            })
            $left_panel.controls.Add($script:metrics_used_combo);
            $left_panel.controls.Add($metrics_used_label);

            ################################################################################
            ######Tools Header##############################################################
            $y_pos = $y_pos + 45;
            $tools_label                          = New-Object system.Windows.Forms.Label
            $tools_label.text                     = "Tools";
            $tools_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
            $tools_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
            $tools_label.Anchor                   = 'top,right'
            $tools_label.width                    = ($sidekick_panel.width)
            $tools_label.height                   = 30
            $tools_label.TextAlign = "MiddleCenter"
            $tools_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
            $tools_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
            $left_panel.Controls.Add($tools_label)


            ################################################################################
            ######EPR Acros Button##########################################################
            $y_pos = $y_pos + 35;
            $gen_epr_acros_button           = New-Object System.Windows.Forms.Button
            $gen_epr_acros_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $gen_epr_acros_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $gen_epr_acros_button.Width     = 200
            $gen_epr_acros_button.height    = 25
            $gen_epr_acros_button.location  = New-Object System.Drawing.Point((($left_panel.width / 2) - ($gen_epr_acros_button.Width / 2)),$y_pos)
            $gen_epr_acros_button.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $gen_epr_acros_button.Text      ="EPR Acronym List"
            $gen_epr_acros_button.Name = ""
            $gen_epr_acros_button.Add_Click({ 
                acronym_list_dialog "EPR"
            })
            $left_panel.controls.Add($gen_epr_acros_button)


            ################################################################################
            ######Award Acros Button########################################################
            $y_pos = $y_pos + 30;
            $gen_1206_acros_button           = New-Object System.Windows.Forms.Button
            $gen_1206_acros_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
            $gen_1206_acros_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
            $gen_1206_acros_button.Width     = 200
            $gen_1206_acros_button.height    = 25
            $gen_1206_acros_button.location  = New-Object System.Drawing.Point((($left_panel.width / 2) - ($gen_1206_acros_button.Width / 2)),$y_pos)
            $gen_1206_acros_button.Font      = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
            $gen_1206_acros_button.Text      ="Award Acronym List"
            $gen_1206_acros_button.Name = ""
            $gen_1206_acros_button.Add_Click({ 
                acronym_list_dialog "Award"
            })
            $left_panel.controls.Add($gen_1206_acros_button)


            ################################################################################
            ######Find Words Input##########################################################
            $y_pos = $y_pos + 30;
            $find_input                   = New-Object system.Windows.Forms.TextBox                       
            $find_input.AutoSize                 = $true
            $find_input.ForeColor                = $script:theme_settings['DIALOG_INPUT_TEXT_COLOR']
            $find_input.BackColor                = $script:theme_settings['DIALOG_INPUT_BACKGROUND_COLOR']
            $find_input.Anchor                   = 'top,left'
            $find_input.width                    = 200
            $find_input.height                   = 30
            $find_input.location                 = New-Object System.Drawing.Point((($left_panel.width / 2) - ($gen_1206_acros_button.Width / 2)),$y_pos)
            $find_input.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 2))
            $find_input.text                     = "Find..."
            $find_input.Add_GotFocus({
                if($this.text -eq "Find...")
                {
                    $this.text = ""
                }
            })
            $find_input.Add_LostFocus({
                if($this.text -eq "")
                {
                    $this.text = "Find..."
                }
            })
            $find_input.Add_TextChanged({
                $value = $this.text.tolower()

                $pattern = "$([regex]::escape($value))"
                $matches = [regex]::Matches($editor.text.tolower(), $pattern)
                    
                $editor.SelectAll();
                $editor.selectionbackcolor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
                $editor.DeselectAll()

                if($matches.Success)
                {  
                    foreach($match in $matches)
                    {
                        if(((($match.index - 1) -ge 0 ) -and (!($editor.text.Substring(($match.index - 1),1) -match "[A-Z0-9]"))) -or (($match.index -1) -lt 0))
                        {
                            $editor.SelectionStart = $match.index
                            $editor.SelectionLength = $match.value.length
                            $editor.selectionbackcolor = [System.Drawing.ColorTranslator]::FromHtml($script:theme_settings['EDITOR_HIGHLIGHT_COLOR'])
                            $editor.DeselectAll();
                        }
                    }
                }     
            })
            $left_panel.controls.Add($find_input);


            ################################################################################
            ######Compression Label#########################################################
            $y_pos = $y_pos + 45;
            $compression_label                          = New-Object system.Windows.Forms.Label
            $compression_label.text                     = "Text Compression";
            $compression_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            #$compression_label.backcolor = "green"
            $compression_label.Anchor                   = 'top,left'
            #$compression_label.autosize = $true
            $compression_label.width                    = 200
            $compression_label.height                   = 30
            $compression_label.TextAlign = "MiddleCenter"
            $compression_label.location                 = New-Object System.Drawing.Point((($left_panel.width / 2) - ($compression_label.Width / 2)),$y_pos)
            $compression_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $left_panel.controls.Add($compression_label);


            ################################################################################
            ######Compression Trackbar######################################################
            $y_pos = $y_pos + 30;
            $script:compression_trackbar_label = New-Object System.Windows.Forms.Label
            $compression_trackbar = New-Object System.Windows.Forms.TrackBar
            $compression_trackbar.Width = 200
            $compression_trackbar.Location = New-Object System.Drawing.Point((($left_panel.width / 2) - ($compression_trackbar.Width / 2)),$y_pos)
            $compression_trackbar.Orientation = "Horizontal"
            $compression_trackbar.Height = 40
            $compression_trackbar.TickFrequency = 1
            $compression_trackbar.TickStyle = "TopLeft"
            $compression_trackbar.SetRange(1, 5) 
            $compression_trackbar.AccessibleName = "Off"
            $compression_trackbar.add_ValueChanged({
                $script:settings['TEXT_COMPRESSION'] = $this.value
                if($this.value -eq 1)
                {
                    $script:space_hash.Clear();
                    $this.AccessibleName = "Off"
                    $script:compression_trackbar_label.Text = $this.AccessibleName
                    $script:space_hash.Clear();
                    $script:bullets_compressed = new-object System.Collections.Hashtable
                }
                if($this.value -eq 2)
                {
                    $script:space_hash.Clear();
                    $script:bullets_compressed = new-object System.Collections.Hashtable
                    $this.AccessibleName = "Reset"
                    $script:compression_trackbar_label.Text = $this.AccessibleName
                    $script:space_hash.Clear();
                    #$script:space_hash.add(" ",14.2159411269359)
                }
                elseif($this.value -eq 3)
                {
                    $this.AccessibleName = "Low"
                    $script:compression_trackbar_label.Text = $this.AccessibleName
                    $script:space_hash.Clear();
                    $script:bullets_compressed = new-object System.Collections.Hashtable
                    #$script:space_hash.add(" ",14.2159411269359)
                    $script:space_hash.add(" ",11.3886113886114)
                }
                elseif($this.value -eq 4)
                {
                    $this.AccessibleName = "Medium"
                    $script:compression_trackbar_label.Text = $this.AccessibleName
                    $script:space_hash.Clear();
                    $script:bullets_compressed = new-object System.Collections.Hashtable
                    #$script:space_hash.add(" ",14.2159411269359)
                    $script:space_hash.add(" ",11.3886113886114)
                    $script:space_hash.add(" ",9.47735191637631)
                }
                elseif($this.value -eq 5)
                {
                    $this.AccessibleName = "High"
                    $script:compression_trackbar_label.Text = $this.AccessibleName
                    $script:space_hash.Clear();
                    $script:bullets_compressed = new-object System.Collections.Hashtable
                    #$script:space_hash.add(" ",14.2159411269359)
                    $script:space_hash.add(" ",11.3886113886114)
                    $script:space_hash.add(" ",9.47735191637631)
                    $script:space_hash.add(" ",4.75524475524475)
                }
                $Script:recent_editor_text = "Changed"
                update_settings
            })
            $compression_trackbar.Value = $script:settings['TEXT_COMPRESSION']
            

            ################################################################################
            ######Compression Trackbar Label################################################
            $y_pos = $y_pos + 35;   
            $script:compression_trackbar_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
            $script:compression_trackbar_label.width = 200
            $script:compression_trackbar_label.Location = New-Object System.Drawing.Point((($left_panel.width / 2) - ($script:compression_trackbar_label.Width / 2)),$y_pos)
            $script:compression_trackbar_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
            #$script:compression_trackbar_label.BackColor = "Green"
            $script:compression_trackbar_label.TextAlign = "MiddleCenter"
            $script:compression_trackbar_label.Text = $compression_trackbar.AccessibleName
            $left_panel.Controls.Add($script:compression_trackbar_label)
            $left_panel.controls.Add($compression_trackbar)


        } 

###############################################################################################################################################################
######Sidekick Refresh Information#############################################################################################################################
        if((Test-Path variable:script:left_panel) -and (Test-Path variable:script:sidekick_results) -and ($script:sidekickgui -eq "Update Values"))
        {
            ################################################################################
            ######Update Sidekick Vars (No Rebuild)#########################################
            $script:location_value.text                     = $script:current_line
            $script:package_name_value.text                 = $script:settings['PACKAGE']
            $script:bullets_loaded_name_value.text          = $Script:bullet_bank.Get_Count()
            $script:acronyms_loaded_name_value.text         = $script:acronym_list.Get_Count()
            $script:headers_count_name_value.text           = $script:sidekick_results.header_count
            $script:bullets_count_name_value.text           = $script:sidekick_results.bullet_count
            $script:word_count_name_value.text              = $script:sidekick_results.word_count
            $script:unique_acro_count_name_value.text       = $script:sidekick_results.unique_acros.get_count();
            $script:total_acro_count_name_value.text        = $script:sidekick_results.total_acronyms;


            ################################################################################
            ######Update Sidekick Formating Errors Combo (No Rebuild)#######################
            $script:formating_errors_combo.AccessibleName = "Off";
            $script:formating_errors_combo.Items.Clear();
            $first = "";
            foreach($format_error in $script:sidekick_results.formating_errors.getEnumerator() | sort value)
            {
                if($first -eq "")
                {
                    $first = $format_error.key
                }
                $script:formating_errors_combo.Items.Add($format_error.key); 
            }
            $script:formating_errors_combo.SelectedItem = $first
            $script:formating_errors_combo.AccessibleName = "On";


            ################################################################################
            ######Update Sidekick Consistency Errors Combo (No Rebuild)#####################
            $script:consistency_errors_combo.AccessibleName = "Off";
            $script:consistency_errors_combo.Items.Clear();
            $first = "";
            foreach($consistency_error in $script:sidekick_results.consistency_errors.getEnumerator())
            {
                $value = "Using both: " + $consistency_error.key + " & " + $consistency_error.value
                if($first -eq "")
                {
                    $first = $value
                }
                $script:consistency_errors_combo.Items.Add($value); 
            }
            $script:consistency_errors_combo.SelectedItem = $first
            $script:consistency_errors_combo.AccessibleName = "On";

            ################################################################################
            ######Update Sidekick Top Acros Combo (No Rebuild)##############################
            $script:top_used_acros_combo.AccessibleName = "Off";
            $script:top_used_acros_combo.Items.Clear();
            $first = "";   
            foreach($acro in $script:sidekick_results.unique_acros.getEnumerator() | sort value -descending | Select-Object -First 20)
            {

                [string]$value = [string]$acro.value + "x - " + $acro.key
                if($first -eq "")
                {
                    $first = $value
                }
                $script:top_used_acros_combo.Items.Add($value); 
            }
            $script:top_used_acros_combo.SelectedItem = $first
            $script:top_used_acros_combo.AccessibleName = "On";


            ################################################################################
            ######Update Sidekick Top Used Words Combo (No Rebuild)#########################
            $script:top_used_words_combo.AccessibleName = "Off";
            $script:top_used_words_combo.Items.Clear();
            $first = "";
            foreach($word in $script:sidekick_results.word_counter.getEnumerator() | sort value -descending | Select-Object -First 10)
            {
                [string]$value = [string]$word.value + "x - " + $word.key
                if($first -eq "")
                {
                    $first = $value
                }
                $script:top_used_words_combo.Items.Add($value); 
            }
            $script:top_used_words_combo.SelectedItem = $first
            $script:top_used_words_combo.AccessibleName = "On";


            ################################################################################
            ######Update Sidekick Metrics Used Combo (No Rebuild)###########################
            $script:metrics_used_combo.AccessibleName = "Off";
            $script:metrics_used_combo.Items.Clear();
            $first = "";
            foreach($metric in $script:sidekick_results.metrics.getEnumerator() | sort value)
            {
                $metric_split = $metric.key -split '::'
                $value = [string]$metric_split[1] + " " + [string]$metric.value + "                                                                                                      ::" + $metric_split[0];
                if($first -eq "")
                {
                    $first = $value
                }
                $script:metrics_used_combo.Items.Add($value); 
            }
            $script:metrics_used_combo.SelectedItem = $first
            $script:metrics_used_combo.AccessibleName = "On";
              

            ################################################################################
            ################################################################################
        }#Sidekick Panel Refresh
    }#Sidekick Panel Width
    else
    {
        $left_panel.Controls.Clear();
        log "Destroyed Sidekick"
    }
}
################################################################################
######Acronym List Dialog#######################################################
function acronym_list_dialog($type)
{
    ######################################################################
    $acronym_list_form = New-Object System.Windows.Forms.Form
    $acronym_list_form.FormBorderStyle = 'Fixed3D'
    $acronym_list_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $acronym_list_form.Location = new-object System.Drawing.Point(0, 0)
    $acronym_list_form.MaximizeBox = $false
    $acronym_list_form.SizeGripStyle = "Hide"
    $acronym_list_form.Width = 500
    $acronym_list_form.height = 600

    $y_pos = 10
    $acronym_title                         = New-Object system.Windows.Forms.Label
    $acronym_title.text                     = "$type Acronym List";
    $acronym_title.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $acronym_title.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $acronym_title.Anchor                   = 'top,right'
    $acronym_title.width                    = ($acronym_list_form.width)
    $acronym_title.height                   = 30
    $acronym_title.TextAlign = "MiddleCenter"
    $acronym_title.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $acronym_title.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $acronym_list_form.Controls.Add($acronym_title)


    $y_pos = $y_pos + 40

    $acronym_box = New-Object System.Windows.Forms.RichTextBox
    $acronym_box.Size = New-Object System.Drawing.Size(($acronym_list_form.width - 30),($acronym_list_form.height - 100))
    $acronym_box.Location = New-Object System.Drawing.Size(10,$y_pos)    
    $acronym_box.ReadOnly = $false
    $acronym_box.WordWrap = $True
    $acronym_box.Multiline = $True
    $acronym_box.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
    $acronym_box.ForeColor = $script:theme_settings['EDITOR_FONT_COLOR']
    $acronym_box.ScrollBars = "Vertical"
    $acronym_box.AccessibleName = "";

    if($type -eq "EPR")
    {
        $line = "";
        foreach($acro in $script:sidekick_results.acro_list.getEnumerator() | sort value)
        {
            $line = $line + $acro.key + " (" + $acro.value + "); "
        }
        if($line.length -ge 2)
        {
            $line = $line.substring(0,($line.Length - 2))
        }
        $acronym_box.Text = $line
        $acronym_box.SelectAll();
        $acronym_box.selectionfont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        $acronym_box.DeselectAll()
    }
    else
    {
        $line = "";
        foreach($acro in $script:sidekick_results.acro_list.getEnumerator() | sort value)
        {
            $line = $line + $acro.value + " - " + $acro.key + "`n";
        }
        $acronym_box.Text = $line
        $acronym_box.SelectAll();
        $acronym_box.selectionfont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
        $acronym_box.DeselectAll()
    }
    $acronym_box.Add_TextChanged({
        if($this.AccessibleName -eq "")
        {
            $this.AccessibleName = "Notified"
            $message = "Any edits you make to your acronym list will not be saved when you close this window. Copy & Paste your work somewhere before exiting!`n"
            [System.Windows.MessageBox]::Show($message,"!!!WARNING!!!",'Ok')
        }
    })

    $acronym_box.Add_MouseDown({     
        if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Right ) 
        {
            if($acronym_box.SelectedText.Length -ge 1)
            {
                $contextMenuStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
                $acronym_box.ContextMenuStrip = $contextMenuStrip1        
                $contextMenuStrip1.Items.Add("Copy").add_Click({clipboard_copy_3})
            }
        }
    })
    $acronym_box.Add_KeyUp({
        if(($_.control) -and ($_.keycode -match "c"))
        {
            #write-host Copy Feeder
            clipboard_copy_3
        }
    })
    $acronym_list_form.Controls.Add($acronym_box)
    $acronym_list_form.ShowDialog()
}
################################################################################
######System Settings Dialog####################################################
function system_settings_dialog
{
    $system_settings_form = New-Object System.Windows.Forms.Form
    $system_settings_form.FormBorderStyle = 'Fixed3D'
    $system_settings_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $system_settings_form.Location = new-object System.Drawing.Point(0, 0)
    $system_settings_form.MaximizeBox = $false
    $system_settings_form.SizeGripStyle = "Hide"
    $system_settings_form.text                     = "System Settings";
    $system_settings_form.Width = 800
    $system_settings_form.Height = 375

    ################################################################################
    ######Title#####################################################################
    $y_pos = 10;
    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "System Settings";
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($system_settings_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point((($system_settings_form.width / 2) - ($title_label.width / 2)),$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $system_settings_form.controls.Add($title_label);



    ################################################################################
    ######Calculator Label##########################################################
    $y_pos = $y_pos + 45
    $text_size_calculator_label                          = New-Object system.Windows.Forms.Label
    $text_size_calculator_label.text                     = "Text Calculation Display:";
    $text_size_calculator_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    #$text_size_calculator_label.backcolor = "green"
    $text_size_calculator_label.Anchor                   = 'top,left'
    #$text_size_calculator_label.autosize = $true
    $text_size_calculator_label.width                    = 225
    $text_size_calculator_label.height                   = 30
    $text_size_calculator_label.TextAlign                 = "middleright"
    $text_size_calculator_label.location                 = New-Object System.Drawing.Point(15,$y_pos)
    $text_size_calculator_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $system_settings_form.controls.Add($text_size_calculator_label);
    

    ################################################################################
    ######Calculator Trackbar#######################################################
    $text_size_calculator_trackbar_label = New-Object System.Windows.Forms.Label
    $text_size_calculator_trackbar = New-Object System.Windows.Forms.TrackBar
    $text_size_calculator_trackbar.Width = 80
    $text_size_calculator_trackbar.Location = New-Object System.Drawing.Point(($text_size_calculator_label.location.x + $text_size_calculator_label.width + 5),($y_pos -2))
    $text_size_calculator_trackbar.Orientation = "Horizontal"
    $text_size_calculator_trackbar.Height = 30
    $text_size_calculator_trackbar.TickFrequency = 1
    $text_size_calculator_trackbar.TickStyle = "TopLeft"
    $text_size_calculator_trackbar.SetRange(1, 2) 
    $text_size_calculator_trackbar.LargeChange = 1;
    $text_size_calculator_trackbar.AccessibleName = "Off"
    $text_size_calculator_trackbar.value = 1
    $text_size_calculator_trackbar.add_ValueChanged({
        if($this.value -eq 1)
        {
            $text_size_calculator_trackbar_label.text = "Count Down"
            $script:settings['SIZER_BOX_INVERTED'] = 1;
        }
        else
        {
            $text_size_calculator_trackbar_label.text = "Count Up"
            $script:settings['SIZER_BOX_INVERTED'] = 2;
            
        }
        update_sizer_box
        #write-host Sizer Box: $script:settings['SIZER_BOX_INVERTED']
        update_settings;
    })
    $text_size_calculator_trackbar.Value = $script:settings['SIZER_BOX_INVERTED']


    ################################################################################
    ######Calculator Trackbar Label#################################################
    $text_size_calculator_trackbar_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $text_size_calculator_trackbar_label.width = 350
    #$text_size_calculator_trackbar_label.backcolor = "green"
    $text_size_calculator_trackbar_label.text = "Count Down"
    $text_size_calculator_trackbar_label.height = 30
    $text_size_calculator_trackbar_label.Location = New-Object System.Drawing.Point(($text_size_calculator_trackbar.location.x + $text_size_calculator_trackbar.width + 5),$y_pos)
    $text_size_calculator_trackbar_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $text_size_calculator_trackbar_label.TextAlign = "Middleleft"

    $system_settings_form.Controls.Add($text_size_calculator_trackbar_label)
    $system_settings_form.controls.Add($text_size_calculator_trackbar)


    ################################################################################
    ######Calculator Trackbar Label 2###############################################
    $y_pos = $y_pos + 50;
    $clock_speed_label                          = New-Object system.Windows.Forms.Label
    $clock_speed_label.text                     = "Application Clock Speed";
    $clock_speed_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    #$clock_speed_label.backcolor = "green"
    $clock_speed_label.Anchor                   = 'top,left'
    #$clock_speed_label.autosize = $true
    $clock_speed_label.width                    = 225
    $clock_speed_label.height                   = 30
    $clock_speed_label.TextAlign = "Middleright"
    $clock_speed_label.location                 = New-Object System.Drawing.Point(15,$y_pos)
    $clock_speed_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $system_settings_form.controls.Add($clock_speed_label);


    ################################################################################
    ######Clock Speed Trackbar######################################################
    $clock_speed_trackbar_label = New-Object System.Windows.Forms.Label
    $clock_speed_trackbar = New-Object System.Windows.Forms.TrackBar
    $clock_speed_trackbar.Width = 250
    $clock_speed_trackbar.Location = New-Object System.Drawing.Point(($clock_speed_label.location.x + $clock_speed_label.width + 5),$y_pos)
    $clock_speed_trackbar.Orientation = "Horizontal"
    $clock_speed_trackbar.Height = 30
    $clock_speed_trackbar.TickFrequency = 500
    $clock_speed_trackbar.TickStyle = "TopLeft"
    $clock_speed_trackbar.SetRange(100, 3000) 
    $clock_speed_trackbar.LargeChange = 500;
    $clock_speed_trackbar.AccessibleName = "Off"
    $clock_speed_trackbar.value = 100
    $clock_speed_trackbar.add_ValueChanged({
        $value1 = 3100 - $this.value
        $value3 = $value1 / 1000;

        $this.AccessibleName = $value1

        $value2 = $this.value 
        $verb = "None"

        if($value2 -le 600)
        {
            $verb = "$value3 Seconds Between Operations`nSlowest PCs"

        }
        elseif($value2 -le 1200)
        {
            $verb = "$value3 Seconds Between Operations`nSlow PCs"
        }
        elseif($value2 -le 1800)
        {
            $verb = "$value3 Seconds Between Operations`nModerate PCs"
        }
        elseif($value2 -le 2400)
        {
            $verb = "$value3 Seconds Between Operations`nFast PCs"
        }
        elseif($value2 -le 3000)
        {
            $verb = "$value3 Seconds Between Operations`nFastest PCs"
        }

        $message = "$verb"

        $clock_speed_trackbar_label.text = $message             
    })
    $clock_speed_trackbar.Value = (3100 - $script:settings['CLOCK_SPEED'])


    ################################################################################
    ######Clock Speed Trackbar Label2######################################################        
    $clock_speed_trackbar_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $clock_speed_trackbar_label.width = 280
    $clock_speed_trackbar_label.height = 30
    $clock_speed_trackbar_label.Location = New-Object System.Drawing.Point(($clock_speed_trackbar.location.x + $clock_speed_trackbar.width + 5),$y_pos)
    $clock_speed_trackbar_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    #$clock_speed_trackbar_label.BackColor = "Green"
    $clock_speed_trackbar_label.TextAlign = "MiddleLeft"
    $system_settings_form.Controls.Add($clock_speed_trackbar_label)
    $system_settings_form.controls.Add($clock_speed_trackbar)



    ################################################################################
    ######Text History Threshold Label##############################################
    $y_pos = $y_pos + 50
    $text_history_threshold_label                          = New-Object system.Windows.Forms.Label
    $text_history_threshold_label.text                     = "Text History Threshold:";
    $text_history_threshold_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    #$text_history_threshold_label.backcolor = "green"
    $text_history_threshold_label.Anchor                   = 'top,left'
    #$text_history_threshold_label.autosize = $true
    $text_history_threshold_label.width                    = 225
    $text_history_threshold_label.height                   = 30
    $text_history_threshold_label.TextAlign                 = "middleright"
    $text_history_threshold_label.location                 = New-Object System.Drawing.Point(15,$y_pos)
    $text_history_threshold_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $system_settings_form.controls.Add($text_history_threshold_label);


    ################################################################################
    ######Clock Speed Trackbar######################################################
    $text_history_threshold_label = New-Object System.Windows.Forms.Label
    $text_history_threshold = New-Object System.Windows.Forms.TrackBar
    $text_history_threshold.Width = 250
    $text_history_threshold.Location = New-Object System.Drawing.Point(($clock_speed_label.location.x + $clock_speed_label.width + 5),$y_pos)
    $text_history_threshold.Orientation = "Horizontal"
    $text_history_threshold.Height = 30
    $text_history_threshold.TickFrequency = 100
    $text_history_threshold.TickStyle = "TopLeft"
    $text_history_threshold.SetRange(50, 1000) 
    $text_history_threshold.LargeChange = 500;
    $text_history_threshold.AccessibleName = "Off"
    $text_history_threshold.value = 100
    $text_history_threshold.add_ValueChanged({
        $value = $this.value
 
        if($value -le 50)
        {
            $verb = "Slowest PCs"
        }
        elseif($value -le 100)
        {
            $verb = "Slow PCs"
        }
        elseif($value -le 500)
        {
            $verb = "Moderate PCs"
        }
        elseif($value -le 700)
        {
            $verb = "Fast PCs"
        }
        elseif($value -le 1000)
        {
            $verb = "Fastest PCs"
        }
        [string]$message = $this.value
        [string]$message = $message + "`n$verb"

        $text_history_threshold_label.text = $message           
    })
    $text_history_threshold.Value = $script:settings['SAVE_HISTORY_THRESHOLD']

    ################################################################################
    ######Clock Speed Trackbar Label2######################################################        
    $text_history_threshold_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $text_history_threshold_label.width = 280
    $text_history_threshold_label.height = 30
    $text_history_threshold_label.Location = New-Object System.Drawing.Point(($text_history_threshold.location.x + $text_history_threshold.width + 5),$y_pos)
    $text_history_threshold_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    #$text_history_threshold_label.BackColor = "Green"
    $text_history_threshold_label.TextAlign = "MiddleLeft"
    $system_settings_form.Controls.Add($text_history_threshold_label)
    $system_settings_form.controls.Add($text_history_threshold)


    ################################################################################
    ######Memory Flushing Label1####################################################
    $y_pos = $y_pos + 50
    $memory_flushing_label                          = New-Object system.Windows.Forms.Label
    $memory_flushing_label.text                     = "Memory Flushing (Beta):";
    $memory_flushing_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    #$memory_flushing_label.backcolor = "green"
    $memory_flushing_label.Anchor                   = 'top,left'
    #$memory_flushing_label.autosize = $true
    $memory_flushing_label.width                    = 225
    $memory_flushing_label.height                   = 30
    $memory_flushing_label.TextAlign                 = "middleright"
    $memory_flushing_label.location                 = New-Object System.Drawing.Point(15,$y_pos)
    $memory_flushing_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $system_settings_form.controls.Add($memory_flushing_label);


    ################################################################################
    ######Memory Flushing Trackbar##################################################
    $memory_flushing_trackbar_label = New-Object System.Windows.Forms.Label
    $memory_flushing_trackbar = New-Object System.Windows.Forms.TrackBar
    $memory_flushing_trackbar.Width = 150
    $memory_flushing_trackbar.Location = New-Object System.Drawing.Point(($memory_flushing_label.location.x + $memory_flushing_label.width + 5),($y_pos -2))
    $memory_flushing_trackbar.Orientation = "Horizontal"
    $memory_flushing_trackbar.Height = 30
    $memory_flushing_trackbar.TickFrequency = 1
    $memory_flushing_trackbar.TickStyle = "TopLeft"
    $memory_flushing_trackbar.SetRange(1, 4) 
    $memory_flushing_trackbar.LargeChange = 1;
    $memory_flushing_trackbar.AccessibleName = "Off"
    $memory_flushing_trackbar.value = 1
    $memory_flushing_trackbar.add_ValueChanged({
        if($this.value -eq 1)
        {
            $memory_flushing_trackbar_label.text = "Off"
        }
        elseif($this.value -eq 2)
        {
            $memory_flushing_trackbar_label.text = "Garbage Collect Only"
        }
        elseif($this.value -eq 3)
        {
            $memory_flushing_trackbar_label.text = "Memory Flushing Only (Recommended)"
        }
        else
        {
            $memory_flushing_trackbar_label.text = "Memory Flushing & Garbage Collect"
        }
    })
    $memory_flushing_trackbar.Value = $script:settings['MEMORY_FLUSHING']


    ################################################################################
    ######Memory Flushing Trackbar Label#############################################
    $memory_flushing_trackbar_label.Font   = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $memory_flushing_trackbar_label.width = 330
    $memory_flushing_trackbar_label.height = 30
    $memory_flushing_trackbar_label.Location = New-Object System.Drawing.Point(($memory_flushing_trackbar.location.x + $memory_flushing_trackbar.width + 5),$y_pos)
    $memory_flushing_trackbar_label.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    #$memory_flushing_trackbar_label.BackColor = "Green"
    $memory_flushing_trackbar_label.TextAlign = "MiddleLeft"
    $system_settings_form.Controls.Add($memory_flushing_trackbar_label)
    $system_settings_form.controls.Add($memory_flushing_trackbar)



    ################################################################################
    ######Memory Usage Label1#######################################################
    $y_pos = $y_pos + 50
    $memory_usage_label1                          = New-Object system.Windows.Forms.Label
    $memory_usage_label1.text                     = "Current Memory Usage:";
    $memory_usage_label1.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    #$memory_usage_label1.backcolor = "green"
    $memory_usage_label1.Anchor                   = 'top,left'
    #$memory_usage_label1.autosize = $true
    $memory_usage_label1.width                    = 225
    $memory_usage_label1.height                   = 30
    $memory_usage_label1.TextAlign                 = "middleright"
    $memory_usage_label1.location                 = New-Object System.Drawing.Point(15,$y_pos)
    $memory_usage_label1.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $system_settings_form.controls.Add($memory_usage_label1);


    ################################################################################
    ######Memory Usage Label 2######################################################
    $current_memory = (Get-Process -id $PID | Sort-Object WorkingSet64 | Select-Object Name,@{Name='WorkingSet';Expression={($_.WorkingSet64)}})
    $current_memory = [System.Math]::Round(($current_memory.WorkingSet)/1mb, 2)
    $memory_usage_label2            = New-Object System.Windows.Forms.Label     
    $memory_usage_label2.Font       = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $memory_usage_label2.Text       = "$current_memory MB"
    $memory_usage_label2.width      = 280
    $memory_usage_label2.height     = 30
    $memory_usage_label2.Location   = New-Object System.Drawing.Point(($memory_usage_label1.location.x + $memory_usage_label1.width + 5),$y_pos)
    $memory_usage_label2.ForeColor  = $script:theme_settings['DIALOG_FONT_COLOR']
    $memory_usage_label2.TextAlign  = "MiddleLeft"
    $system_settings_form.Controls.Add($memory_usage_label2)


    ################################################################################
    ######Submit Button#############################################################
    $y_pos = $y_pos + 40;
    $submit_button           = New-Object System.Windows.Forms.Button
    $submit_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $submit_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $submit_button.Width     = 110
    $submit_button.height     = 25
    $submit_button.Location  = New-Object System.Drawing.Point((($system_settings_form.width / 2) - ($submit_button.width)),$y_pos);
    $submit_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $submit_button.Text      ="Save"
    $submit_button.Name = ""
    $submit_button.Add_Click({
        $script:settings['CLOCK_SPEED'] = $clock_speed_trackbar.AccessibleName
        $Script:Timer.Interval = $script:settings['CLOCK_SPEED']

        $script:settings['SAVE_HISTORY_THRESHOLD'] = $text_history_threshold.value
        $script:settings['MEMORY_FLUSHING'] = $memory_flushing_trackbar.value
        update_settings;
        $system_settings_form.close();

    });
    $system_settings_form.controls.Add($submit_button)

    ################################################################################
    ######Submit Button#############################################################
    $cancel_button           = New-Object System.Windows.Forms.Button
    $cancel_button.BackColor = $script:theme_settings['DIALOG_BUTTON_BACKGROUND_COLOR']
    $cancel_button.ForeColor = $script:theme_settings['DIALOG_BUTTON_TEXT_COLOR']
    $cancel_button.Width     = 110
    $cancel_button.height     = 25
    $cancel_button.Location  = New-Object System.Drawing.Point((($system_settings_form.width / 2)),$y_pos);
    $cancel_button.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'])    
    $cancel_button.Text      ="Cancel"
    $cancel_button.Add_Click({
        $system_settings_form.close();
    });
    $system_settings_form.controls.Add($cancel_button) 

    $system_settings_form.ShowDialog()
}
################################################################################
######Logger####################################################################
Function Log($message) 
{
    if(($script:print_to_log -eq 1) -or ($script:print_to_console -eq 1))
    {
        if($message -eq "BLANK")
        {
            Add-Content -literalpath "$script:logfile" -Value ""
            write-host ""
        }
        elseif(($message.length -ge 7) -and ($message.Substring(0,7) -eq "SUBLOG "))
        {
             $message = $message.Substring(7,($message.Length - 7))
             write-host $message
        }
        else
        {

            ###Get Total Memory
            #$memory_total          = [System.Math]::Round((((Get-Process -Id $PID | workingset64 –auto).PrivateMemorySize))/1mb, 2)
            $memory_total = (Get-Process -id $PID | Sort-Object WorkingSet64 | Select-Object Name,@{Name='WorkingSet';Expression={($_.WorkingSet64)}})
            $memory_total = [System.Math]::Round(($memory_total.WorkingSet)/1mb, 2)

            ###Calculate Thread changes
            if($message -match "Start")
            {
                $script:log_mem_change = $memory_total
                [string]$mem_change_string = 0
            }
            else
            {
                $script:log_mem_change = [System.Math]::Round(($memory_total - $script:log_mem_change),2)
                [string]$mem_change_string = $script:log_mem_change
            }
            
            
            
                 
            $memory_diff           = [System.Math]::Round(($memory_total - $script:memBefore),2)
            While($memory_diff.tostring().length -lt 8)
            {
                $memory_diff = "$memory_diff" + " ";
            }
            While($memory_total.tostring().length -lt 8)
            {
                $memory_total = "$memory_total" + " ";
            }
            While($mem_change_string.length -lt 5)
            {
                $mem_change_string = $mem_change_string + " ";
            }
            $var_count = (get-variable).count
            $stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
            $line = "$Stamp - Mem Start Diff: $memory_diff - Mem Total: $memory_total - Mem Thread Change: $mem_change_string - Vars: $var_count $Message"
            $script:log_mem_change = $memory_total
            if($script:print_to_log -eq "1")
            {
                Add-Content -literalpath "$script:logfile" -Value $line
            }
            if($script:print_to_console -eq "1")
            {
                Write-host $line
            }
        }
    }
}
################################################################################
######Variable Sizer############################################################
function var_sizes
{
    if($script:var_size_detection -eq 1)
    {
        $var_sizes2 = New-Object system.collections.hashtable 
        foreach($var in $var_sizes.GetEnumerator())
        {
            $var_sizes2.add($var.key,$var.value);
        }
        foreach($var in (Get-variable))
        {
            if($var.name -and $var.value)
            {
                $name = $var.name
                $type = $var.value.GetType()
                #write-host $name = $type      
                if($type -match "Hash")
                {
                    #write-host $name
                    $size = 0;
                    foreach($table in ($var.name | Get-variable -ValueOnly))
                    {
                        foreach($key in $table.keys)
                        {
                            #write-host K-----
                            #write-host $key - $size - $key.Length
                            $size = $size + ($key.tostring()).Length
                            #write-host $key - $size - $key.Length
                    
                        }
                        foreach($value in $table.values)
                        {
                            #write-host V-----
                            #write-host $value - $size - $value.Length
                            $size = $size + ($value.tostring()).Length
                            #write-host $value - $size - $value.Length
                    
                        }
                    }
                    $var_sizes["Hash   - $name"] = $size
                    #write-host $name = $size     
                }
                elseif($type -match "^String")
                {
                    #write-host ----------------$type
                    $size = ($var.name.Length) + ($var.value.length)
                    #write-host $var.name $var.value
                    $var_sizes["String - $name"] = $size

                }
                elseif($type -match "Array")
                {
                    #write-host ------------------$type $var.name #$var.value
                    $size = 0;
                    foreach($item in $var.value)
                    {
                        if(Test-path variable:item)
                        {
                            foreach($part in $item)
                            {
                                if($part -ne "")
                                {
                                    foreach($seg in $part)
                                    {
                                        if((Test-path variable:seg) -and ($seg -ne $null) -and ($seg.psobject.properties['name']))
                                        {
                                            #write-host $seg.name
                                            $size = $size + (($seg.name).ToString()).Length
                                        }
                                    }
                                    if((Test-path variable:seg) -and ($seg -ne $null) -and ($seg.psobject.properties['value']))
                                    {
                                        foreach($seg in $part.value)
                                        {
                                        
                                                #write-host $seg
                                                $size = $size + (($seg).ToString()).Length
                                            
                                        }
                                    }
                                }
                            }
                        }
                    }
                    $var_sizes["Array  - $name"] = $size
                }
                elseif($type -match "Form|custom|window")
                {
                    #$var_sizes["Form - $name"] = "$type"
                    [string]$size1 = $var.value
                    $size1 =  $size1.Length
                
                    $var_sizes["Form   - $name"] = $size1

                }
                elseif($type -match "^int|^double|^float|^long")
                {
                    #write-host ----------------$type
                    $size = ($var.name.Length) + ($var.value.tostring()).length
                    #write-host $size = $var.name $var.value
                    $var_sizes["Int    - $name"] = $size
                }
                else
                {
                    #$var_sizes["Unk - $name"] = "$type"
                }
            }
        }
        #################################################
        write-host ---------------------VarList Start
        foreach($var in $var_sizes2.GetEnumerator())
        {
            if(!($var_sizes.Contains($var.key)))
            {
                write-host Added Var $var.key
            }
            else
            {
                if([int]$var.value -gt [int]$var_sizes[$var.key])
                {
                    write-host Shrunk $var.key $var.value to $var_sizes[$var.key]
                }
                elseif([int]$var.value -lt [int]$var_sizes[$var.key])
                {
                    write-host Grew $var.key $var.value to $var_sizes[$var.key]
                }
            }
        }
        foreach($var in $var_sizes.GetEnumerator() | sort Value)
        {
            #write-host $var.key = $var.value
        }
        write-host ---------------------VarList End
    }
}
################################################################################
######FAQ Dialog################################################################
function FAQ_dialog
{
    $FAQ_form = New-Object System.Windows.Forms.Form
    $FAQ_form.FormBorderStyle = 'Fixed3D'
    $FAQ_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $FAQ_form.Location = new-object System.Drawing.Point(0, 0)
    $FAQ_form.MaximizeBox = $false
    $FAQ_form.SizeGripStyle = "Hide"
    $FAQ_form.Width = 1200
    $FAQ_form.Height = 600

    $y_pos = 10;

    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "Frequently Asked Questions"
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($FAQ_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $FAQ_form.controls.Add($title_label);

    $y_pos = 45;

    $FAQ_box = New-Object System.Windows.Forms.RichTextBox
    $FAQ_box.Size = New-Object System.Drawing.Size(($FAQ_form.width - 30),($FAQ_form.height - 90))
    $FAQ_box.Location = New-Object System.Drawing.Size(10,$y_pos)    
    $FAQ_box.ReadOnly = $true
    $FAQ_box.WordWrap = $True
    $FAQ_box.Multiline = $True
    $FAQ_box.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
    $FAQ_box.ForeColor = $script:theme_settings['EDITOR_FONT_COLOR']
    $FAQ_box.ScrollBars = "Vertical"
    $FAQ_box.AccessibleName = "";
    $FAQ_box.text = 
"Q - Did you use any duty hours or government resources to make this?
    A - No, this entire program/script was written on my free-time after hours on my personal computer.

Q - Why is the text size different from Adobe's?
    A - Microsoft's native `"Times New Roman`" font is slightly different, and the character sizes are completely different.

Q - Why does the text size not calculate by letter?
    A - Each letter is a different number of pixels, the space provided on an EPR is not determined by the number of letters, rather by the number of pixels. 

Q - Why does it run so slow?
    A - Microsoft's PowerShell language is not designed for robust GUI environments. 
          Additionally, this was written on a very powerful PC, and most AF computers do not have the resources to properly run it.
          If your system is struggling to run it, you can change the Clock Speed under (Options -> System Settings) to a lower setting.

Q - How do I report a problem?
    A - You can e-mail me at anthony.brechtel@gmail.com, but I probably won't entertain your problem unless it's a significate issue or you've donated to the project.

Q - I'd like to request a new feature?
    A - You can e-mail me at anthony.brechtel@gmail.com, but I probably won't entertain your idea unless you've donated to the project.

Q - Why not include a list of bullets?
    A - Bullets contain a large amount of OPSEC material in them, I didn't feel comfortable publishing a list of bullets.

Q - Why can't I import/export to/from Abobe PDF files?
    A - Adobe has security features and a proprietary API. It is not very friendly to outside applications. 

Q - I'm having problems running it on other computers.
    A - Your most likely problem is PowerShell's execution policy settings or your computer does not have PowerShell V5+ installed.

Q - Why does it keep locking up?
    A - Not really sure exactly why this is happening, but it appears to happen during idle. (Still looking into this)

Q - I noticed a problem with your code, can I send a solution?
    A - Absolutely, this program was written in under 2 months, if you see a better way to approach something I'm all ears.
        I probably will not re-write entire sections unless you provide the code. 

Q - I hate this program.
    A - Not a question, but you don't have to use it.

Q - Did anyone help you?
    A - My wife did some testing, but most of the help I received was from the thousands of anonymous programmers on the web.

Q - Why did you make this?
    A - I like programming, I think it's fun, and I wanted a useful/fun challenge.
    "

    $FAQ_box.SelectAll();
    $FAQ_box.selectionfont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    $FAQ_box.DeselectAll()


    $FAQ_form.controls.Add($FAQ_box);
    $FAQ_form.ShowDialog()
}
################################################################################
######About Dialog##############################################################
function about_dialog
{
      
    $about_form = New-Object System.Windows.Forms.Form
    $about_form.FormBorderStyle = 'Fixed3D'
    $about_form.BackColor             = $script:theme_settings['DIALOG_BACKGROUND_COLOR']
    $about_form.Location = new-object System.Drawing.Point(0, 0)
    $about_form.MaximizeBox = $false
    $about_form.SizeGripStyle = "Hide"
    $about_form.Width = 800
    $about_form.Height = 600

    $y_pos = 10;

    $title_label                          = New-Object system.Windows.Forms.Label
    $title_label.text                     = "About",$script:program_title
    $title_label.ForeColor                = $script:theme_settings['DIALOG_TITLE_FONT_COLOR']
    $title_label.Backcolor                = $script:theme_settings['DIALOG_TITLE_BANNER_COLOR']
    $title_label.Anchor                   = 'top,right'
    $title_label.width                    = ($about_form.width)
    $title_label.height                   = 30
    $title_label.TextAlign = "MiddleCenter"
    $title_label.location                 = New-Object System.Drawing.Point(0,$y_pos)
    $title_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] + 4))
    $about_form.controls.Add($title_label);

    $y_pos = 45;

    $version_name_label                          = New-Object system.Windows.Forms.Label
    $version_name_label.text                     = "Version:";
    $version_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $version_name_label.Anchor                   = 'top,right'
    $version_name_label.autosize = $true
    $version_name_label.width                    = 120
    $version_name_label.height                   = 30
    $version_name_label.TextAlign = "MiddleLeft"
    $version_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
    $version_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $about_form.controls.Add($version_name_label);

    $version_name_value                          = New-Object system.Windows.Forms.Label
    $version_name_value.text                     = $script:program_version
    $version_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $version_name_value.Anchor                   = 'top,right'
    $version_name_value.autosize = $true
    $version_name_value.TextAlign = "MiddleLeft"
    $version_name_value.width                    = 150
    $version_name_value.height                   = 30
    $version_name_value.location                 = New-Object System.Drawing.Point((10 + $version_name_label.width),($y_pos + 3));
    $version_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
    $about_form.controls.Add($version_name_value);

    $y_pos = $y_pos + 25;

    $author_name_label                          = New-Object system.Windows.Forms.Label
    $author_name_label.text                     = "Written By:";
    $author_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $author_name_label.Anchor                   = 'top,right'
    $author_name_label.autosize = $true
    $author_name_label.width                    = 120
    $author_name_label.height                   = 30
    $author_name_label.TextAlign = "MiddleLeft"
    $author_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
    $author_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $about_form.controls.Add($author_name_label);

    $author_name_value                          = New-Object system.Windows.Forms.Label
    $author_name_value.text                     = "Anthony V. Brechtel"
    $author_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $author_name_value.Anchor                   = 'top,right'
    $author_name_value.autosize = $true
    $author_name_value.TextAlign = "MiddleLeft"
    $author_name_value.width                    = 150
    $author_name_value.height                   = 30
    $author_name_value.location                 = New-Object System.Drawing.Point((10 + $author_name_label.width),($y_pos + 3));
    $author_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
    $about_form.controls.Add($author_name_value);

    $y_pos = $y_pos + 25

    $donate_name_label                          = New-Object system.Windows.Forms.Label
    $donate_name_label.text                     = "Donate:";
    $donate_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $donate_name_label.Anchor                   = 'top,right'
    $donate_name_label.autosize = $true
    $donate_name_label.width                    = 120
    $donate_name_label.height                   = 30
    $donate_name_label.TextAlign = "MiddleLeft"
    $donate_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
    $donate_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $about_form.controls.Add($donate_name_label);

    $donate_name_value                          = New-Object system.Windows.Forms.Label
    $donate_name_value.text                     = "https://donorbox.org/bullet-blender (Click Here)"
    $donate_name_value.ForeColor                = $script:theme_settings['DIALOG_FONT_COLOR']
    $donate_name_value.Anchor                   = 'top,right'
    $donate_name_value.autosize = $true
    $donate_name_value.TextAlign = "MiddleLeft"
    $donate_name_value.width                    = 150
    $donate_name_value.height                   = 30
    $donate_name_value.location                 = New-Object System.Drawing.Point((10 + $donate_name_label.width),($y_pos + 3));
    $donate_name_value.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], [Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] - 1)
    $donate_name_value.add_click({
        Start-Process "https://donorbox.org/bullet-blender"
    })
    $about_form.controls.Add($donate_name_value);
    
    $y_pos = $y_pos + 25

    $version_history_name_label                          = New-Object system.Windows.Forms.Label
    $version_history_name_label.text                     = "Version History:";
    $version_history_name_label.ForeColor                = $script:theme_settings['DIALOG_SUB_HEADER_COLOR']
    $version_history_name_label.Anchor                   = 'top,right'
    $version_history_name_label.autosize = $true
    $version_history_name_label.width                    = 120
    $version_history_name_label.height                   = 30
    $version_history_name_label.TextAlign = "MiddleLeft"
    $version_history_name_label.location                 = New-Object System.Drawing.Point(5,$y_pos)
    $version_history_name_label.Font                     = [Drawing.Font]::New($script:theme_settings['INTERFACE_FONT'], ([Decimal]$script:theme_settings['INTERFACE_FONT_SIZE'] ))
    $about_form.controls.Add($version_history_name_label);

    $y_pos = $y_pos + 35

    $version_box = New-Object System.Windows.Forms.RichTextBox
    $version_box.Size = New-Object System.Drawing.Size(($about_form.width - 30),($about_form.height - 200))
    $version_box.Location = New-Object System.Drawing.Size(10,$y_pos)    
    $version_box.ReadOnly = $true
    $version_box.WordWrap = $True
    $version_box.Multiline = $True
    $version_box.BackColor = $script:theme_settings['EDITOR_BACKGROUND_COLOR']
    $version_box.ForeColor = $script:theme_settings['EDITOR_FONT_COLOR']
    $version_box.ScrollBars = "Vertical"
    $version_box.AccessibleName = "";
    $version_box.text = "
    --------------------------------------------------------------------
    Version 1.4:
    --------------------------------------------------------------------
    Date: 4 Dec 2021
    Performance: Eliminated Invisible ENV errors, causing memory leaks
    Performance: Re-arranged Script Scope variables 
    New Feature: Show total used Acronyms
    Bug Fixed: Unique Acronyms count incorrect
    Bug Fixed: Duplicate word detection more accurate
    Bug Fixed: Extra space detection now visible

    --------------------------------------------------------------------
    Version 1.3:
    --------------------------------------------------------------------
    Date: 2 Dec 2021
    New Feature: Ability to lookup words via WordHippo.
    New Feature: Overhauled System Settings
    Update: Added/Updated Default Theme `"Dark Castle`"
    Performance: History File now has limited growth in memory (Beta)  
    Performance: Ability to control memory flushing (Beta)
    Bug Fixed: Changed Text Size calculation by 1 pixel.               
    Bug Fixed: Fixed issue with changing theme colors
    Bug Fixed: Removed gap above sidekick window
    Bug Fixed: Fixed issue with left mouse clicks
    Bug Fixed: Removed Invalid Match types
    Bug Fixed: Ensured Settings file validations
    
    --------------------------------------------------------------------
    Version 1.2:
    --------------------------------------------------------------------
    Bug Fixed: Closing of Themes Dialog causes Bullet Blender & ISE to crash. 

    --------------------------------------------------------------------
    Version 1.1:
    --------------------------------------------------------------------
    Date: 13 May 2021
    Bug Fixed: Excel/Word Comm Object failing to release
    Bug Fixed: Some platforms not allowing Unicode CSV saves
    Bug Fixed: csv_to_line function not returning Array
    Bug Fixed: Refined Acronym processing thresholds
    Bug Fixed: Feeder Function no longer continuously runs
    Bug Fixed: Text Calculator Rebuilt to prevent flashing
    Bug Fixed: Text Compression Not working as intended
    Bug Fixed: Refined Repeated Usage Detection
    Known Issue: System memory expands until PowerShell crashes
    Known Issue: Saving Theme sometimes causes system crash
    Known Issue: Sidekick does not always display properly

    --------------------------------------------------------------------
    Version 1.0:
    --------------------------------------------------------------------
    Project Started: 26 January 2021
    Feature Added: Text Editor
    Feature Added: Bullet Feeder
    Feature Added: Text Size Calculation
    Feature Added: Bullet Lists
    Feature Added: Acronym Scanning
    Feature Added: Acronym Lists
    Feature Added: Package Management
    Feature Added: Themes
    Feature Added: Error Detection
    Feature Added: Metrics
    Feature Added: Text Compression
    Project Released: 27 April 2021

    --------------------------------------------------------------------  
 

    "


    $version_box.SelectAll();
    $version_box.selectionfont = [Drawing.Font]::New($script:theme_settings['EDITOR_FONT'], [Decimal]$script:theme_settings['EDITOR_FONT_SIZE'])
    $version_box.DeselectAll()


    $about_form.controls.Add($version_box);
    $about_form.ShowDialog()
}
################################################################################
######Main Sequence Start#######################################################

##########Setup Pre-Vars
var_sizes
##Interface Vars
$script:Form                                           = New-Object system.Windows.Forms.Form
$script:editor                                         = New-Object CustomRichTextBox
$script:ghost_editor                                   = New-Object CustomRichTextBox
$script:MenuBar                                        = New-Object System.Windows.Forms.MenuStrip
$script:bullet_feeder_panel                            = New-Object system.Windows.Forms.Panel
$script:feeder_box                                     = New-Object System.Windows.Forms.RichTextBox
$script:sizer_box                                      = New-Object System.Windows.Forms.RichTextBox
$script:sizer_art                                      = new-object system.windows.forms.label
$script:sidekick_panel                                 = New-Object system.Windows.Forms.Panel
$script:left_panel                                     = New-Object system.Windows.Forms.Panel
$script:main_vars = Get-Variable | Select-Object -ExpandProperty Name   #Contains List of Startup Variables

#########Run Sequence
Log "Initial Checks Start"
initial_checks;
Log "Initial Checks End"
Log "BLANK"
Log "Loading Settings Start"
load_settings;
Log "Loading Settings End"
Log "BLANK"
Log "Loading Theme Start"
load_theme
Log "Loading Theme End"
Log "BLANK"
Log "Loading Character Blocks Start"
load_character_blocks;
Log "Loading Character Blocks End"
Log "BLANK"
Log "Loading Dictionary Start"
load_dictionary;
Log "Loading Dictionary End"
Log "BLANK"
Log "Main Start"
main;