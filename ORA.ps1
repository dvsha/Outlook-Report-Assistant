Add-Type -AssemblyName PresentationFramework

 $Global:inputfile=$null
 $Global:ExportFile = $null
 $Global:Staring_Row=1
 $Global:Max_Row_To_Count = 20
 $Global:Max_Column_To_Count = 20
 $Global:Max_Column_Tolerance = 10
 $Global:Maximum_Rows_Used = $null
 $Global:used_columns
 $Global:My_sheet = $null
 $Global:objExcel = $null
 $Global:workbook = $null
 $Global:heading_row_number

 $Global:Log_Location = $null
 $Global:App_Reg_Path = "HKCU:\Software\OutlookAutomation"
 $Global:Default_Log_Path = "$env:APPDATA"
 $Global:Default_Template_Path = "$env:APPDATA"
 $Global:Template_file="$env:APPDATA\Templates.xlsx"
 $Global:Installdir="${env:ALLUSERSPROFILE}\ORA"

 $Global:PSTinputfile=$null
 $Global:new_mail_request = $true
 $Global:new_task_request = $False
 $Global:Textbox_To_Selected = $False
 $Global:Textbox_CC_Selected = $False
 $Global:Textbox_BCC_Selected = $False
 $Global:Textbox_Subject_Selected = $False
 $Global:Textbox_Body_Selected = $False
 $Global:Mails_Read = [System.Collections.ArrayList]@()
 $Global:All_Checked_Nodes = [System.Collections.ArrayList]@()
 $Global:NS = $null


function GenerateForm {

      [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
      [reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null
      [System.Windows.Forms.Application]::EnableVisualStyles()
      [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

      $displays= Get-WmiObject -class "Win32_DisplayConfiguration"
      foreach ($display in $displays) {
      $w=$display.PelsWidth
      $h=$display.PelsHeight
      $w=(4*$w/5)
      $h=(4*$h/5)
      }


      $help_Message =  "Outlook Report Assistant
      Version 1.0
      Powered by PowerShell
      Developed by ATOS Global IT Solutions
      Contact dhirendra-vikash.sharma@atos.net for any assistance"



      function Get-ScriptDirectory
      {
        $Invocation = (Get-Variable MyInvocation -Scope 1).Value;
        if($Invocation.PSScriptRoot)
        {
            $Invocation.PSScriptRoot;
        }
        Elseif($Invocation.MyCommand.Path)
        {
            Split-Path $Invocation.MyCommand.Path
        }
        else
        {
            $Invocation.InvocationName.Substring(0,$Invocation.InvocationName.LastIndexOf("\"));
        }
      }

      $directorypath = Get-ScriptDirectory
      $Processing_file = "$Global:Installdir\File Repository\Processing.gif"
      $background_image_file="$Global:Installdir\File Repository\background.jpg"
      $elements_back_file="$Global:Installdir\File Repository\elements_backcolor.jpg"
      $icon_file="$Global:Installdir\File Repository\icon.ico"

      $Image = [system.drawing.image]::FromFile($Processing_file)
      $background_image=[system.drawing.image]::FromFile($background_image_file)
      $element_backImage = [system.drawing.image]::FromFile($elements_back_file)
      $Icon = New-Object system.drawing.icon($icon_file)

 Function Get_Reg_Value ($reg_path, $name)
 {
    if (Test-Path $reg_Path)
    {
        $test_value=Get-ItemProperty -Path $reg_Path | Select-Object -ExpandProperty $name -ErrorAction SilentlyContinue

        if($test_value -ne $null)
        {
	        return $test_value
        }
        else
        {
            return $false
        }
    }
    else
    {
        return $false
    }
 }

 Function Set_Reg_StringValue($reg_path, $name, $value)
 {
     if (Test-Path $reg_Path)
    {
        $test_value=Get-ItemProperty -Path $reg_Path | Select-Object -ExpandProperty $name

        if($test_value -ne $null)
        {
	        Set-ItemProperty -Path $reg_Path -Name $name -Value $value
        }
        else
        {
            New-ItemProperty -Path $reg_Path -Name $name -PropertyType String -Value $value
        }
    }
    else
    {
        New-Item $reg_Path
        New-ItemProperty -Path $reg_Path -Name $name -PropertyType String -Value $value
    }
 }


 Function Load_Template($template_name, $template_path)
 {
    $Template_objExcel = New-Object -ComObject Excel.Application
    $Template_workbook = $Template_objExcel.Workbooks.Open($template_path)
    $Template_sheet = $Template_workbook.Worksheets.Item(1)
    $Template_objExcel.Visible=$false
    $Template_Maximum_Rows_Used = ($Template_sheet.UsedRange.Rows).count
    for($i=1; $i -le $Template_Maximum_Rows_Used; $i++)
    {
        if( $Template_sheet.Cells.Item($i, 1).text -eq $template_name)
        {
            $Textbox_To.Text = $Template_sheet.Cells.Item($i, 2).text
            $Textbox_CC.Text = $Template_sheet.Cells.Item($i, 3).text
            $Textbox_BCC.Text = $Template_sheet.Cells.Item($i, 4).text
            $Textbox_Subject.Text = $Template_sheet.Cells.Item($i, 5).text
            $Textbox_Body.Text = $Template_sheet.Cells.Item($i, 6).text
        }
    }
 }


 #------------Check and Set Logging Path----------------
 $logging_path=Get_Reg_Value $Global:App_Reg_Path "LoggingPath"
 if($logging_path -ne $False)
 {
    $Global:Log_Location = $logging_path
 }
 else
 {
    Set_Reg_StringValue $Global:App_Reg_Path "LoggingPath" $Global:Default_Log_Path
    $Global:Log_Location = $Global:Default_Log_Path
 }


   #------------Check and Set Staring Row----------------
 $start_row=Get_Reg_Value $Global:App_Reg_Path "Staring_Row"
 if($start_row -ne $False)
 {
    $Global:Staring_Row = [int]$start_row
 }
 else
 {
    Set_Reg_StringValue $Global:App_Reg_Path "Staring_Row" "1"
    $Global:Staring_Row = 1
 }

  #------------Check and Set Max Row to Count----------------
 $max_row_temp=Get_Reg_Value $Global:App_Reg_Path "Max_Row_To_Count"
 if($max_row_temp -ne $False)
 {
    $Global:Max_Row_To_Count = [int]$max_row_temp
 }
 else
 {
    Set_Reg_StringValue $Global:App_Reg_Path "Max_Row_To_Count" "20"
    $Global:Max_Row_To_Count = 20
 }

   #------------Check and Set Max Column to Count----------------
 $max_col_temp=Get_Reg_Value $Global:App_Reg_Path "Max_Column_To_Count"
 if($max_col_temp -ne $False)
 {
    $Global:Max_Column_To_Count = [int]$max_col_temp
 }
 else
 {
    Set_Reg_StringValue $Global:App_Reg_Path "Max_Column_To_Count" "20"
    $Global:Max_Column_To_Count = 20
 }

  #------------Check and Set Max Column Tolerance----------------
 $max_col_tol=Get_Reg_Value $Global:App_Reg_Path "Max_Column_Tolerance"
 if($max_col_tol -ne $False)
 {
    $Global:Max_Column_Tolerance = [int]$max_col_tol
 }
 else
 {
    Set_Reg_StringValue $Global:App_Reg_Path "Max_Column_Tolerance" "10"
    $Global:Max_Column_Tolerance = 10
 }
 #------------------------------------------------------------

 Function Get_Mail_ID ($mixed_id) {
    $indexofsymbol=$mixed_id.IndexOf("@")
    $string_lenth=$mixed_id.Length
    $arr=$mixed_id.ToCharArray()
    for($i=$indexofsymbol; $i -le $string_lenth; $i++)
    {
        $char=$arr[$i]
        if($char -eq ' ' -or $char -eq '(' -or $char -eq ';' -or $char -eq '#' -or $char -eq ')'  ) {$end_position=$i ; break}
    }

    for($j=$indexofsymbol; $j -ge 0; $j--)
    {
        $char2=$arr[$j]
        if($char2 -eq ' ' -or $char2 -eq '(' -or $char2 -eq ';' -or $char2 -eq '#' -or $char2 -eq ')'  ) {$start_position=$j ; break}
    }

    if($start_position -eq $null) {$start_position=-1}
    if($end_position -eq $null) {$end_position=$string_lenth}

    $mail_id=$mixed_id.Substring($start_position+1, ($end_position - $start_position -1))
    return $mail_id
} 

 Function Number_of_mail_ids($mail_string){
    $count_of_ids=0
    $temp_mail_string=$mail_string
    do
    {
        $index_temp=$temp_mail_string.IndexOf("@")
        if($index_temp -ne -1){
            $count_of_ids=$count_of_ids+1
            $temp_mail_string=$temp_mail_string.Remove($index_temp,1)
        }
    }while ($index_temp -ne -1)
    return $count_of_ids
}

 Function Generate_Recepients_List($mixed_list){
   $mail_ids_temp=$mixed_list
   $num=Number_of_mail_ids $mixed_list
   for($i=1 ; $i -le $num; $i++)
    {

        $mail= Get_Mail_ID $mail_ids_temp
        $mail_ids_temp=$mail_ids_temp.Replace($mail, "x")
        $recepients_list=$recepients_list+$mail+";"

    }
    return $recepients_list
}
 
 Function Check_Mail_Errors1($text, $type){

        $temp=$text
        $start_index=0
        $len=$temp.length
        $temp=$temp.Replace("#Col", "#col")
        $temp=$temp.Replace("#cOl", "#col")
        $temp=$temp.Replace("#coL", "#col")
        $temp=$temp.Replace("#COl", "#col")
        $temp=$temp.Replace("#cOL", "#col")
        $temp=$temp.Replace("#CoL", "#col")
        $temp=$temp.Replace("#COL", "#col")
        do
        {
            $num_col=$temp.IndexOf("#col", $start_index)
            if($num_col -ne -1)
            {
                if($len-$num_col -eq 4)
                {
                    if($type -eq "to"){$Textbox_To.SelectionStart = $num_col; $Textbox_To.SelectionLength = 4;$Textbox_To.SelectionColor = 'red'}
                    if($type -eq "cc"){$Textbox_CC.SelectionStart = $num_col; $Textbox_CC.SelectionLength = 4;$Textbox_CC.SelectionColor = 'red'}
                    if($type -eq "bcc"){$Textbox_BCC.SelectionStart = $num_col; $Textbox_BCC.SelectionLength = 4;$Textbox_BCC.SelectionColor = 'red'}
                    if($type -eq "sub"){$Textbox_Subject.SelectionStart = $num_col; $Textbox_Subject.SelectionLength = 4;$Textbox_Subject.SelectionColor = 'red'}
                    if($type -eq "body"){$Textbox_Body.SelectionStart = $num_col; $Textbox_Body.SelectionLength = 4;$Textbox_Body.SelectionColor = 'red'}
                    return
                }
                
                try
                {
                    $num_temp=$temp.Substring($num_col+4,1)
                    if([int]$num_temp -is [int]){}
                }
                catch
                {
                    if($type -eq "to"){$Textbox_To.SelectionStart = $num_col; $Textbox_To.SelectionLength = 4;$Textbox_To.SelectionColor = 'red'}
                    if($type -eq "cc"){$Textbox_CC.SelectionStart = $num_col; $Textbox_CC.SelectionLength = 4;$Textbox_CC.SelectionColor = 'red'}
                    if($type -eq "bcc"){$Textbox_BCC.SelectionStart = $num_col; $Textbox_BCC.SelectionLength = 4;$Textbox_BCC.SelectionColor = 'red'}
                    if($type -eq "sub"){$Textbox_Subject.SelectionStart = $num_col; $Textbox_Subject.SelectionLength = 4;$Textbox_Subject.SelectionColor = 'red'}
                    if($type -eq "body"){$Textbox_Body.SelectionStart = $num_col; $Textbox_Body.SelectionLength = 4;$Textbox_Body.SelectionColor = 'red'}
                }

                $start_index=$num_col+1
            }

       }while($num_col -ne -1)

}


Function Check_Mail_Errors2($text, $type){

        $temp=$text
        $start_index=0
        $len=$temp.length
        $temp=$temp.Replace("Col", "col")
        $temp=$temp.Replace("cOl", "col")
        $temp=$temp.Replace("coL", "col")
        $temp=$temp.Replace("COl", "col")
        $temp=$temp.Replace("cOL", "col")
        $temp=$temp.Replace("CoL", "col")
        $temp=$temp.Replace("COL", "col")
      do
      {
        $new_num_col=$temp.IndexOf("col", $start_index)
        if($new_num_col -ne -1)
        {
            if($len-$new_num_col -gt 3)
            {
                $new_num_temp=$temp.Substring($new_num_col+3,1)
                try
                {
                    if([int]$new_num_temp -is [int]){}
                }
                catch{$start_index=$new_num_col +1; continue}

                if($new_num_col -ne 0){$try_hash =$temp.Substring($new_num_col-1,1)}
                if($try_hash -ne "#") 
                {
                    if($type -eq "to"){$Textbox_To.SelectionStart = $new_num_col; $Textbox_To.SelectionLength = 4;$Textbox_To.SelectionColor = 'red'}
                    if($type -eq "cc"){$Textbox_CC.SelectionStart = $new_num_col; $Textbox_CC.SelectionLength = 4;$Textbox_CC.SelectionColor = 'red'}
                    if($type -eq "bcc"){$Textbox_BCC.SelectionStart = $new_num_col; $Textbox_BCC.SelectionLength = 4;$Textbox_BCC.SelectionColor = 'red'}
                    if($type -eq "sub"){$Textbox_Subject.SelectionStart = $new_num_col; $Textbox_Subject.SelectionLength = 4;$Textbox_Subject.SelectionColor = 'red'}
                    if($type -eq "body"){$Textbox_Body.SelectionStart = $new_num_col; $Textbox_Body.SelectionLength = 4;$Textbox_Body.SelectionColor = 'red'}
                }

            }

            $start_index=$new_num_col +1
        }

       }while($new_num_col -ne -1)

}


Function Replace_Cols_with_Columns($text, $row_num){

        $temp=$text
        $start_index=0
        $len=$temp.length
        $temp=$temp.Replace("#Col", "#col")
        $temp=$temp.Replace("#cOl", "#col")
        $temp=$temp.Replace("#coL", "#col")
        $temp=$temp.Replace("#COl", "#col")
        $temp=$temp.Replace("#cOL", "#col")
        $temp=$temp.Replace("#CoL", "#col")
        $temp=$temp.Replace("#COL", "#col")
        do
        {
            $num_col=$temp.IndexOf("#col", $start_index)
            if($num_col -ne -1)
            {
                if($len-$num_col -eq 4) 
                {
                    return $temp
                }
                
                try
                {
                    $num_temp=$null
                    $num_counter=0
                    do{
                        $char_next=$temp.Substring(  ($num_col+4+$num_counter),1)
                        if([int]$char_next -is [int]){$num_temp = $num_temp+$char_next}
                        $num_counter=$num_counter+1
                        #Write-Host "Num Temp1 is :$num_temp"
                    }while([int]$char_next -is [int])
                }
                catch
                {
                    $start_index=($temp.IndexOf("#col"+$num_temp))+1
                    #Write-Host "Num Temp2 is :$num_temp"
                    #continue
                }

                if(([int]$num_temp -is [int]) -eq $true) 
                {
                    #Write-Host "Num Temp3 is :$num_temp"
                    $Col_temp="#col"+$num_temp
                    #Write-Host "Col temp is :$Col_temp"
                    $temp=$temp.Replace($Col_temp, $Global:My_sheet.Cells.Item([int]$row_num,[int]$num_temp).text)
                }

            }
        }while($num_col -ne -1)


        return $temp

}
    
   

      Function Get-FileName($initialDirectory)
      {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.Title="Select the source Excel file"
        $OpenFileDialog.filter = "Excel File| *.xls; *.xlsx; *.xlsm"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
      }

      Function Get-Multi_FileName($initialDirectory)
      {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenMultiFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenMultiFileDialog.Multiselect = $true
        $OpenMultiFileDialog.initialDirectory = $initialDirectory
        $OpenMultiFileDialog.Title="Select attachments files"
        $OpenMultiFileDialog.filter = "All files (*.*)| *.*"
        $OpenMultiFileDialog.ShowDialog() | Out-Null
        $OpenMultiFileDialog.FileNames
      }

      Function Get-SaveFolder($initialDirectory)
      { 
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.ShowNewFolderButton = $true
        $FolderBrowser.RootFolder = $initialDirectory
        [void]$FolderBrowser.ShowDialog()
        return $FolderBrowser.SelectedPath
      }

      Function Get-SaveFile($initialDirectory)
      { 
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.initialDirectory = $initialDirectory
        $SaveFileDialog.filter = "Log File| *.xlsx; *.xls"
        $SaveFileDialog.ShowDialog() | Out-Null
        $SaveFileDialog.filename
      }

      Function Get-PSTFileName($initialDirectory)
      {
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.Title="Select the source Excel file"
        $OpenFileDialog.filter = "PST File| *.pst"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
      }

      Function Fill_Tree($tree, $root_folder)
      {
        if($root_folder.Folders.Count -gt 0) 
        {
            foreach($sub_folder in $root_folder.Folders)
            {
                $subnode = New-Object System.Windows.Forms.TreeNode
                $tree.Nodes.Add($subnode)
                Fill_Tree $subnode $sub_folder
            }
        }

        $tree.Text=$root_folder.Name
      }

   
        Function Read_Selected_mail_Folders($root_folder, $all_checked_mail_nodes, $pst_RootFolder)
        {
            if($root_folder.Folders.Count -gt 0)
            {
                foreach($sub_folder in $root_folder.Folders)
                {
                    Read_Selected_mail_Folders $sub_folder $all_checked_mail_nodes $pst_RootFolder
                }
            }
            foreach($checked_mail_nodes in $all_checked_mail_nodes)
            {
                $full_path=Join-Path $pst_RootFolder.FullFolderPath $checked_mail_nodes.FullPath
                if( $root_folder.FullFolderPath -eq $full_path)
                {
                    foreach ($mail in $root_folder.Items)
                    {
                        #$msubject =$mail.subject
                        #Write-Host $msubject
                        $received_time = $mail.ReceivedTime
                        if ($received_time -gt $datePicker_Search_StartDate.Value -and $received_time -lt $datePicker_Search_DueDate.Value)
                        {
                            $Global:Mails_Read.Add($mail)
                        }
                        
                    }
                 }
             }
         }



        Function Get-Checked_Nodes ($tree)
        {
            if($tree.Nodes.Count -gt 0)
            {
                foreach($node in $tree.Nodes) {Get-Checked_Nodes $node}
            }


            foreach($node in $tree.Nodes)
            {
                if($node.Checked) {$Global:All_Checked_Nodes.Add($node)}
            }


        }

        Function Populate_Heading_Tree ($sheet, $Heading_treeView)
        {
            $row_counter=($Global:Staring_Row-1)
            $temp_column_counter=1
            $Global:Maximum_Rows_Used = ($sheet.UsedRange.Rows).count

            do{
                if ($row_counter -eq $Global:Max_Row_To_Count) {$row_counter = 0; $temp_column_counter=$temp_column_counter+1}
                if($temp_column_counter -eq $Global:Max_Column_To_Count) {[System.Windows.MessageBox]::Show("Unable to read Excel file for data");return}
                $row_counter = $row_counter +1
                $temp_row_item=$sheet.Cells.Item($row_counter, $temp_column_counter).text
                #Write-Host $row_counter  "   "  $temp_column_counter  "  "  $temp_row_item
            }while( ($temp_row_item -eq "") -or ($temp_row_item -eq $null) )

            $Global:heading_row_number= [int]$row_counter
            #Write-host "Headings start at row number :" $Global:heading_row_number

            $column_counter=$temp_column_counter
            #Write-host "Columns are sterting from :" $column_counter

            $temp_column_item=$sheet.Cells.Item($row_counter, $column_counter).text
            while(($temp_column_item -ne "") )
            {
                $column_counter=$column_counter+1
                #Write-Host "Checking cell number :" $row_counter " " $column_counter
                $temp_column_item=$sheet.Cells.Item($row_counter, $column_counter).text

                if ($temp_column_item -eq "")
                {
                    #Write-Host "Found empty cell in column no :" $column_counter
                    for($i = $column_counter; $i -le $column_counter+$Global:Max_Column_Tolerance; $i++)
                    {
                        if( $sheet.Cells.Item($row_counter, $i).text -ne "") {Write-Host "Resuming column at "$i; $column_counter=$i; $temp_column_item=$true; break}
                    }

                }

            }

            $Global:used_columns=$column_counter-1

            #Write-Host "Total columns used is :" $Global:used_columns

            for($j=1; $j -le $used_columns; $j++)
            {
                $subnode = New-Object System.Windows.Forms.TreeNode
                $Heading_treeView.Nodes.Add($subnode)
                $subnode.Text = "#Col" + $j + " - " + $sheet.Cells.Item($heading_row_number, $j).text
                $DropDown_Headings.Items.Add($sheet.Cells.Item($heading_row_number, $j).text)
            }
            
        }

        Function Write_To_Log ($Selection, $text, $style, $text_color, $font_type)
        {
            $Selection.TypeParagraph()
            $Selection.Style = $style
            $Selection.Font.Color = $text_color
            if($font_type -eq "bold")  { $Selection.Font.Bold = 1}
            else { $Selection.Font.Bold = 0}
            $Selection.TypeText("$(Get-Date) - $text")
        }


    Function Send_Mail ($mail_To, $row_temp, $Selection)
    {
            $to_recepients_list = $mail_To

            $CC_recepients_list=$Textbox_CC.Text
            if($CC_recepients_list -ne ""){
                $CC_recepients_list=Replace_Cols_with_Columns $CC_recepients_list $row_temp
                $cc_people_number=Number_of_mail_ids $CC_recepients_list
                if($cc_people_number -le 0)
                {
                    Write_To_Log $Selection "The CC field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }
            
            $BCC_recepients_list=$Textbox_BCC.Text
            if($BCC_recepients_list -ne ""){
                $BCC_recepients_list=Replace_Cols_with_Columns $BCC_recepients_list $row_temp
                $bcc_people_number=Number_of_mail_ids $BCC_recepients_list
                if($bcc_people_number -le 0)
                {
                    Write_To_Log $Selection "The BCC field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }

            $mail_subject=$Textbox_Subject.Text
            if($mail_subject -ne ""){
                $mail_subject=Replace_Cols_with_Columns $mail_subject $row_temp
                if($mail_subject -eq "") {Write_To_Log $Selection "The subject is blank or cannot be resolved" "Normal" "wdColorRed" "no_bold"}
            }

            $mail_body=$Textbox_Body.Text
            if($mail_body -ne ""){
                $mail_body=Replace_Cols_with_Columns $mail_body $row_temp
                if($mail_body -eq ""){Write_To_Log $Selection "The body is blank or cannot be resolved" "Normal" "wdColorRed" "no_bold"}
            }


            $outlook= New-Object -ComObject outlook.application
            $msg= $outlook.CreateItem(0)
            $msg.Body=$mail_body
            $msg.Subject=$mail_subject
            #----Managing Recepients-------------
            $temp_People_To = Generate_Recepients_List $to_recepients_list
            $msg.To = $temp_People_To
            if ($CC_recepients_list -ne "") {$msg.CC= Generate_Recepients_List $CC_recepients_list}
            if ($BCC_recepients_list -ne ""){$msg.BCC= Generate_Recepients_List $BCC_recepients_list}
            #-----Managing Attachments----------
            if ($attachment_box.Items.Count -gt 0)
            {
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                foreach($attachment in $attachment_box.Items) 
                {
                    $msg.Attachments.Add($attachment.ToString())
                }
            }
            #----------Checking and Sending---------
            if($check_first -eq 1)
            {
                $check_first =0
                $msg.Display()
                $input=[System.Windows.MessageBox]::Show('Is the item correct?','Check Item','YesNo','Info')
                if($input -eq "No"){return}
            }
            else
            {
                try
                {
                    $msg.send()
                    Write_To_Log $Selection "Sending mail to $temp_People_To" "Normal" "wdColorBlack" "no_bold"
                }
                catch{Write_To_Log $Selection "There was an issue sending mail for this row number" "Normal" "wdColorRed" "no_bold"}
            }
    }

    $form_main_add_Closing =  {
        if($Global:objExcel -ne $null)
        {
            [void]$Global:objExcel.Quit()
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$objExcel)
        }
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable -Name All_Checked_Nodes -ErrorAction SilentlyContinue
        Remove-Variable -Name App_Reg_Path -ErrorAction SilentlyContinue
        Remove-Variable -Name Default_Log_Path -ErrorAction SilentlyContinue
        Remove-Variable -Name heading_row_number -ErrorAction SilentlyContinue
        Remove-Variable -Name inputfile -ErrorAction SilentlyContinue
        Remove-Variable -Name Log_Location -ErrorAction SilentlyContinue
        Remove-Variable -Name Mails_Read -ErrorAction SilentlyContinue
        Remove-Variable -Name Max_Column_To_Count -ErrorAction SilentlyContinue
        Remove-Variable -Name Max_Row_To_Count -ErrorAction SilentlyContinue
        Remove-Variable -Name Max_Column_Tolerance -ErrorAction SilentlyContinue
        Remove-Variable -Name Maximum_Rows_Used -ErrorAction SilentlyContinue
        Remove-Variable -Name My_sheet -ErrorAction SilentlyContinue
        Remove-Variable -Name new_mail_request -ErrorAction SilentlyContinue
        Remove-Variable -Name new_task_request -ErrorAction SilentlyContinue
        Remove-Variable -Name NS -ErrorAction SilentlyContinue
        Remove-Variable -Name objExcel -ErrorAction SilentlyContinue
        Remove-Variable -Name PSTinputfile -ErrorAction SilentlyContinue
        Remove-Variable -Name Textbox_BCC_Selected -ErrorAction SilentlyContinue
        Remove-Variable -Name Textbox_Body_Selected -ErrorAction SilentlyContinue
        Remove-Variable -Name Textbox_CC_Selected -ErrorAction SilentlyContinue
        Remove-Variable -Name Textbox_Subject_Selected -ErrorAction SilentlyContinue
        Remove-Variable -Name Textbox_To_Selected -ErrorAction SilentlyContinue
        Remove-Variable -Name used_columns -ErrorAction SilentlyContinue
        Remove-Variable -Name workbook -ErrorAction SilentlyContinue
    }
     
   
     
    $Go_button_OnClick = {

     if($label_heading.Text -eq "Mail Replies")
     {
        
            $Go_button.Enabled = $false
            $label_User_Info.Text = "Reading mails from selected folders......"
            $progressBar1.Maximum = ([int]$Global:Maximum_Rows_Used-[int]$Global:heading_row_number)
            $progressBar1.Value=0

            #------Opening Word Document--------------
            $Word = New-Object -ComObject Word.Application
            $Document = $Word.Documents.Add()
            $Selection = $Word.Selection

            $Global:Mails_Read.Clear()
            [string]$pstPath = $Global:PSTinputfile
            #if outlook is not running, launch a hidden instance.
            $oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
            if ( $oProc -eq $null ) { Start-Process outlook -WindowStyle Hidden; Start-Sleep -Seconds 5 }
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $namespace.AddStoreEx($pstPath, "olStoreDefault")
            $pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $pstPath } )
            $pstRootFolder = $pstStore.GetRootFolder()
            Write-host $pstRootFolder.FullFolderPath

            $Global:All_Checked_Nodes.Clear()
            Get-Checked_Nodes $Mail_treeView
            Write-Host "Total nodes selected " $Global:All_Checked_Nodes.Count

            foreach($temp_node in $Global:All_Checked_Nodes) {Write-Host $temp_node.FullPath}

            $excel_sender_column = ($DropDown_Headings.SelectedIndex+1)
            Read_Selected_mail_Folders $pstRootFolder $Global:All_Checked_Nodes $pstRootFolder
            Write-Host $Global:Mails_Read.Count
            $label_User_Info.Text = "Starting to process rows......"

            for($i=$Global:heading_row_number+1; $i -le $Global:Maximum_Rows_Used; $i++)
            {
                $label_User_Info.Text = "Proccessing row number $i"
                $found=0
                $excel_temp_sender=$Global:My_sheet.Cells.Item([int]$i,[int]$excel_sender_column).text
                $excel_sender= Get_Mail_ID $excel_temp_sender
                $excel_subject = Replace_Cols_with_Columns $Textbox_Subject_New.Text $i

                Foreach($mail_object in $Global:Mails_Read)
                {
                    if($mail_object -ne $null)
                    {
                        if($mail_object.SenderEmailType -eq "EX") {$user=$mail_object.Sender.GetExchangeUser();$mail_sender = $user.PrimarySmtpAddress}
                        if($mail_object.SenderEmailType -eq "SMTP") {$mail_sender = $mail_object.SenderEmailAddress}
                        
                        $mail_subject = $mail_object.Subject
                        

                        #Write-Host $mail_sender.ToLower()
                        #Write-Host $excel_sender.ToLower()
                        #Write-Host $mail_subject.ToLower()
                        #Write-Host $excel_subject.ToLower()
                        #Write-Host "`n"

                        $mail_sender= $mail_sender.ToLower()
                        $excel_sender= $excel_sender.ToLower()
                        $mail_subject= $mail_subject.ToLower()
                        $excel_subject= $excel_subject.ToLower()

                        if(  ($mail_sender).contains($excel_sender) -and ($mail_subject).contains($excel_subject)  )
                        {
                            if(  ($mail_subject).StartsWith("automatic reply:")  ) {Write_To_Log $Selection "Automatic reply from $mail_sender" "Normal" "wdColorBlack" "no_bold"}
                            if(  ($mail_subject).StartsWith("re:")  ) 
                            {
                                Write_To_Log $Selection "Received reply from $mail_sender" "Normal" "wdColorGreen" "no_bold"
                                Write-Host "Received reply from $mail_sender"
                                $Global:My_sheet.Cells.Item([int]$i,[int]$excel_sender_column).Interior.ColorIndex = 4
                                $found=1
                                break
                            }
                        }
                    }
                }
                
                if($found -ne 1)
                {
                    Write-Host "Sending mail to $excel_sender"
                    if($CheckBox_Check_Reply.Checked -eq $true ){Send_Mail $excel_sender $i $Selection}
                }

                $progressBar1.PerformStep()
            }

            $Go_button.Enabled = $true
            $label_User_Info.Text = "Saving the log file..........."
            $date_time=((Get-Date).ToString()).Replace(":", "-")
            $Report = $Global:Log_Location+"\Checked_Mails_$date_time"
            $Document.SaveAs($Report)
            $Document.close()
            $word.Quit()
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            Remove-Variable word
            $label_User_Info.Text = "Log file created successfully as $Report"
            $Global:workbook.Save()
            $Global:workbook.Close
            $Global:objExcel.Save()
            
    }



     if($label_heading.Text -eq "Export to Excel")
     {
            $new_mail_counter=2
            $Go_button.Enabled = $false
            $label_User_Info.Text = "Exporting mails from selected folders......"

            $Global:Mails_Read.Clear()
            [string]$pstPath = $Global:PSTinputfile
            #if outlook is not running, launch a hidden instance.
            $oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
            if ( $oProc -eq $null ) { Start-Process outlook -WindowStyle Hidden; Start-Sleep -Seconds 5 }
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $namespace.AddStoreEx($pstPath, "olStoreDefault")
            $pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $pstPath } )
            $pstRootFolder = $pstStore.GetRootFolder()
            Write-host $pstRootFolder.FullFolderPath

            $Global:All_Checked_Nodes.Clear()
            Get-Checked_Nodes $Mail_treeView
            Write-Host "Total nodes selected " $Global:All_Checked_Nodes.Count

            foreach($temp_node in $Global:All_Checked_Nodes) {Write-Host $temp_node.FullPath}

            Read_Selected_mail_Folders $pstRootFolder $Global:All_Checked_Nodes $pstRootFolder
            Write-Host $Global:Mails_Read.Count
            $progressBar1.Maximum = $Global:Mails_Read.Count
            $progressBar1.Value=0

            $Report_Excel=New-Object -ComObject excel.application
            $Report_Excel.Visible = $false
            $Report_workbook=$Report_Excel.Workbooks.Add()
            $report_sheet=$Report_workbook.Worksheets.Item(1)
            $report_sheet.name = "Mail Content"

            $report_sheet.Cells.Item(1,1) = "Received Time"
            $report_sheet.Cells.Item(1,1).Font.Size=14
            $report_sheet.Cells.Item(1,1).Font.Bold= $true
            $report_sheet.Cells.Item(1,1).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,1).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,2) = "From"
            $report_sheet.Cells.Item(1,2).Font.Size=14
            $report_sheet.Cells.Item(1,2).Font.Bold= $true
            $report_sheet.Cells.Item(1,2).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,2).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,3) = "To"
            $report_sheet.Cells.Item(1,3).Font.Size=14
            $report_sheet.Cells.Item(1,3).Font.Bold= $true
            $report_sheet.Cells.Item(1,3).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,3).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,4) = "CC"
            $report_sheet.Cells.Item(1,4).Font.Size=14
            $report_sheet.Cells.Item(1,4).Font.Bold= $true
            $report_sheet.Cells.Item(1,4).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,4).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,5) = "BCC"
            $report_sheet.Cells.Item(1,5).Font.Size=14
            $report_sheet.Cells.Item(1,5).Font.Bold= $true
            $report_sheet.Cells.Item(1,5).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,5).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,6) = "Subject"
            $report_sheet.Cells.Item(1,6).Font.Size=14
            $report_sheet.Cells.Item(1,6).Font.Bold= $true
            $report_sheet.Cells.Item(1,6).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,6).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,7) = "Body"
            $report_sheet.Cells.Item(1,7).Font.Size=14
            $report_sheet.Cells.Item(1,7).Font.Bold= $true
            $report_sheet.Cells.Item(1,7).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,7).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,8) = "Attachments"
            $report_sheet.Cells.Item(1,8).Font.Size=14
            $report_sheet.Cells.Item(1,8).Font.Bold= $true
            $report_sheet.Cells.Item(1,8).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,8).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,9) = "Importance"
            $report_sheet.Cells.Item(1,9).Font.Size=14
            $report_sheet.Cells.Item(1,9).Font.Bold= $true
            $report_sheet.Cells.Item(1,9).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,9).Font.ThemeColor=4

            $report_sheet.Cells.Item(1,10) = "Unread?"
            $report_sheet.Cells.Item(1,10).Font.Size=14
            $report_sheet.Cells.Item(1,10).Font.Bold= $true
            $report_sheet.Cells.Item(1,10).Font.ThemeFont=1
            $report_sheet.Cells.Item(1,10).Font.ThemeColor=4

            $label_User_Info.Text = "Starting to read mails......"

                Foreach($mail_object in $Global:Mails_Read)
                {
                    if($mail_object -ne $null)
                    {
                        #if($mail_object.SenderEmailType -eq "EX") {$user=$mail_object.Sender.GetExchangeUser();$mail_sender = $user.PrimarySmtpAddress}
                        #if($mail_object.SenderEmailType -eq "SMTP") {$mail_sender = $mail_object.SenderEmailAddress}
                        
                        $received_time = ($mail_object.ReceivedTime).ToString()
                        $from = $mail_object.SenderName
                        $TO=$mail_object.To
                        $CC=$mail_object.CC
                        $BCC=$mail_object.BCC
                        $mail_subject = $mail_object.Subject
                        $Body=$mail_object.Body
                        $attachment_names = ""
                        $mail_object.Attachments | ForEach-Object {$attachment_names = $attachment_names + $_.FileName + ";"}
                        $Importance = ($mail_object.Importance).ToString()
                        $unread = $mail_object.Unread

                        $report_sheet.Cells.Item($new_mail_counter ,1) = $received_time ; #$report_sheet.Cells($new_mail_counter ,1).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,2) = $from ;#$report_sheet.Cells($new_mail_counter ,2).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,3) = $TO ;#$report_sheet.Cells($new_mail_counter ,3).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,4) = $CC ;#$report_sheet.Cells($new_mail_counter ,4).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,5) = $BCC ;#$report_sheet.Cells($new_mail_counter ,5).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,6) = $mail_subject ;#$report_sheet.Cells($new_mail_counter ,6).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,7) = $Body ;#$report_sheet.Cells($new_mail_counter ,7).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,8) = $attachment_names ;#$report_sheet.Cells($new_mail_counter ,8).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,9) = $Importance ;#$report_sheet.Cells($new_mail_counter ,9).VerticalAlignmant = -4108
                        $report_sheet.Cells.Item($new_mail_counter ,10) = $unread ;#$report_sheet.Cells($new_mail_counter ,10).VerticalAlignmant = -4108

                    }
                    $progressBar1.PerformStep()
                    $new_mail_counter = $new_mail_counter+1
                }
                

            $Go_button.Enabled = $true
            $label_User_Info.Text = "Saving the file..........."
            $usedRange = $report_sheet.UsedRange
            $Report_Excel.Rows.VerticalAlignment = -4108
            $usedRange.EntireColumn.AutoFit() | Out-Null
            $Report_workbook.SaveAs($Global:ExportFile)
            $Report_Excel.Quit()
            $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Report_Excel)
            [gc]::Collect()
            [gc]::WaitForPendingFinalizers()
            $label_User_Info.Text = "Mails exported successfully as $Global:ExportFile"
            
    }

    
    
   }



      $attachment_button_OnClick ={
        $attachment_files = Get-Multi_FileName "C:\Windows"
        if($attachment_files -ne "")
        {
            foreach ($attachment_file in $attachment_files)
            {
                $attachment_box.Items.Add($attachment_file)
            }
        }
     }

     $attachment_remove_button_OnClick = {
        if ($attachment_box.SelectedItems.Count -gt 0){
            $attachment_box.Items.RemoveAt($attachment_box.SelectedIndex)
        }
     }

     $Heading_TV_NodeMouseDoubleClick = {

        if ($Global:Textbox_To_Selected) {$Textbox_To.SelectedText = "#col"+($_.Node.Index+1)}
        if ($Global:Textbox_CC_Selected) {$Textbox_CC.SelectedText = "#col"+($_.Node.Index+1)}
        if ($Global:Textbox_BCC_Selected) {$Textbox_BCC.SelectedText = "#col"+($_.Node.Index+1)}
        if ($Global:Textbox_Subject_Selected) {$Textbox_Subject.SelectedText = "#col"+($_.Node.Index+1)}
        if ($Global:Textbox_Body_Selected) {$Textbox_Body.SelectedText = "#col"+($_.Node.Index+1)}
     }

     $Textbox_To_OnClick = {
        $Global:Textbox_To_Selected = $True
        $Global:Textbox_CC_Selected = $False
        $Global:Textbox_BCC_Selected = $False
        $Global:Textbox_Subject_Selected = $False
        $Global:Textbox_Body_Selected = $False
     }

     $Textbox_CC_OnClick = {
        $Global:Textbox_To_Selected = $False
        $Global:Textbox_CC_Selected = $True
        $Global:Textbox_BCC_Selected = $False
        $Global:Textbox_Subject_Selected = $False
        $Global:Textbox_Body_Selected = $False
     }

     $Textbox_BCC_OnClick = {
        $Global:Textbox_To_Selected = $False
        $Global:Textbox_CC_Selected = $False
        $Global:Textbox_BCC_Selected = $True
        $Global:Textbox_Subject_Selected = $False
        $Global:Textbox_Body_Selected = $False
     }

     $Textbox_Subject_OnClick = {
        $Global:Textbox_To_Selected = $False
        $Global:Textbox_CC_Selected = $False
        $Global:Textbox_BCC_Selected = $False
        $Global:Textbox_Subject_Selected = $True
        $Global:Textbox_Body_Selected = $False
     }

     $Textbox_Body_OnClick = {
        $Global:Textbox_To_Selected = $False
        $Global:Textbox_CC_Selected = $False
        $Global:Textbox_BCC_Selected = $False
        $Global:Textbox_Subject_Selected = $False
        $Global:Textbox_Body_Selected = $True
     }

     $help_about_OnClick = {
        [System.Windows.MessageBox]::Show($help_Message,'ORA Help','OK','Info')
     }

      $New_Task_Item_OnClick = {
        $label_To.Visible=$true
        $Textbox_To.Visible=$true

        $label_CC.Visible = $false
        $Textbox_CC.Visible=$false
        $label_BCC.Visible = $false
        $Textbox_BCC.Visible=$false

        $source_file_button.Visible = $true
        $export_file_button.Visible = $false

        $Global:new_mail_request = $False
        $Global:new_task_request = $true

        $label_StartDate.Visible=$true
        $label_DueDate.Visible=$true
        $datePicker_StartDate.Visible=$true
        $datePicker_DueDate.Visible=$true
        $label_Search_StartDate.Visible=$False
        $label_Search_DueDate.Visible=$False
        $datePicker_Search_StartDate.Visible=$False
        $datePicker_Search_DueDate.Visible=$False
        

        $attachment_button.Visible=$true
        $attachment_remove_button.Visible=$true
        $attachment_box.Visible=$true

        $check_first_item.Visible=$true
        $send_button.Visible=$true
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $false
        $Go_button.Visible = $False
        $Textbox_Body.Enabled = $True
        $label_heading.Text = "Task Assignment"

        $Mail_treeView.Visible = $False

        $Heading_treeView.Visible = $True
        
        $label_Subject.Visible = $true
        $Textbox_Subject_New.Visible = $false
        $Textbox_Subject.Visible = $true
        $DropDown_Headings.Visible = $false        
        $DropDown_OutlookProfile.Visible = $false
      }

      $New_Mail_Item_OnClick = {
        $label_To.Visible=$true
        $Textbox_To.Visible=$true

        $label_CC.Visible = $true
        $Textbox_CC.Visible=$true
        $label_BCC.Visible = $true
        $Textbox_BCC.Visible=$true

        $source_file_button.Visible = $true
        $export_file_button.Visible = $false

        $Global:new_mail_request = $true
        $Global:new_task_request = $False

        $label_StartDate.Visible=$true
        $label_DueDate.Visible=$true
        $datePicker_StartDate.Visible=$true
        $datePicker_DueDate.Visible=$true
        $label_Search_StartDate.Visible=$False
        $label_Search_DueDate.Visible=$False
        $datePicker_Search_StartDate.Visible=$False
        $datePicker_Search_DueDate.Visible=$False

        $attachment_button.Visible=$true
        $attachment_remove_button.Visible=$true
        $attachment_box.Visible=$true
                
        $check_first_item.Visible=$true
        $send_button.Visible=$true
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $false
        $Go_button.Visible = $False
        $Textbox_Body.Enabled = $True
        $label_heading.Text = "New Mail"

        $Mail_treeView.Visible = $False

        $Heading_treeView.Visible = $True
        
        $label_Subject.Visible = $true
        $Textbox_Subject_New.Visible = $false
        $Textbox_Subject.Visible = $true
        $DropDown_Headings.Visible = $false       
        $DropDown_OutlookProfile.Visible = $false
      }

      $Check_mail_response_OnClick = {
        $label_To.Visible=$false
        $Textbox_To.Visible=$false

        $label_CC.Visible = $false
        $Textbox_CC.Visible=$false
        $label_BCC.Visible = $false
        $Textbox_BCC.Visible=$false

        $source_file_button.Visible = $true
        $export_file_button.Visible = $false

        $label_StartDate.Visible=$false
        $label_DueDate.Visible=$false
        $datePicker_StartDate.Visible=$false
        $datePicker_DueDate.Visible=$false
        $label_Search_StartDate.Visible=$True
        $label_Search_DueDate.Visible=$True
        $datePicker_Search_StartDate.Visible=$True
        $datePicker_Search_DueDate.Visible=$True

        $attachment_button.Visible=$false
        $attachment_remove_button.Visible=$false
        $attachment_box.Visible=$false
               
        $check_first_item.Visible=$false
        $send_button.Visible=$False
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $true
        $Go_button.Visible = $True
        $Textbox_Body.Enabled = $false
        $label_heading.Text = "Mail Replies"
        $CheckBox_Check.Checked = $true 

        $Mail_treeView.Visible = $true

        $Heading_treeView.Visible = $False

        $label_Subject.Visible = $true
        $Textbox_Subject_New.Visible = $True
        $Textbox_Subject.Visible = $false

        $DropDown_Headings.Visible = $true

        $DropDown_OutlookProfile.Visible = $true
     }

     $Check_task_response_OnClick = {
        $label_To.Visible=$false
        $Textbox_To.Visible=$false

        $label_CC.Visible = $false
        $Textbox_CC.Visible=$false
        $label_BCC.Visible = $false
        $Textbox_BCC.Visible=$false

        $source_file_button.Visible = $true
        $export_file_button.Visible = $false

        $label_StartDate.Visible=$false
        $label_DueDate.Visible=$false
        $datePicker_StartDate.Visible=$false
        $datePicker_DueDate.Visible=$false
        $label_Search_StartDate.Visible=$True
        $label_Search_DueDate.Visible=$True
        $datePicker_Search_StartDate.Visible=$True
        $datePicker_Search_DueDate.Visible=$True

        $attachment_button.Visible=$false
        $attachment_remove_button.Visible=$false
        $attachment_box.Visible=$false 
               
        $check_first_item.Visible=$false
        $send_button.Visible=$False
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $true
        $Go_button.Visible = $True
        $Textbox_Body.Enabled = $false
        $label_heading.Text = "Task Responses"
        $CheckBox_Check.Checked = $true 

        $Mail_treeView.Visible = $true

        $Heading_treeView.Visible = $False

        $label_Subject.Visible = $true
        $Textbox_Subject_New.Visible = $True
        $Textbox_Subject.Visible = $false

        $DropDown_Headings.Visible = $true

        $DropDown_OutlookProfile.Visible = $true
     }


     $Set_Replies_OnClick = {
        $label_To.Visible=$False
        $Textbox_To.Visible=$False

        $label_CC.Visible = $true
        $Textbox_CC.Visible=$true
        $label_BCC.Visible = $true
        $Textbox_BCC.Visible=$true

        $source_file_button.Visible = $false
        $export_file_button.Visible = $false

        $Global:new_mail_request = $true
        $Global:new_task_request = $False

        $label_StartDate.Visible=$true
        $label_DueDate.Visible=$true
        $datePicker_StartDate.Visible=$true
        $datePicker_DueDate.Visible=$true
        $label_Search_StartDate.Visible=$False
        $label_Search_DueDate.Visible=$False
        $datePicker_Search_StartDate.Visible=$False
        $datePicker_Search_DueDate.Visible=$False

        $attachment_button.Visible=$true
        $attachment_remove_button.Visible=$true
        $attachment_box.Visible=$true
                
        $check_first_item.Visible=$true
        $send_button.Visible=$false
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $false
        $Go_button.Visible = $False
        $Textbox_Body.Enabled = $True
        $label_heading.Text = "Set Reply"

        $Mail_treeView.Visible = $False

        $Heading_treeView.Visible = $True
        
        $label_Subject.Visible = $true
        $Textbox_Subject_New.Visible = $false
        $Textbox_Subject.Visible = $true
        $DropDown_Headings.Visible = $false       
        $DropDown_OutlookProfile.Visible = $false
     }

     $export_OnClick = {
        $label_To.Visible=$false
        $Textbox_To.Visible=$false

        $label_CC.Visible = $false
        $Textbox_CC.Visible=$false
        $label_BCC.Visible = $false
        $Textbox_BCC.Visible=$false

        $source_file_button.Visible = $false
        $export_file_button.Visible = $true

        $label_StartDate.Visible=$false
        $label_DueDate.Visible=$false
        $datePicker_StartDate.Visible=$false
        $datePicker_DueDate.Visible=$false
        $label_Search_StartDate.Visible=$True
        $label_Search_DueDate.Visible=$True
        $datePicker_Search_StartDate.Visible=$True
        $datePicker_Search_DueDate.Visible=$True

        $attachment_button.Visible=$false
        $attachment_remove_button.Visible=$false
        $attachment_box.Visible=$false
               
        $check_first_item.Visible=$false
        $send_button.Visible=$False
        $check_button.Visible=$True
        $MyGroupBox_Action.Visible = $False
        $Go_button.Visible = $True
        $Textbox_Body.Enabled = $false
        $label_heading.Text = "Export to Excel"
        $CheckBox_Check.Checked = $False

        $Mail_treeView.Visible = $true

        $Heading_treeView.Visible = $False

        $Textbox_Subject_New.Visible = $False
        $Textbox_Subject.Visible = $false
        $label_Subject.Visible = $false

        $DropDown_Headings.Visible = $false

        $DropDown_OutlookProfile.Visible = $true
     }


          
     $CheckBox_Check_OnClick = {
        #$Textbox_Body.Enabled = $False
     }

     $CheckBox_Check_Reply_OnClick = {
        #$Textbox_Body.Enabled = $true
     }


     $Edit_Log_Location_OnClick = {
        $change_Log_Location_To= Get-SaveFolder 'MyComputer'
        if($change_Log_Location_To -ne "")
        {
            Set_Reg_StringValue $Global:App_Reg_Path "LoggingPath" $change_Log_Location_To
            $Global:Log_Location = $change_Log_Location_To
        }
     }

     $Edit_starting_Row_OnClick = {
        $Temp = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter starting row number', 'Staring Row')
        if($Temp -ne "")
        {
            Set_Reg_StringValue $Global:App_Reg_Path "Staring_Row" $Temp
            $Global:Staring_Row= [int]$Temp
        }
     }

     $Edit_Max_Row_To_Count_OnClick = {
        $Temp = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter Maximum Rows to count', 'Maximum Row Count')
        if($Temp -ne "")
        {
            Set_Reg_StringValue $Global:App_Reg_Path "Max_Row_To_Count" $Temp
            $Global:Max_Row_To_Count = [int]$Temp
        }
     }

     $Edit_Max_Column_To_Count_OnClick = {
        $Temp = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter Maximum Column to count', 'Maximum Column')
        if($Temp -ne "")
        {
            Set_Reg_StringValue $Global:App_Reg_Path "Max_Column_To_Count" $Temp
            $Global:Max_Column_To_Count = [int]$Temp
        }
     }

     $Edit_Max_Column_Tolerance_OnClick = {
        $Temp = [Microsoft.VisualBasic.Interaction]::InputBox('Please enter Maximum Column Tolerance', 'Maximum Column Tolerance')
        if($Temp -ne "")
        {
            Set_Reg_StringValue $Global:App_Reg_Path "Max_Column_Tolerance" $Temp
            $Global:Max_Column_Tolerance = [int]$Temp
        }
     }

      
      $check_button_OnClick = {


        if($label_heading.Text -eq "Task Assignment")
        {
            if($Global:inputfile -eq $null){[System.Windows.MessageBox]::Show("Please choose a source Excel file");return}
            if($Textbox_To.text -eq ""){[System.Windows.MessageBox]::Show("Please add recepients in To field");return}
            if($Textbox_Subject.text -eq ""){[System.Windows.MessageBox]::Show("Please add something in subject");return}
            if($Textbox_Body.text -eq ""){[System.Windows.MessageBox]::Show("Please write some mail body");return}
            if($datePicker_DueDate.Value -le $datePicker_StartDate.Value){[System.Windows.MessageBox]::Show("Due date cannot be less than start date");return}
            if($datePicker_DueDate.Value.DayOfYear -eq (Get-Date).DayOfYear){[System.Windows.MessageBox]::Show("Please note: The due date is of today")}
        }

        if($label_heading.Text -eq "New Mail")
        {
            if($Global:inputfile -eq $null){[System.Windows.MessageBox]::Show("Please choose a source Excel file");return}
            if($Textbox_To.text -eq ""){[System.Windows.MessageBox]::Show("Please add recepients in To field");return}
            if($Textbox_Subject.text -eq ""){[System.Windows.MessageBox]::Show("Please add something in subject");return}
            if($Textbox_Body.text -eq ""){[System.Windows.MessageBox]::Show("Please write some mail body");return}
        }

        if($label_heading.Text -eq "Set Reply")
        {
            if($Textbox_Subject.text -eq ""){[System.Windows.MessageBox]::Show("Please add something in subject");return}
            if($Textbox_Body.text -eq ""){[System.Windows.MessageBox]::Show("Please write some mail body");return}
        }

        if( ($label_heading.Text -eq "Mail Replies") -or ($label_heading.Text -eq "Task Responses") )
        {
            $Global:All_Checked_Nodes.Clear()
            Get-Checked_Nodes $Mail_treeView
            Write-Host "Total nodes selected " $Global:All_Checked_Nodes.Count   
            #if($Global:PSTinputfile -eq $null) {[System.Windows.MessageBox]::Show("Please select a PST file"); return}
            if($Global:inputfile -eq $null){[System.Windows.MessageBox]::Show("Please choose a source Excel file"); return}
            if($DropDown_OutlookProfile.SelectedIndex -eq -1) {[System.Windows.MessageBox]::Show("Please select a profile"); return}
            if($Global:All_Checked_Nodes.Count -eq 0) {[System.Windows.MessageBox]::Show("Please select at least one mail folder to check"); return}
            if($DropDown_Headings.SelectedIndex -eq -1) {[System.Windows.MessageBox]::Show("Please select -Check Reply From- field"); return}
            if($Textbox_Subject_New.Text -eq ""){[System.Windows.MessageBox]::Show("Please enter the subject to search for"); return}
            if($datePicker_Search_DueDate.Value -le $datePicker_Search_StartDate.Value){[System.Windows.MessageBox]::Show("End date cannot be less than start date"); return}
            if(  [string]($datePicker_Search_StartDate.Value) -eq [string]($datePicker_Search_DueDate.Value) ) {[System.Windows.MessageBox]::Show("Start date should be less than End date"); return}
        }

        if($label_heading.Text -eq "Export to Excel")
        {
            $Global:All_Checked_Nodes.Clear()
            Get-Checked_Nodes $Mail_treeView
            Write-Host "Total nodes selected " $Global:All_Checked_Nodes.Count   
            #if($Global:PSTinputfile -eq $null) {[System.Windows.MessageBox]::Show("Please select a PST file"); return}
            if($Global:ExportFile -eq $null){[System.Windows.MessageBox]::Show("Please select a file to export the mail contents"); return}
            if($DropDown_OutlookProfile.SelectedIndex -eq -1) {[System.Windows.MessageBox]::Show("Please select a profile"); return}
            if($Global:All_Checked_Nodes.Count -eq 0) {[System.Windows.MessageBox]::Show("Please select at least one mail folder to check"); return}
            if($datePicker_Search_DueDate.Value -le $datePicker_Search_StartDate.Value){[System.Windows.MessageBox]::Show("End date cannot be less than start date"); return}
            if([string]($datePicker_Search_StartDate.Value) -eq [string]($datePicker_Search_DueDate.Value) ) {[System.Windows.MessageBox]::Show("Start date should be less than End date"); return}
            $Go_button.Enabled = $True
            Return
        }


        $to_recepients_list=$Textbox_To.Text
        $CC_recepients_list=$Textbox_CC.Text
        $BCC_recepients_list=$Textbox_BCC.Text
        $mail_subject=$Textbox_Subject.Text
        $mail_body=$Textbox_Body.Text


        $Textbox_To.SelectAll()
        $Textbox_To.SelectionColor = 'black'
        $Textbox_CC.SelectAll()
        $Textbox_CC.SelectionColor = 'black'
        $Textbox_BCC.SelectAll()
        $Textbox_BCC.SelectionColor = 'black'
        $Textbox_Subject.SelectAll()
        $Textbox_Subject.SelectionColor = 'black'
        $Textbox_Body.SelectAll()
        $Textbox_Body.SelectionColor = 'black'

        if($to_recepients_list -ne "")
        {
            Check_Mail_Errors1 $to_recepients_list "to"
            Check_Mail_Errors2 $to_recepients_list "to"
        }

        if($CC_recepients_list -ne "")
        {
            Check_Mail_Errors1 $CC_recepients_list "cc"
            Check_Mail_Errors2 $CC_recepients_list "cc"
        }

        if($BCC_recepients_list -ne "")
        {
            Check_Mail_Errors1 $BCC_recepients_list "bcc"
            Check_Mail_Errors2 $BCC_recepients_list "bcc"
        }

        if($mail_subject -ne "")
        {
            Check_Mail_Errors1 $mail_subject "sub"
            Check_Mail_Errors2 $mail_subject "sub"
        }

        if($mail_body -ne "")
        {
            Check_Mail_Errors1 $mail_body "body"
            Check_Mail_Errors2 $mail_body "body"
        }

        $send_button.Enabled = $True
        $Go_button.Enabled = $True

      }

      $send_button_OnClick = {

        $label_User_Info.Text = ""
        $check_first=0
        if (-not (Test-path "$env:APPDATA\Outlook_Automation")) {New-Item -Type directory -path "$env:APPDATA\Outlook_Automation"}
        $string_sub_text = $Textbox_Subject.Text + "`n"
        $string_sub_text | Out-File -Encoding Ascii -append "$env:APPDATA\Outlook_Automation\History_List.txt"

        #----Opening Word Document for logging----------------------------------------

        $Word = New-Object -ComObject Word.Application
        $Document = $Word.Documents.Add()
        $Selection = $Word.Selection
        $rowMax = $Global:Maximum_Rows_Used
        $starting_row=[int]($Global:heading_row_number+1)
        $progressBar1.Maximum = ($rowMax-$starting_row+1)
        $progressBar1.Value=0
        if($check_first_item.Checked -eq $true){$check_first=1}
        Write_To_Log $Selection "New Report" "Heading 1" "wdColorBlue" "no_bold"

        for($row_temp =$starting_row; $row_temp -le $rowMax; $row_temp++)
        {
          $label_User_Info.Text = "Processing row number $row_temp"
          Write_To_Log $Selection "Processing row number $row_temp" "Normal" "wdColorBlack" "bold"
        #---------------FOR NEW MAIL-----------------------------------
         if($Global:new_mail_request){

            $to_recepients_list=$Textbox_To.Text
            if($to_recepients_list -ne ""){
                $to_recepients_list=Replace_Cols_with_Columns $to_recepients_list $row_temp
                $to_people_number=Number_of_mail_ids $to_recepients_list
                if($to_people_number -le 0) 
                {
                    Write_To_Log $Selection "The To field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }

            $CC_recepients_list=$Textbox_CC.Text
            if($CC_recepients_list -ne ""){
                $CC_recepients_list=Replace_Cols_with_Columns $CC_recepients_list $row_temp
                $cc_people_number=Number_of_mail_ids $CC_recepients_list
                if($cc_people_number -le 0)
                {
                    Write_To_Log $Selection "The CC field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }
            
            $BCC_recepients_list=$Textbox_BCC.Text
            if($BCC_recepients_list -ne ""){
                $BCC_recepients_list=Replace_Cols_with_Columns $BCC_recepients_list $row_temp
                $bcc_people_number=Number_of_mail_ids $BCC_recepients_list
                if($bcc_people_number -le 0)
                {
                    Write_To_Log $Selection "The BCC field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }

            $mail_subject=$Textbox_Subject.Text
            if($mail_subject -ne ""){
                $mail_subject=Replace_Cols_with_Columns $mail_subject $row_temp
                if($mail_subject -eq "") {Write_To_Log $Selection "The subject is blank or cannot be resolved" "Normal" "wdColorRed" "no_bold"}
            }

            $mail_body=$Textbox_Body.Text
            if($mail_body -ne ""){
                $mail_body=Replace_Cols_with_Columns $mail_body $row_temp
                if($mail_body -eq ""){Write_To_Log $Selection "The body is blank or cannot be resolved" "Normal" "wdColorRed" "no_bold"}
            }


            $outlook= New-Object -ComObject outlook.application
            $msg= $outlook.CreateItem(0)
            $msg.Body=$mail_body
            $msg.Subject=$mail_subject
            #----Managing Recepients-------------
            $temp_People_To = Generate_Recepients_List $to_recepients_list
            $msg.To = $temp_People_To
            if ($CC_recepients_list -ne "") {$msg.CC= Generate_Recepients_List $CC_recepients_list}
            if ($BCC_recepients_list -ne ""){$msg.BCC= Generate_Recepients_List $BCC_recepients_list}
            #-----Managing Attachments----------
            if ($attachment_box.Items.Count -gt 0)
            {
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                foreach($attachment in $attachment_box.Items) 
                {
                    $msg.Attachments.Add($attachment.ToString())
                }
            }
            #----------Checking and Sending---------
            if($check_first -eq 1)
            {
                $check_first =0
                $msg.Display()
                $input=[System.Windows.MessageBox]::Show('Is the item correct?','Check Item','YesNo','Info')
                if($input -eq "No"){$label_User_Info.Text = ""; return}
            }
            else
            {
                try
                {
                    $msg.send()
                    Write_To_Log $Selection "Sending mail to $temp_People_To" "Normal" "wdColorBlack" "no_bold"
                }
                catch{Write_To_Log $Selection "There was an issue sending mail for this row number" "Normal" "wdColorRed" "no_bold"}
            }

            $progressBar1.PerformStep()

        }#----Closing If new mail---------

         #---------------FOR NEW TASK-----------------------------------
         if($Global:new_task_request){

            $to_recepients_list=$Textbox_To.Text
            if($to_recepients_list -ne ""){
                $to_recepients_list=Replace_Cols_with_Columns $to_recepients_list $row_temp
                $to_people_number=Number_of_mail_ids $to_recepients_list
                if($to_people_number -le 0)
                {
                    Write_To_Log $Selection "The To field cannot be resolved in proper mail ID/IDs" "Normal" "wdColorRed" "no_bold"
                }
            }

            $start_date=$datePicker_StartDate.Value

            $due_date=$datePicker_DueDate.Value
            

            $mail_subject=$Textbox_Subject.Text
            if($mail_subject -ne ""){
                $mail_subject=Replace_Cols_with_Columns $mail_subject $row_temp
            }

            $mail_body=$Textbox_Body.Text
            if($mail_body -ne ""){
                $mail_body=Replace_Cols_with_Columns $mail_body $row_temp
            }


            $to_assigned_to=Generate_Recepients_List $to_recepients_list
            $to_assigned_to=$to_assigned_to.Substring(0, ($to_assigned_to.length-1))

            $outlook= New-Object -ComObject outlook.application
            $msg= $outlook.CreateItem(3)
            $msg.Assign()
            $msg.Body=$mail_body
            $msg.Subject=$mail_subject
            $msg.Recipients.Add($to_assigned_to)
            $msg.Owner = $to_assigned_to
            $msg.StartDate = $datePicker_StartDate.Value
            $msg.DueDate = $datePicker_DueDate.Value
            #-----Managing Attachments----------
            if ($attachment_box.Items.Count -gt 0)
            {
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                $msg.Body=$msg.Body+"`n"
                foreach($attachment in $attachment_box.Items) 
                {
                    $msg.Attachments.Add($attachment.ToString())
                }
            }
            #----------Checking and Sending---------
            if($check_first -eq 1)
            {
                $check_first =0
                $msg.Display()
                $input=[System.Windows.MessageBox]::Show('Is the item correct?','Check Item','YesNo','Info')
                if($input -eq "No"){$label_User_Info.Text = ""; return}
            }
            else
            {
                try
                {
                    $msg.send()
                    Write_To_Log $Selection "Assigning task request to $to_assigned_to" "Normal" "wdColorBlack" "no_bold"
                }
                catch{Write_To_Log $Selection "There was an issue assigning task for this row number" "Normal" "wdColorRed" "no_bold"}
            }

            $progressBar1.PerformStep()
     
    #------------------------------------------------------
        }#----Closing If task request---------
       }#----Closing for loop---------
        #---Closing and Saving the log file---------------------
        $label_User_Info.Text = "Saving the log file..........."
        $date_time=((Get-Date).ToString()).Replace(":", "-")
        $Report = $Global:Log_Location+"\Sent_Mails_$date_time"
        $Document.SaveAs($Report)
        $Document.close()
        $word.Quit()
        $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Remove-Variable word
        $label_User_Info.Text = "Log file created successfully as $Report" 
     }#----Closing send button function---------

     


     $source_file_button_OnClick = {
        $browse_file = Get-FileName "C:\Windows"
        if($browse_file -ne "")
        {
            $DropDown_Headings.Enabled = $true
            $Heading_treeView.Nodes.Clear()
            $Global:inputfile = $browse_file
            #$label_Choose_Source_File.Text = $inputfile
            $pictureBox.Visible = $true
            $label_User_Info.Text = "Please wait while your Excel headings are being populated............"
            $Heading_treeView.Enabled = $true
            $Global:objExcel = New-Object -ComObject Excel.Application
            $Global:workbook = $Global:objExcel.Workbooks.Open($Global:inputfile)
            $sheet = $Global:workbook.Worksheets.Item(1)
            $Global:objExcel.Visible=$false
            $Global:My_sheet = $sheet
            Populate_Heading_Tree $sheet $Heading_treeView
            $pictureBox.Visible = $false
            $label_User_Info.Text = ""
            #$Global:objExcel.quit()

            #--------Populating the PST profiles------------------
            $DropDown_OutlookProfile.Enabled = $True
            $DropDown_OutlookProfile.Items.Clear()
            $oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
            if ( $oProc -eq $null ) { Start-Process outlook -WindowStyle Hidden; Start-Sleep -Seconds 5 }
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $Global:NS = $namespace
            #$namespace.AddStoreEx($pstPath, "olStoreDefault")
            $pstStores = $nameSpace.Stores
            ForEach ($pstStore in $pstStores) 
            {
                $pstRootFolder = $pstStore.GetRootFolder()
                if( ($pstStore.FilePath).EndsWith(".pst") ) {$DropDown_OutlookProfile.Items.Add($pstStore.DisplayName)}
            }
        }
        
      }


      $export_file_button_OnClick = {
        $export_file = Get-SaveFile "C:\Windows"
        if($export_file -ne "")
        {
            $Global:ExportFile=$export_file
            $label_User_Info.Text = $export_file
            #--------Populating the PST profiles------------------
            $DropDown_OutlookProfile.Enabled = $True
            $DropDown_OutlookProfile.Items.Clear()
            $oProc = ( Get-Process | where { $_.Name -eq "OUTLOOK" } )
            if ( $oProc -eq $null ) { Start-Process outlook -WindowStyle Hidden; Start-Sleep -Seconds 5 }
            $outlook = New-Object -ComObject Outlook.Application
            $namespace = $outlook.GetNamespace("MAPI")
            $Global:NS = $namespace
            #$namespace.AddStoreEx($pstPath, "olStoreDefault")
            $pstStores = $nameSpace.Stores
            ForEach ($pstStore in $pstStores) 
            {
                $pstRootFolder = $pstStore.GetRootFolder()
                if( ($pstStore.FilePath).EndsWith(".pst") ) {$DropDown_OutlookProfile.Items.Add($pstStore.DisplayName)}
            }
        }
        
      }



      $DropDown_OutlookProfile_SelectedIndexChanged = {
        #Write-Host $DropDown_OutlookProfile.SelectedItem
        $pstStores = $Global:NS.Stores
        ForEach($pstStore in $pstStores)
        {
            if ($pstStore.DisplayName -eq $DropDown_OutlookProfile.SelectedItem) {$Global:PSTinputfile= $pstStore.FilePath}
        }
            $Mail_treeView.Nodes.Clear()          
            $Mail_treeView.Enabled = $True
            $pictureBox.Visible = $true
            $label_User_Info.Text = "Please wait while your mail folders are being populated..........."
            [string]$pstPath = $Global:PSTinputfile
            $namespace = $Global:NS
            $namespace.AddStoreEx($pstPath, "olStoreDefault")
            $pstStore = ( $nameSpace.Stores | where { $_.FilePath -eq $pstPath } )
            $pstRootFolder = $pstStore.GetRootFolder()
            Fill_Tree $Mail_treeView $pstRootFolder
            $pictureBox.Visible = $false
            $label_User_Info.Text = ""
      }


     $form_main = New-Object System.Windows.Forms.Form
     $progressBar1 = New-Object System.Windows.Forms.ProgressBar
     $InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
         
     $MS_Main = new-object System.Windows.Forms.MenuStrip

     $Sending_Items = new-object System.Windows.Forms.ToolStripMenuItem
     $New_Mail_Item = new-object System.Windows.Forms.ToolStripMenuItem
     $New_Task_Item = new-object System.Windows.Forms.ToolStripMenuItem
     $Tools=new-object System.Windows.Forms.ToolStripMenuItem
     $help = new-object System.Windows.Forms.ToolStripMenuItem

     $Checking_Items = new-object System.Windows.Forms.ToolStripMenuItem
     $Check_mail_response = new-object System.Windows.Forms.ToolStripMenuItem
     $Check_task_response = new-object System.Windows.Forms.ToolStripMenuItem
     $Set_Replies = new-object System.Windows.Forms.ToolStripMenuItem

     $Edit_Items = new-object System.Windows.Forms.ToolStripMenuItem
     $Edit_Log_Location = new-object System.Windows.Forms.ToolStripMenuItem

     $Excel_Data = new-object System.Windows.Forms.ToolStripMenuItem
     $Edit_starting_Row = new-object System.Windows.Forms.ToolStripMenuItem
     $Edit_Max_Row_To_Count = new-object System.Windows.Forms.ToolStripMenuItem
     $Edit_Max_Column_To_Count = new-object System.Windows.Forms.ToolStripMenuItem
     $Edit_Max_Column_Tolerance = new-object System.Windows.Forms.ToolStripMenuItem

     $write_to_excel=new-object System.Windows.Forms.ToolStripMenuItem

     $help_about = new-object System.Windows.Forms.ToolStripMenuItem
    #
    # MS_Main
    #
    $MS_Main.Items.AddRange(@(
    $Sending_Items,
    $Checking_Items,
    $Edit_Items,
    $Tools,
    $help))

    $MS_Main.Location = new-object System.Drawing.Point(0, 0)
    $MS_Main.Name = "MS_Main"
    $MS_Main.Size = new-object System.Drawing.Size(354, 24)
    $MS_Main.TabIndex = 0
    $MS_Main.BackColor ="LightBlue"
    #$MS_Main.BackgroundImage = $element_backImage
    $MS_Main.Text = "menuStrip1"

    $Sending_Items.Name = "Sending items"
    $Sending_Items.Size = new-object System.Drawing.Size(35, 20)
    $Sending_Items.Text = "&New"
    $Sending_Items.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Sending_Items.DropDownItems.AddRange(@($New_Mail_Item,$New_Task_Item ))

    $New_Mail_Item.Name = "New mail"
    $New_Mail_Item.Size = new-object System.Drawing.Size(35, 20)
    $New_Mail_Item.Text = "&New Mail"
    $New_Mail_Item.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $New_Mail_Item.add_Click($New_Mail_Item_OnClick)

    $New_Task_Item.Name = "New Task"
    $New_Task_Item.Size = new-object System.Drawing.Size(35, 20)
    $New_Task_Item.Text = "&Task Request"
    $New_Task_Item.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $New_Task_Item.add_Click($New_Task_Item_OnClick)

    $Checking_Items.Name = "CheckingItems"
    $Checking_Items.Size = new-object System.Drawing.Size(51, 20)
    $Checking_Items.Text = "&Check Reply"
    $Checking_Items.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Checking_Items.DropDownItems.AddRange(@($Check_mail_response,$Check_task_response,$Set_Replies ))

    $Check_mail_response.Name = "Check Mail"
    $Check_mail_response.Size = new-object System.Drawing.Size(35, 20)
    $Check_mail_response.Text = "&Check mail reply"
    $Check_mail_response.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Check_mail_response.add_Click($Check_mail_response_OnClick)

    $Check_task_response.Name = "Check Task Response"
    $Check_task_response.Size = new-object System.Drawing.Size(35, 20)
    $Check_task_response.Text = "&Check Task status"
    $Check_task_response.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Check_task_response.Enabled = $False
    $Check_task_response.add_Click($Check_task_response_OnClick)

    $Set_Replies.Name = "Set Replies"
    $Set_Replies.Size = new-object System.Drawing.Size(35, 20)
    $Set_Replies.Text = "&Set Reply"
    $Set_Replies.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Set_Replies.add_Click($Set_Replies_OnClick)

    $Edit_Items.Name = "Editing items"
    $Edit_Items.Size = new-object System.Drawing.Size(35, 20)
    $Edit_Items.Text = "&Edit"
    $Edit_Items.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_Items.DropDownItems.AddRange(@($Edit_Log_Location, $Excel_Data))

    $Edit_Log_Location.Name = "Edit Log Location"
    $Edit_Log_Location.Size = new-object System.Drawing.Size(35, 20)
    $Edit_Log_Location.Text = "&Change Log Location"
    $Edit_Log_Location.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_Log_Location.add_Click($Edit_Log_Location_OnClick)

    $Excel_Data.Name = "ExcelData"
    $Excel_Data.Size = new-object System.Drawing.Size(35, 20)
    $Excel_Data.Text = "&Excel Data"
    $Excel_Data.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Excel_Data.DropDownItems.AddRange(@($Edit_starting_Row, $Edit_Max_Row_To_Count, $Edit_Max_Column_To_Count, $Edit_Max_Column_Tolerance))

    $Edit_starting_Row.Name = "Starting Row"
    $Edit_starting_Row.Size = new-object System.Drawing.Size(35, 20)
    $Edit_starting_Row.Text = "&Set Starting Row"
    $Edit_starting_Row.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_starting_Row.add_Click($Edit_starting_Row_OnClick)

    $Edit_Max_Row_To_Count.Name = "Set Max Row Count"
    $Edit_Max_Row_To_Count.Size = new-object System.Drawing.Size(35, 20)
    $Edit_Max_Row_To_Count.Text = "&Set Max Row Count"
    $Edit_Max_Row_To_Count.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_Max_Row_To_Count.add_Click($Edit_Max_Row_To_Count_OnClick)

    $Edit_Max_Column_To_Count.Name = "Set Max Column Count"
    $Edit_Max_Column_To_Count.Size = new-object System.Drawing.Size(35, 20)
    $Edit_Max_Column_To_Count.Text = "&Set Max Column Count"
    $Edit_Max_Column_To_Count.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_Max_Column_To_Count.add_Click($Edit_Max_Column_To_Count_OnClick)

    $Edit_Max_Column_Tolerance.Name = "Set Max Column Tolerance"
    $Edit_Max_Column_Tolerance.Size = new-object System.Drawing.Size(35, 20)
    $Edit_Max_Column_Tolerance.Text = "&Set Max Column Tolerance"
    $Edit_Max_Column_Tolerance.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Edit_Max_Column_Tolerance.add_Click($Edit_Max_Column_Tolerance_OnClick)

    $Tools.Name = "Tools"
    $Tools.Size = new-object System.Drawing.Size(51, 20)
    $Tools.Text = "&Tools"
    $Tools.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $Tools.DropDownItems.AddRange(@($write_to_excel))

    $write_to_excel.Name = "Export_to_Excel"
    $write_to_excel.Size = new-object System.Drawing.Size(35, 20)
    $write_to_excel.Text = "&Export to Excel"
    $write_to_excel.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $write_to_excel.add_Click($export_OnClick)

    $help.Name = "Help"
    $help.Size = new-object System.Drawing.Size(51, 20)
    $help.Text = "&Help"
    $help.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $help.DropDownItems.AddRange(@($help_about))

    $help_about.Name = "Set Max Column Count"
    $help_about.Size = new-object System.Drawing.Size(35, 20)
    $help_about.Text = "&About"
    $help_about.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
    $help_about.add_Click($help_about_OnClick)

      $OnLoadForm_StateCorrection = {
      #Correct the initial state of the form to prevent the .Net maximized form issue
      $form_main.WindowState = $InitialFormWindowState
      $form_main.Icon = $Icon
      }

      $form_main.MaximumSize = New-Object System.Drawing.Size((40*$w/30),(35*$h/30))
      $form_main.MinimumSize = New-Object System.Drawing.Size((40*$w/30),(35*$h/30))
      $form_main.ClientSize = New-Object System.Drawing.Size((40*$w/30),(35*$h/30))
      $form_main.Text = 'Outlook Report Assistant'
      $form_main.ControlBox = $true
      $form_main.Name = 'form_main'
      #$form_main.ShowIcon = $False
      $form_main.StartPosition = 1
      $form_main.DataBindings.DefaultDataSourceUpdateMode = 0
      $form_main.Controls.Add($MS_Main)
      $form_main.MainMenuStrip = $MS_Main
      $form_main.BackgroundImage=$background_image
      $form_main.BackgroundImageLayout = "Center"
      #$form_main.FormBorderStyle = 'None'     
      $form_main.add_Closing($form_main_add_Closing)

      $pictureBox = new-object Windows.Forms.PictureBox
      $pictureBox.Size = New-Object System.Drawing.Size($Image.Size.Width,$Image.Size.Height)
      $pictureBox.Location = New-Object System.Drawing.Point((10*$w/30),(15*$h/30))
      $pictureBox.Image = $Image
      $pictureBox.Visible = $False
      $form_main.Controls.Add($pictureBox)
      
      $label_heading = New-Object System.Windows.Forms.Label
      $label_heading.TabIndex = 1
      $label_heading.Size = New-Object System.Drawing.Size((20*$w/30),(2*$h/30))
      $label_heading.Location = New-Object System.Drawing.Point((10*$w/30),($h/30))
      $label_heading.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/50))
      $label_heading.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_heading.Name = 'label_heading'
      $label_heading.Text = "New Mail"
      $label_heading.TextAlign = "BottomCenter"
      $label_heading.BackColor="Transparent"
      $form_main.Controls.Add($label_heading)


      $label_Choose_Source_File = New-Object System.Windows.Forms.Label
      $label_Choose_Source_File.TabIndex = 2
      $label_Choose_Source_File.Size = New-Object System.Drawing.Size((14*$w/30),(3*$h/60))
      $label_Choose_Source_File.Location = New-Object System.Drawing.Point((2*$w/30),(5*$h/30))
      $label_Choose_Source_File.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_Choose_Source_File.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_Choose_Source_File.Name = 'label_Source_File'
      $label_Choose_Source_File.Text = "Choose source File"
      $label_Choose_Source_File.TextAlign = "MiddleLeft"
      $label_Choose_Source_File.BackColor="Transparent"
      #$form_main.Controls.Add($label_Choose_Source_File)

     
      $source_file_button = New-Object System.Windows.Forms.Button
      $source_file_button.TabIndex = 3
      $source_file_button.Name = 'button1'
      $source_file_button.Size = New-Object System.Drawing.Size((6*$w/30),(3*$h/60))
      $source_file_button.Location = New-Object System.Drawing.Point((2*$w/30),(5*$h/30))
      $source_file_button.UseVisualStyleBackColor = $True
      $source_file_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $source_file_button.Text = "Browse Source File"
      $source_file_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $source_file_button.add_Click($source_file_button_OnClick)
      $form_main.Controls.Add($source_file_button)

      $export_file_button = New-Object System.Windows.Forms.Button
      $export_file_button.TabIndex = 3
      $export_file_button.Name = 'button1'
      $export_file_button.Size = New-Object System.Drawing.Size((6*$w/30),(3*$h/60))
      $export_file_button.Location = New-Object System.Drawing.Point((2*$w/30),(5*$h/30))
      $export_file_button.UseVisualStyleBackColor = $True
      $export_file_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $export_file_button.Text = "Export to File"
      $export_file_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $export_file_button.Visible = $False
      $export_file_button.add_Click($export_file_button_OnClick)
      $form_main.Controls.Add($export_file_button)


      #-----------Group Box for action type-----------------------------
      $MyGroupBox_Action = New-Object System.Windows.Forms.GroupBox
      $MyGroupBox_Action.TabIndex = 4
      $MyGroupBox_Action.Location = New-Object System.Drawing.Point((20*$w/30),(9*$h/30))
      $MyGroupBox_Action.size = New-Object System.Drawing.Size((7*$w/30),(8*$h/60))
      $MyGroupBox_Action.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $MyGroupBox_Action.text = "Select Action"
      $MyGroupBox_Action.Visible = $False
      $MyGroupBox_Action.BackColor = "Transparent"
      $form_main.Controls.Add($MyGroupBox_Action)

      # Create the collection of radio buttons
      $CheckBox_Check = New-Object System.Windows.Forms.CheckBox
      $CheckBox_Check.TabIndex = 5
      $CheckBox_Check.Location = New-Object System.Drawing.Point(10,25)
      $CheckBox_Check.size = New-Object System.Drawing.Size(170,20)
      $CheckBox_Check.Checked = $true 
      $CheckBox_Check.Text = "Check Response."
      $CheckBox_Check.add_Click($CheckBox_Check_OnClick)
 
      $CheckBox_Check_Reply = New-Object System.Windows.Forms.CheckBox
      $CheckBox_Check_Reply.TabIndex = 6
      $CheckBox_Check_Reply.Location = New-Object System.Drawing.Point(10,50)
      $CheckBox_Check_Reply.size = New-Object System.Drawing.Size(170,20)
      $CheckBox_Check_Reply.Checked = $false
      $CheckBox_Check_Reply.Text = "Send Replies"
      $CheckBox_Check_Reply.add_Click($CheckBox_Check_Reply_OnClick)

      $MyGroupBox_Action.Controls.AddRange(@($CheckBox_Check,$CheckBox_Check_Reply))

      #-------------Attachments--------------------------------------------

      $attachment_button = New-Object System.Windows.Forms.Button
      $attachment_button.TabIndex = 32
      $attachment_button.Name = 'button1'
      $attachment_button.Size = New-Object System.Drawing.Size((6*$w/30),(3*$h/60))
      $attachment_button.Location = New-Object System.Drawing.Point((2*$w/30),(7*$h/30))
      $attachment_button.UseVisualStyleBackColor = $True
      $attachment_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $attachment_button.Text = "Add Attachments"
      $attachment_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $attachment_button.add_Click($attachment_button_OnClick)
      $form_main.Controls.Add($attachment_button)

      $attachment_remove_button = New-Object System.Windows.Forms.Button
      $attachment_remove_button.TabIndex = 34
      $attachment_remove_button.Name = 'button11'
      $attachment_remove_button.Size = New-Object System.Drawing.Size((6*$w/30),(3*$h/60))
      $attachment_remove_button.Location = New-Object System.Drawing.Point((10*$w/30),(7*$h/30))
      $attachment_remove_button.UseVisualStyleBackColor = $True
      $attachment_remove_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $attachment_remove_button.Text = "Remove Attachments"
      $attachment_remove_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $attachment_remove_button.add_Click($attachment_remove_button_OnClick)
      $form_main.Controls.Add($attachment_remove_button)

      $attachment_box = New-Object System.Windows.Forms.ListBox
      $attachment_box.TabIndex = 33
      $attachment_box.Size = New-Object System.Drawing.Size((14*$w/30),(5*$h/60))
      $attachment_box.Location = New-Object System.Drawing.Point((2*$w/30),(9*$h/30))
      $form_main.Controls.Add($attachment_box)

      #---------------------------------------------------------------------------

      $Mail_treeView = New-Object System.Windows.Forms.TreeView
      $Mail_treeView.Size = New-Object System.Drawing.Size((15*$w/60),(45*$h/60))
      $Mail_treeView.Location = New-Object System.Drawing.Point((31*$w/30),(14*$h/60))
      $Mail_treeView.CheckBoxes = $true
      $Mail_treeView.Enabled = $False
      $Mail_treeView.Add_AfterCheck($TV_AfterCheck)
      $Mail_treeView.Visible = $False
      $Mail_treeView.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $form_main.Controls.Add($Mail_treeView)

      $Heading_treeView = New-Object System.Windows.Forms.TreeView
      $Heading_treeView.Size = New-Object System.Drawing.Size((15*$w/60),(45*$h/60))
      $Heading_treeView.Location = New-Object System.Drawing.Point((31*$w/30),(14*$h/60))
      $Heading_treeView.Enabled = $False
      $Heading_treeView.Add_AfterCheck($TV_AfterCheck)
      $Heading_treeView.Visible = $True
      $Heading_treeView.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $Heading_treeView.Add_NodeMouseDoubleClick($Heading_TV_NodeMouseDoubleClick)
      $form_main.Controls.Add($Heading_treeView)



      $label_To = New-Object System.Windows.Forms.Label
      $label_To.TabIndex = 7
      $label_To.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_To.Location = New-Object System.Drawing.Point((17*$w/30),(29*$h/120))
      $label_To.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_To.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_To.Name = 'label_To'
      $label_To.Text = "To"
      $label_To.TextAlign = "MiddleLeft"
      $label_To.BackColor = "Transparent"
      $label_To.BorderStyle = "FixedSingle"
      $form_main.Controls.Add($label_To)

      $Textbox_To = New-Object System.Windows.Forms.RichTextBox
      $Textbox_To.TabIndex = 8
      $Textbox_To.Location = New-Object System.Drawing.Point((20*$w/30),(29*$h/120))
      $Textbox_To.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $Textbox_To.add_Click($Textbox_To_OnClick)
      $form_main.Controls.Add($Textbox_To)


      $label_CC = New-Object System.Windows.Forms.Label
      $label_CC.TabIndex = 9
      $label_CC.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_CC.Location = New-Object System.Drawing.Point((17*$w/30),(34*$h/120))
      $label_CC.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_CC.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_CC.Name = 'label_CC'
      $label_CC.Text = "CC"
      $label_CC.TextAlign = "MiddleLeft"
      $label_CC.BackColor="Transparent"
      $label_CC.BorderStyle = "FixedSingle"
      $form_main.Controls.Add($label_CC)

      $Textbox_CC = New-Object System.Windows.Forms.RichTextBox
      $Textbox_CC.TabIndex = 10
      $Textbox_CC.Location = New-Object System.Drawing.Point((20*$w/30),(34*$h/120))
      $Textbox_CC.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $Textbox_CC.add_Click($Textbox_CC_OnClick)
      $form_main.Controls.Add($Textbox_CC)

      $label_BCC = New-Object System.Windows.Forms.Label
      $label_BCC.TabIndex = 11
      $label_BCC.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_BCC.Location = New-Object System.Drawing.Point((17*$w/30),(39*$h/120))
      $label_BCC.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_BCC.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_BCC.Name = 'label_BCC'
      $label_BCC.Text = "BCC"
      $label_BCC.TextAlign = "MiddleLeft"
      $label_BCC.BackColor="Transparent"
      $label_BCC.BorderStyle = "FixedSingle"
      $form_main.Controls.Add($label_BCC)

      $Textbox_BCC = New-Object System.Windows.Forms.RichTextBox
      $Textbox_BCC.TabIndex = 12
      $Textbox_BCC.Location = New-Object System.Drawing.Point((20*$w/30),(39*$h/120))
      $Textbox_BCC.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $Textbox_BCC.add_Click($Textbox_BCC_OnClick)
      $form_main.Controls.Add($Textbox_BCC)

      $check_first_item = New-Object System.Windows.Forms.Checkbox 
      $check_first_item.Location = New-Object System.Drawing.Point((20*$w/30),(45*$h/120))
      $check_first_item.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $check_first_item.Text = "I want to check the first item"
      $check_first_item.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $check_first_item.TabIndex = 31
      $check_first_item.BackColor = "Transparent"
      $form_main.Controls.Add($check_first_item)

      #----------------------------------------------------------------------
      $label_StartDate = New-Object System.Windows.Forms.Label
      $label_StartDate.TabIndex = 13
      $label_StartDate.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_StartDate.Location = New-Object System.Drawing.Point((17*$w/30),(34*$h/120))
      $label_StartDate.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_StartDate.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_StartDate.Name = 'label_StartDate'
      $label_StartDate.Text = "Start Date"
      $label_StartDate.TextAlign = "MiddleLeft"
      $label_StartDate.BackColor="Transparent"
      $label_StartDate.BorderStyle = "FixedSingle"
      $label_StartDate.Visible=$false
      $form_main.Controls.Add($label_StartDate)

      $label_DueDate = New-Object System.Windows.Forms.Label
      $label_DueDate.TabIndex = 14
      $label_DueDate.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_DueDate.Location = New-Object System.Drawing.Point((17*$w/30),(39*$h/120))
      $label_DueDate.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_DueDate.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_DueDate.Name = 'label_DueDate'
      $label_DueDate.Text = "Due Date"
      $label_DueDate.TextAlign = "MiddleLeft"
      $label_DueDate.BackColor="Transparent"
      $label_DueDate.BorderStyle = "FixedSingle"
      $label_DueDate.Visible=$false
      $form_main.Controls.Add($label_DueDate)

      $datePicker_StartDate = New-Object Windows.Forms.DateTimePicker
      $datePicker_StartDate.ShowUpDown = $false
      $datePicker_StartDate.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $datePicker_StartDate.Location = New-Object System.Drawing.Point((20*$w/30),(34*$h/120))
      $datePicker_StartDate.Visible=$False
      $form_main.Controls.Add($datePicker_StartDate)

      $datePicker_DueDate = New-Object Windows.Forms.DateTimePicker
      $datePicker_DueDate.ShowUpDown = $false
      $datePicker_DueDate.Size = New-Object System.Drawing.Size((10*$w/30),(2*$h/60))
      $datePicker_DueDate.Location = New-Object System.Drawing.Point((20*$w/30),(39*$h/120))
      $datePicker_DueDate.Visible=$False
      $form_main.Controls.Add($datePicker_DueDate)

      #--------------------------------------------------------------------------------

      $label_Search_StartDate = New-Object System.Windows.Forms.Label
      $label_Search_StartDate.TabIndex = 51
      $label_Search_StartDate.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_Search_StartDate.Location = New-Object System.Drawing.Point((2*$w/30),(37*$h/120))
      $label_Search_StartDate.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_Search_StartDate.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_Search_StartDate.Name = 'label_StartDate'
      $label_Search_StartDate.Text = "Start Date"
      $label_Search_StartDate.TextAlign = "MiddleLeft"
      $label_Search_StartDate.BackColor="Transparent"
      $label_Search_StartDate.BorderStyle = "FixedSingle"
      $label_Search_StartDate.Visible=$false
      $form_main.Controls.Add($label_Search_StartDate)

      $label_Search_DueDate = New-Object System.Windows.Forms.Label
      $label_Search_DueDate.TabIndex = 52
      $label_Search_DueDate.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_Search_DueDate.Location = New-Object System.Drawing.Point((2*$w/30),(42*$h/120))
      $label_Search_DueDate.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_Search_DueDate.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_Search_DueDate.Name = 'label_DueDate'
      $label_Search_DueDate.Text = "End Date"
      $label_Search_DueDate.TextAlign = "MiddleLeft"
      $label_Search_DueDate.BackColor="Transparent"
      $label_Search_DueDate.BorderStyle = "FixedSingle"
      $label_Search_DueDate.Visible=$false
      $form_main.Controls.Add($label_Search_DueDate)

      $datePicker_Search_StartDate = New-Object Windows.Forms.DateTimePicker
      $datePicker_Search_StartDate.ShowUpDown = $false
      $datePicker_Search_StartDate.Size = New-Object System.Drawing.Size((11*$w/30),(2*$h/60))
      $datePicker_Search_StartDate.Location = New-Object System.Drawing.Point((5*$w/30),(37*$h/120))
      $datePicker_Search_StartDate.Visible=$False
      $form_main.Controls.Add($datePicker_Search_StartDate)

      $datePicker_Search_DueDate = New-Object Windows.Forms.DateTimePicker
      $datePicker_Search_DueDate.ShowUpDown = $false
      $datePicker_Search_DueDate.Size = New-Object System.Drawing.Size((11*$w/30),(2*$h/60))
      $datePicker_Search_DueDate.Location = New-Object System.Drawing.Point((5*$w/30),(42*$h/120))
      $datePicker_Search_DueDate.Visible=$False
      $form_main.Controls.Add($datePicker_Search_DueDate)

      #------------------------------------------------------------------------

      $label_Subject = New-Object System.Windows.Forms.Label
      $label_Subject.TabIndex = 15
      $label_Subject.Size = New-Object System.Drawing.Size((5*$w/60),(2*$h/60))
      $label_Subject.Location = New-Object System.Drawing.Point((2*$w/30),(12*$h/30))
      $label_Subject.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_Subject.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_Subject.Name = 'label_Subject'
      $label_Subject.Text = "Subject"
      $label_Subject.TextAlign = "MiddleLeft"
      $label_Subject.BackColor="Transparent"
      $label_Subject.BorderStyle = "FixedSingle"
      $form_main.Controls.Add($label_Subject)

      $Textbox_Subject = New-Object System.Windows.Forms.RichTextBox
      $Textbox_Subject.TabIndex = 16
      $Textbox_Subject.Location = New-Object System.Drawing.Point((5*$w/30),(12*$h/30))
      $Textbox_Subject.Size = New-Object System.Drawing.Size((11*$w/30),(2*$h/60))
      $Textbox_Subject.add_Click($Textbox_Subject_OnClick)
      $form_main.Controls.Add($Textbox_Subject)

      $Textbox_Subject_New = New-Object System.Windows.Forms.TextBox
      $Textbox_Subject_New.TabIndex = 16
      $Textbox_Subject_New.Location = New-Object System.Drawing.Point((5*$w/30),(12*$h/30))
      $Textbox_Subject_New.Size = New-Object System.Drawing.Size((11*$w/30),(2*$h/60))
      $Textbox_Subject_New.add_Click($Textbox_Subject_New_OnClick)
      $Textbox_Subject_New.Visible = $false
      $Textbox_Subject_New.AutoCompleteSource = 'CustomSource'
      $Textbox_Subject_New.AutoCompleteMode = 'SuggestAppend'
      Get-content "$env:APPDATA\Outlook_Automation\History_List.txt" | % {$Textbox_Subject_New.AutoCompleteCustomSource.AddRange($_) }
      $form_main.Controls.Add($Textbox_Subject_New)

      $DropDown_Headings = new-object System.Windows.Forms.ComboBox
      $DropDown_Headings.TabIndex = 37
      $DropDown_Headings.Size = New-Object System.Drawing.Size((56*$w/120),($h/22))
      $DropDown_Headings.Location = New-Object System.Drawing.Point((2*$w/30),(30*$h/120))
      $DropDown_Headings.Enabled = $False
      $DropDown_Headings.Visible = $false
      $DropDown_Headings.Text = "Check Reply From"
      $DropDown_Headings.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $form_main.Controls.Add($DropDown_Headings)

      $DropDown_OutlookProfile = new-object System.Windows.Forms.ComboBox
      $DropDown_OutlookProfile.TabIndex = 38
      $DropDown_OutlookProfile.Size = New-Object System.Drawing.Size((15*$w/60),($h/22))
      $DropDown_OutlookProfile.Location = New-Object System.Drawing.Point((31*$w/30),(5*$h/30))
      $DropDown_OutlookProfile.Enabled = $False
      $DropDown_OutlookProfile.Visible = $false
      $DropDown_OutlookProfile.Text = "Select Outlook Profile"
      $DropDown_OutlookProfile.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $DropDown_OutlookProfile.add_SelectedIndexChanged($DropDown_OutlookProfile_SelectedIndexChanged)
      $form_main.Controls.Add($DropDown_OutlookProfile)

      $Textbox_Body = New-Object System.Windows.Forms.RichTextBox
      $Textbox_Body.TabIndex = 17
      $Textbox_Body.Location = New-Object System.Drawing.Point((2*$w/30),(27*$h/60))
      $Textbox_Body.Size = New-Object System.Drawing.Size((28*$w/30),(16*$h/30))
      $Textbox_Body.add_Click($Textbox_Body_OnClick)
      $form_main.Controls.Add($Textbox_Body)


      $send_button = New-Object System.Windows.Forms.Button
      $send_button.TabIndex = 18
      $send_button.Name = 'button3'
      $send_button.Size = New-Object System.Drawing.Size((4*$w/30),(5*$h/60))
      $send_button.Location = New-Object System.Drawing.Point((2*$w/30),(2*$h/30))
      $send_button.UseVisualStyleBackColor = $True
      $send_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $send_button.Text = "Send"
      $send_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $send_button.Enabled = $False
      $send_button.add_Click($send_button_OnClick)
      $form_main.Controls.Add($send_button)

      $Go_button = New-Object System.Windows.Forms.Button
      $Go_button.TabIndex = 42
      $Go_button.Name = 'button6'
      $Go_button.Size = New-Object System.Drawing.Size((4*$w/30),(5*$h/60))
      $Go_button.Location = New-Object System.Drawing.Point((2*$w/30),(2*$h/30))
      $Go_button.UseVisualStyleBackColor = $True
      $Go_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $Go_button.Text = "Go"
      $Go_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $Go_button.Enabled = $False
      $Go_button.Visible = $False
      $Go_button.add_Click($Go_button_OnClick)
      $form_main.Controls.Add($Go_button)

      $check_button = New-Object System.Windows.Forms.Button
      $check_button.TabIndex = 19
      $check_button.Name = 'button3'
      $check_button.Size = New-Object System.Drawing.Size((5*$w/30),(5*$h/60))
      $check_button.Location = New-Object System.Drawing.Point((31*$w/30),(2*$h/30))
      $check_button.UseVisualStyleBackColor = $True
      $check_button.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $check_button.Text = "Check for Errors"
      $check_button.DataBindings.DefaultDataSourceUpdateMode = 0
      $check_button.add_Click($check_button_OnClick)
      $form_main.Controls.Add($check_button)

     
      $progressBar1.DataBindings.DefaultDataSourceUpdateMode = 0
      $progressBar1.Size = New-Object System.Drawing.Size((28*$w/30),(6*$h/120))
      $progressBar1.Location = New-Object System.Drawing.Point((2*$w/30),(60*$h/60))
      $progressBar1.Step = 1
      $progressBar1.TabIndex = 0
      $progressBar1.Style = 1
      $progressBar1.Name = 'progressBar1'
      $form_main.Controls.Add($progressBar1)

      $label_User_Info = New-Object System.Windows.Forms.Label
      $label_User_Info.TabIndex = 50
      $label_User_Info.Size = New-Object System.Drawing.Size((28*$w/30),(5*$h/120))
      $label_User_Info.Location = New-Object System.Drawing.Point((2*$w/30),(63*$h/60))
      $label_User_Info.Font = New-Object System.Drawing.Font("Trebuchet MS",($w/90))
      $label_User_Info.DataBindings.DefaultDataSourceUpdateMode = 0
      $label_User_Info.Name = 'label_Info'
      $label_User_Info.TextAlign = "MiddleLeft"
      $label_User_Info.BackColor="Transparent"
      $form_main.Controls.Add($label_User_Info)
     
     
      $InitialFormWindowState = $form_main.WindowState
      $form_main.add_Load($OnLoadForm_StateCorrection)
      $form_main.ShowDialog()| Out-Null
     
    }
     
    #Call the Function
    GenerateForm