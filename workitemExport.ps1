#Credit to James Palmer for initial request
Import-Module smlets
$ErrorActionPreference = "SilentlyContinue"

$form = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Export WorkItem to HTML" Height="372.6" Width="665" MinHeight="361" MinWidth="664" MaxHeight="373" MaxWidth="666">
    <TabControl HorizontalAlignment="Left" Height="340" Margin="0,0,0,0" VerticalAlignment="Top" Width="657">
        <TabItem Header="Single">
            <Grid Background="#FFE5E5E5">
                <ComboBox Name="cmbWIType" HorizontalAlignment="Left" Margin="125,28,0,0" VerticalAlignment="Top" Width="148" />
                <Label Content="WorkItem Type:" HorizontalAlignment="Left" Margin="19,25,0,0" VerticalAlignment="Top" Width="101"/>
                <ListBox Name="lstSource" HorizontalAlignment="Left" Height="100" Margin="19,87,0,0" VerticalAlignment="Top" Width="254"/>
                <ListBox Name="lstTarget" HorizontalAlignment="Left" Height="100" Margin="380,87,0,0" VerticalAlignment="Top" Width="254"/>
                <TextBox Name="txtWID" HorizontalAlignment="Left" Height="23" Margin="514,28,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
                <Label Content="WorkItem ID:" HorizontalAlignment="Left" Margin="402,25,0,0" VerticalAlignment="Top" Width="107"/>
                <Label Content="Available Properties:" HorizontalAlignment="Left" Margin="19,60,0,0" VerticalAlignment="Top" Width="125"/>
                <Label Content="Properties to Show:" HorizontalAlignment="Left" Margin="380,60,0,0" VerticalAlignment="Top" Width="113"/>
                <Button Name="btnAddAll" Content="Add All" HorizontalAlignment="Left" Margin="292,110,0,0" VerticalAlignment="Top" Width="75" Background="#FFD8833C"/>
                <Button Name="btnAdd" Content="Add Selected" HorizontalAlignment="Left" Margin="292,134,0,0" VerticalAlignment="Top" Width="75" Background="#FFD8833C"/>
                <Button Name="btnExport" Content="Export" HorizontalAlignment="Left" Margin="250,261,0,0" VerticalAlignment="Top" Width="153" Height="42" Background="#FFD8833C"/>
                <Button Name="btnRemove" Content="Remove Selected" HorizontalAlignment="Left" Margin="535,192,0,0" VerticalAlignment="Top" Width="99" Background="#FFD8833C"/>
                <Button Name="btnClear" Content="Clear" HorizontalAlignment="Left" Margin="431,192,0,0" VerticalAlignment="Top" Width="75" Background="#FFD8833C"/>
                <CheckBox Name="chkHistory" Content="Include History" HorizontalAlignment="Left" Margin="250,206,0,0" VerticalAlignment="Top" Width="153" IsChecked="True"/>
                <CheckBox Name="chkRelationships" Content="Get Relationships" HorizontalAlignment="Left" Margin="19,192,0,0" VerticalAlignment="Top" Width="151" IsChecked="True"/>
                <CheckBox Name="chkFiles" Content="Save Attachments" HorizontalAlignment="Left" Margin="250,224,0,0" VerticalAlignment="Top" IsChecked="True"/>
                <CheckBox Name="chkChildren" Content="Save Child Activities" HorizontalAlignment="Left" Margin="250,242,0,0" VerticalAlignment="Top" IsChecked="True"/>
                <CheckBox Name="chkRemote" Content="Remote Computer" HorizontalAlignment="Left" Margin="19,242,0,0" VerticalAlignment="Top"/>
                <TextBox Name="txtRemote" HorizontalAlignment="Left" Height="23" Margin="19,260,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Visibility="Hidden"/>
            </Grid>
        </TabItem>
        <TabItem Header="Multi">
            <Grid Background="#FFE5E5E5" Margin="0,-2,0.4,2.2">
                <DatePicker Name="dateFrom" HorizontalAlignment="Left" Margin="69,60,0,0" VerticalAlignment="Top" />
                <DatePicker Name="dateTo" HorizontalAlignment="Left" Margin="215,60,0,0" VerticalAlignment="Top"/>
                <TextBox Name="txtWIDFrom" HorizontalAlignment="Left" Height="23" Margin="69,133,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="102"/>
                <TextBox Name="txtWIDTo" HorizontalAlignment="Left" Height="23" Margin="215,133,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="102"/>
                <TextBox Name="txtCSVPath" HorizontalAlignment="Left" Height="23" Margin="87,206,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="162"/>
                <Button Name="btnCSVBrowse" Content="Browse" HorizontalAlignment="Left" Margin="258,207,0,0" VerticalAlignment="Top" Width="75" Background="#FFD8833C"/>
                <Label Content="From:" HorizontalAlignment="Left" Margin="21,58,0,0" VerticalAlignment="Top"/>
                <Label Content="To:" HorizontalAlignment="Left" Margin="180,58,0,0" VerticalAlignment="Top"/>
                <Label Content="From:" HorizontalAlignment="Left" Margin="21,133,0,0" VerticalAlignment="Top"/>
                <Label Content="To:" HorizontalAlignment="Left" Margin="180,133,0,0" VerticalAlignment="Top"/>
                <Label Content="CSV Path:" HorizontalAlignment="Left" Margin="21,207,0,0" VerticalAlignment="Top"/>
                <RadioButton Name="radDate" Content="Created Date Range:" HorizontalAlignment="Left" Margin="19,38,0,0" VerticalAlignment="Top"/>
                <RadioButton Name="radWID" Content="Work Item ID Range:" HorizontalAlignment="Left" Margin="21,113,0,0" VerticalAlignment="Top"/>
                <RadioButton Name="radCSV" Content="Read Work Items to export from CSV:" HorizontalAlignment="Left" Margin="19,186,0,0" VerticalAlignment="Top"/>
                <Label Content="Note: CSV should be in the format WorkItemID,WorkItemType on each line.  Example: IR123,Incident" HorizontalAlignment="Left" Margin="21,229,0,0" VerticalAlignment="Top"/>
                <Label Content="Note: Date and WorkItem Range will use WorkItem" HorizontalAlignment="Left" Margin="331,52,0,0" VerticalAlignment="Top" Width="296" />
                <Label Content="type from Single tab.  Current type:" HorizontalAlignment="Left" Margin="331,67,0,0" VerticalAlignment="Top" Width="204" />
                <Label Name="lblWIType" Content="" HorizontalAlignment="Left" Margin="520,67,0,0" VerticalAlignment="Top" Width="105" FontWeight="Bold" />
                <Button Name="btnExportMulti" Content="Export" HorizontalAlignment="Left" Margin="418,134,0,0" VerticalAlignment="Top" Width="103" Height="38" Background="#FFD8833C"/>
            </Grid>
        </TabItem>
    </TabControl>
</Window>
"@

Function Validate-Remote {
   ## Credit to Tom Hendricks for Remote Computer request
   if ($txtRemote.text -eq "") {
      $computername = $env:COMPUTERNAME
   }
   else {
      $computername = $txtRemote.Text
   }

   if ($chkRemote.isChecked) {
      if (Test-Connection -ComputerName $computername -Count 1 -ea SilentlyContinue) {
         Set-Variable -Name smdefaultcomputer -Value $computername -Scope Script
         return $true
      }
      else {
         [System.Windows.MessageBox]::Show("$($txtRemote.text) Couldn't be reached.  Falling Back to $env:COMPUTERNAME")
         Set-Variable -Name smdefaultcomputer -Value $env:COMPUTERNAME -Scope Script
         return $false
      }
      
   }
   else {
      Set-Variable -Name smdefaultcomputer -Value $env:COMPUTERNAME -Scope Script
      return $false
   }
   
}

Function Load-Dialog {
    Param(
        [Parameter(Mandatory=$True,Position=1)]
        [string]$XamlPath
    )
    [xml]$xmlWPF = $XamlPath
    try{
        Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
    } 
    catch {
        Throw "Failed to load Windows Presentation Framework assemblies."
    }
    $xamGUI = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xmlWPF))
    $xmlWPF.SelectNodes("//*[@Name]") | %{
        Set-Variable -Name ($_.Name) -Value $xamGUI.FindName($_.Name) -Scope Global
    }
    return $xamGUI
}

Function Process-Enum {
    param(
        [parameter(mandatory=$true)][string]$EnumToProcess
    )
    $Enum = Get-SCSMEnumeration -Name ($EnumToProcess + "$")
    return $Enum.DisplayName
}

Function Date-FromUTC {
   param(
      [parameter(mandatory=$true,position=1)][DateTime]$UTCDate
   )
   $LocalTZ = (Get-TimeZone).StandardName
   $Zone = [System.TimeZoneInfo]::FindSystemTimeZoneById($LocalTZ)
   $LocalDate = [system.timezoneinfo]::ConvertTimeFromUtc($UTCDate, $Zone)
   return $LocalDate
}

Function Process-Property {
    param(
        [parameter(mandatory=$true)][string]$Property
    )
    
    $ProcessDate = get-date
    if ($Property -like "*enum*") {
        $RO = Process-Enum -EnumToProcess $Property
    }
    elseif ([DateTime]::TryParse($Property, [ref]$ProcessDate)) {
       $RO = Date-FromUTC -UTCDate $ProcessDate
    }
    else {
    [guid]$GuidToProcess = '00000000-0000-0000-0000-000000000000'
       if ([guid]::TryParse($Property, [ref]$GuidToProcess)) {
          $SMObject = Get-SCSMObject -id $GuidToProcess -ea SilentlyContinue -ErrorVariable $WhatEror
    
          if (!($SMObject)) {
             $relObject = Get-SCSMRelationshipObject -id $GuidToProcess -ea SilentlyContinue
             if ($relObject) {
                $RO = $relObject.TargetObject
             }
          }
          else {
             if ($SMObject.ClassName -like "System.WorkItem.Trouble*") {
                if ($SMObject.Comment) {
                    $RO = $SMObject.Comment
                }
                else {
                    $RO = $SMObject.Title
                }
             }
             elseif ($SMObject.Classname -eq "system.reviewer") {
                
                $reviewer = Get-SCSMObject -Class (Get-SCSMClass -name "System.Reviewer") -Filter "id -eq $Property" 
                if ($reviewer) {
                    $reviwerRel = Get-SCSMRelationshipObject -BySource $reviewer
                    if ($reviwerRel) {
                        if ($reviewerRel.length -gt 1) {
                            $RO = $reviewerRel[1].TargetObject
                        }
                        else {
                            $RO = $reviwerRel.TargetObject
                        }
                    }
                }
                
                else {
                    $RO = "No Reviewer Assigned"
                }
             }
             else {
                $RO = $SMObject.displayname    
             }
          }
       }
    }
    if (!($RO)) { return $Property }
    else { return $RO }
}


Function Get-WorkItemFromType {
   param(
   [parameter(mandatory=$true)][string]$WorkItemID,
   [parameter(mandatory=$true)][string]$WorkItemType
   )
   
   switch ($WorkItemType) {
      "Incident" { $wiClass = Get-SCSMClass -name "System.Workitem.Incident$" }
      "Service Request" { $wiClass = Get-SCSMClass -name "system.workitem.serviceRequest$"  }
      "Change Request" { $wiClass = Get-SCSMClass -name "system.workitem.changeRequest$" }
      "Problem" { $wiClass = Get-SCSMClass -name "system.workitem.problem$" }
      "Activity" {
         $wiClass = Get-SCSMClass -name "System.workitem.activity$"
         $activity = $true
      }

   }
   if ($activity) {
      $workItem = Get-SCSMObject -Class $wiClass -Filter "name -eq $WorkItemID" -ea SilentlyContinue
   }
   else {
      $workItem = Get-SCSMObject -Class $wiClass -Filter "id -eq $WorkItemID" -ea SilentlyContinue
   }
   return $workItem
}

Function Get-WorkItemProperties {
   param(
      [parameter(mandatory=$true)][string]$WorkItemID,
      [parameter(mandatory=$true)][string]$WorkItemType,
      [parameter(mandatory=$false)][bool]$GetRelationships,
      [parameter(mandatory=$false)][switch]$NamesOnly,
      [parameter(mandatory=$false)][switch]$FromUI
   )
   $workItem = Get-WorkItemFromType -WorkItemID $WorkItemID -WorkItemType $WorkItemType
   if ($workItem) {
      $values = $workItem | Get-Member
      $PropertyNames = @()
      $PropWithValue = @{}

      foreach ($v in $values) {
         if ($v.membertype -eq 'NoteProperty' -and $v.name -ne 'Values') {
            $PropertyNames += ($v.name)
            $PropWithValue | Add-Member -MemberType NoteProperty -Name $v.Name -Value $workItem.($v.name)
         }
      }
      if ($GetRelationships) {
      $rels = Get-SCSMRelationshipObject -BySource $workItem | where {$_.targetobject.classname -notlike "System.WorkItem.TroubleTicket*"}
      $i = 0
         foreach ($r in $rels) {
            $rc = Get-SCSMRelationshipClass -id $r.relationshipid 
            $PropertyNames += ("Relationship:" + $rc.displayName + $i.ToString())
            $PropWithValue | Add-Member -MemberType NoteProperty -Name ("Relationship:" + $rc.displayName + $i.tostring()) -Value $r.targetobject.DisplayName
            if ($FromUI) {
               Set-Variable -Scope script -Name $($rc.displayname.replace(" ", "") + $i.tostring())  -Value $r.targetobject
            }
            
            else {
               Set-Variable -Scope script -Name $($WorkItem.Name + "." + $rc.displayname.replace(" ", "") + $i.tostring())  -Value $r.targetobject
            }
            
            $i++
         }
      } 
   }
   if ($NamesOnly) {
      return $PropertyNames
   }
   else {
      return $PropsWithValue
   }
}

Function Populate-Source {
   Validate-Remote
   $lstSource.items.Clear()
   $lstTarget.items.Clear()
   if ($txtWID.Text.Length -gt 3) {    
       $properties = Get-WorkItemProperties -WorkItemID $txtWID.Text -WorkItemType $cmbWIType.SelectedItem -GetRelationships $chkRelationships.IsChecked -NamesOnly -FromUI
       foreach ($p in $properties) {
          $lstSource.Items.Add($p)
       }
   }
   
}

Function Get-Attachment {
   param(
      [Guid] $Id,
      [string]$FileSavePAth
   )
   
   $WorkItem = Get-SCSMObject -Id $Id
   $WIhasAttachMentClass = Get-SCSMRelationshipClass -name System.WorkItemHasFileAttachment
   $WIClass = Get-SCSMClass System.WorkItem$   

   $files = Get-SCSMRelatedObject -SMObject $WorkItem -Relationship $WIhasAttachMentClass

   if($files -ne $Null) {
      if (!(test-path $FileSavePAth\Attachments)) {
         mkdir $FileSavePAth\Attachments
      }   
      foreach ($f in $files) {    
         $fs = [IO.File]::OpenWrite(($FileSavePath + "\Attachments\" + $f.DisplayName))
         $memoryStream = New-Object IO.MemoryStream
         $buffer = New-Object byte[] 8192
         [int]$bytesRead|Out-Null
         while (($bytesRead = $f.Content.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $memoryStream.Write($buffer, 0, $bytesRead)
         }        
         $memoryStream.WriteTo($fs)
       }

       $fs.Close()
       $memoryStream.Close()             
   }
}

Function Export-WorkItem {
   param(
      [parameter(mandatory=$true)][string]$WorkItemID,
      [parameter(mandatory=$false)]$Properties,
      [parameter(mandatory=$true)][string]$WorkItemType,
      [parameter(mandatory=$false)][bool]$GetRelationships,    
      [parameter(mandatory=$false)][string]$SavePath,
      [parameter(mandatory=$false)][switch]$FromUI,
      [parameter(mandatory=$false)][bool]$GetAttachments,
      [parameter(mandatory=$false)][bool]$GetHistory,
      [parameter(mandatory=$false)][bool]$ExportChildren
   )
   $WorkItem = Get-WorkItemFromType -WorkItemID $WorkItemID -WorkItemType $WorkItemType

   if ($workItem) {
      if (!($Properties)) {
         $props = Get-WorkItemProperties -WorkItemID $WorkItemID -WorkItemType $WorkItemType -GetRelationships $true -NamesOnly
      }
      else {
         $props = $Properties
      }
      $htmlOutput = "<html><head><title>$($workItem.id)</title><style>BODY{background-color:#FFFFFF;font-family:Verdana,sans-serif; font-size: small;}
      TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse; width: 98%;}
      TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#ff9333; font-weight: bold;text-align:left;}
      TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#ffc999; padding: 2px;}</style></head><body><h1>$($cmbWIType.SelectedItem) Properties</h1><table><tr><th><strong>PropertyName</strong></th><th><strong>PropertyValue</strong></th></tr>"
      if ($GetRelationships) {
         $wiWithProps = @{}
         $propTime = get-date
         foreach ($p in $props) {
            if ($workitem.($p) -like "*enum*") {
               $propToDisplay = Process-Enum -EnumToProcess ($workItem.($p))
            }
            elseif ([DateTime]::TryParse($workItem.($p), [ref]$propTime)) {
               $propToDisplay = Date-FromUTC -UTCDate $propTime

            }
            else {
               $propToDisplay = $workitem.($p)
            }
               
            $splitP = $p.split(":")
            if ($splitP[0] -eq "Relationship") {
               if ($FromUI) {
                  $relationshipTarget = Get-Variable -Name ($splitP[1].replace(" ", "")) -ValueOnly
               }
               else {
                  $relationshipTarget = Get-Variable -Name ($workItem.name + "." + $splitP[1].replace(" ", "")) -ValueOnly
               }
               $propToDisplay = $relationshipTarget.displayName
               $wiWithProps | Add-Member -Name ($splitP[1]) -Value ($propToDisplay) -MemberType NoteProperty
            }
            else {
               $wiWithProps | Add-Member -Name $p -Value $propToDisplay -MemberType NoteProperty
            }
         }
               
      }
      else{
         $wiWithProps = @{}
         foreach ($p in $props) {
            if ($workitem.($p) -like "*enum*") {
               $propToDisplay = Process-Enum -EnumToProcess ($workItem.($p))
            }
            else {
               $propToDisplay = $workitem.($p)
            }
            $wiWithProps | Add-Member -Name $p -Value $propToDisplay -MemberType NoteProperty
         }
      }

      $values = $wiWithProps | Get-Member
         
         
      foreach ($v in $values) {
         if ($v.MemberType -eq "NoteProperty") {
            $htmlOutput += "<tr><td><strong>$($v.Name):</strong></td><td>$($wiWithProps.($v.Name))</td></tr>"
         }
      }
      $htmlOutput += "</table>"
      ##Adding Reviwers table to RAs per request by Brad McKenna
      if ($WorkItem.className -eq 'System.WorkItem.Activity.ReviewActivity') {
          
         $rels = Get-SCSMRelationshipObject -Bysource $WorkItem | where {$_.relationshipId -eq '6e05d202-38a4-812e-34b8-b11642001a80'}
         $revTable = @()
         foreach ($r in $rels) {
            $revO = $r.TargetObject
            if ($revO.DisplayName.Length -gt 0) {
                $VotedBy = Get-SCSMRelationshipObject -BySource $revO | where {$_.relationshipid -eq '9441a6d1-1317-9520-de37-6c54512feeba'}
                $assignedReviewer = Get-SCSMRelationshipObject -BySource $revO | where {$_.relationshipid -eq '90da7d7c-948b-e16e-f39a-f6e3d1ffc921'}
        
                $values  = $revO.Values
                $revProps = @{}
                $revProps | Add-Member -MemberType NoteProperty -Name Reviewer -Value $assignedReviewer.TargetObject.DisplayName
                $revProps | Add-Member -MemberType NoteProperty -Name VotedBy -Value $VotedBy.TargetObject.DisplayName
                 
                foreach ($v in $values) {
                    switch ($v.Type) {
                        "Comments" {$revProps | Add-Member -MemberType NoteProperty -Name Comments -Value $v.value.Replace("<br/>","")}
                        "Decision" {$revProps | Add-Member -MemberType NoteProperty -Name Decision -Value $v.value.DisplayName}
                        "DecisionDate" {$revProps | Add-Member -MemberType NoteProperty -Name DecisionDate -Value (Date-FromUTC -UTCDate $v.value)}
                        "Veto" {$revProps | Add-Member -MemberType NoteProperty -Name Veto -Value $v.value}
                        "MustVote" {$revProps | Add-Member -MemberType NoteProperty -Name MustVote -Value $v.value}
                     }
                }
        
                $revTable += $revProps | select Reviewer,Decision,DecisionDate,VotedBy,Veto,MustVote,Comments
            }
         }
         $htmlOutput += "<h1>Reviwers</h1>"
         $htmlOutput += $revTable | ConvertTo-Html -Fragment
      }
      $htmlOutput += "<h1>Action Log</h1>"
      $ActionLog =  Get-SCSMRelationshipObject -BySource $workItem | ?{$_.targetobject.classname -like "System.workitem.troubleticket*"}
      $actionLogObject = @()
      foreach ($a in $ActionLog) {
         $thisTarget = Get-SCSMObject -id $a.TargetObject.id
         $actionLogObject += $thisTarget | select EnteredBy, Title, @{Label="EnteredDate"; Expression={Date-FromUTC $_.EnteredDate}}, Comment 
      }
      $htmlOutput += $actionLogObject | sort enteredDate -Descending | ConvertTo-Html -Fragment
      
      if ($getHistory) {
         $htmlOutput += "<h1>History</h1><table>"
         $history = Get-SCSMObjectHistory -Object $workItem 
         foreach ($h in $history.history) {
            $htmlOutput += ($h | select Username, LastModified) | ConvertTo-Html -Fragment
               
            $changes = @()
            foreach ($i in $h.Changes) {
               $newChangesOuput = @{}
               $newValue = Process-Property -Property $i.NewValue
               $oldValue = Process-Property -Property $i.OldValue

               $newChangesOuput | Add-Member -MemberType NoteProperty -Name Name -Value $i.Name 
               $newChangesOuput | Add-Member -MemberType NoteProperty -Name NewValue -Value $newValue 
               $newChangesOuput | Add-Member -MemberType NoteProperty -name OldValue -Value $oldValue
               $newChangesOuput | Add-Member -MemberType NoteProperty -Name WhatChanged -Value $i.WhatChanged
               $newchangesOuput | Add-Member -MemberType NoteProperty -Name TypeOfChange -Value $i.TpyeOfChange
        
               $changes += $newChangesOuput | Select Name,WhatChanged,OldValue,NewValue
            }
            $htmlOutput += $changes | ConvertTo-Html -Fragment
            $htmlOutput += "<br>"
         }

         $htmlOutput += "<br>"
      }
         
      $htmlOutput += "</body></html>"
      if (!($SavePath)) {
         $savePath = "$env:USERPROFILE\Desktop\$($workitem.id)"
      }
      if (!(Test-Path $savePath)) {
         mkdir $savePath
      }
      $htmlOutput | Out-File $savePath\$($workItem.id).html -Force
      
      if ($GetAttachments) {
         Get-Attachment -Id $WorkItem.get_id() -FileSavePAth $SavePath
      }

      if ($ExportChildren) {
         $kids = Get-Children -WorkItem $WorkItem
         foreach ($k in $kids) {
            Export-WorkItem -WorkItemID $k.name -WorkItemType "Activity" -GetRelationships $true -GetHistory $true -SavePath $SavePath -GetAttachments $true
         }
      }
      
   }   

 }

Function Get-Children {
   param(
       $WorkItem
   ) 
   $RelClass = Get-SCSMRelationshipClass -name System.WorkItemContainsActivity 
   $Children = Get-SCSMRelatedObject -SMObject $WorkItem -Relationship $RelClass
   return $Children
}

Function Toggle-Visibility {
   param(
      [parameter(mandatory=$true)]$Control    
   )
   if ($Control.visibility -eq "Hidden"){
      $Control.Visibility = "Visible"
   }
   else {
      $Control.Visibility = "Hidden"
   }
}

Function Get-Prefix {
   param(
      [parameter(mandatory=$true)][string]$WorkItemType
   )
   switch ($WorkItemType) {
      "Incident" {
         $incSettings = Get-SCSMSetting -id "613c9f3e-9b94-1fef-4088-16c33bfd0be1"
         $prefix = $incSettings.PrefixForId
      }
      "Service Request" {
         $srSettings = Get-SCSMSetting -id "fa662352-1660-33ae-6316-7fe1c9fecc6d"
         $prefix = $srSettings.ServiceRequestPrefix
      }
      "Change Request" {
         $crSettings = Get-SCSMSetting -id "c7fe33bb-9760-3f88-59fc-0951e3221be4"
         $prefix = $crSettings.SystemWorkItemChangeRequestIdPrefix
      }
      "Problem" {
         $prSettings = Get-SCSMSetting -id "da0eeac9-9c85-e72b-f321-44a3fcec9c9a"
         $prefix = $prSettings.ProblemIdPrefix
      }
      "Activity" {
         $listOfPrefixes = @()
         $activitySettings = Get-SCSMSetting -id 5e04a50d-01d1-6fce-7946-15580aa8681d
         $listOfPrefixes += $activitySettings.MicrosoftSystemCenterOrchestratorRunbookAutomationActivityBaseIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivityDependentActivityIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivityIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivityManualActivityIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivityParallelActivityIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivityReviewActivityIdPrefix
         $listOfPrefixes += $activitySettings.SystemWorkItemActivitySequentialActivityIdPrefix
         return $listOfPrefixes
      }
   }
   return $prefix

}

Function GetWorkItems-FromRange {
   param(
      [parameter(mandatory=$true)][string]$prefix,
      [parameter(mandatory=$true)][int]$start,
      [parameter(mandatory=$true)][int]$end
   )
   $workItemClass = Get-SCSMClass -name system.workitem$
   $workItemObject = @()
   for ($i = $start; $i -le $end; $i++) {
      $thisWorkItem = Get-SCSMObject -Class $workItemClass -Filter "Name -eq $($prefix + $i.tostring())"
      $workItemObject += $thisWorkItem
   }
   return $workItemObject
}

Function Get-FileName($initialDirectory) {
    ## This is not a function I wrote.  I cleaned up its formatting to bring it inline with the rest of the code   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
} 

$win = Load-Dialog $Form
$cmbWIType.items.add("Incident") | out-null
$cmbWIType.items.add("Service Request") | Out-Null
$cmbWIType.items.add("Change Request") | Out-Null
$cmbWIType.items.add("Problem") | Out-Null
$cmbWIType.items.add("Activity") | Out-Null

### Single Tab ###

$txtWID.add_textChanged({
   Populate-Source
})

$cmbWIType.add_selectionChanged({
   Populate-Source
   $lblWIType.Content = $cmbWIType.SelectedItem
})

$btnAdd.add_click({
   $lstTarget.items.add($lstSource.SelectedItem)
   $lstSource.items.Remove($lstSource.SelectedItem)
})

$btnAddAll.add_click({
   foreach ($i in $lstSource.items) {
      $lstTarget.items.add($i)
   }
   $lstSource.items.Clear()
})

$btnClear.add_click({
   $lstTarget.Items.clear()
   Populate-Source
})

$btnRemove.add_click({
   $lstSource.items.add($lstTarget.SelectedItem)
   $lstTarget.items.Remove($lstTarget.SelectedItem)
})

$btnExport.add_click({
  if ($lstTarget.HasItems) {
     Export-WorkItem -WorkItemID $txtWID.text -Properties $lstTarget.Items -WorkItemType $cmbWIType.SelectedItem -GetRelationships $chkRelationships.IsChecked -FromUI -GetAttachments $chkFiles.IsChecked -ExportChildren $chkChildren.IsChecked -GetHistory $chkHistory.IsChecked
     Invoke-Item "$env:USERPROFILE\desktop\$($txtWID.text)\$($txtWID.text).html"
  }
  else {
     [System.Windows.MessageBox]::Show("No properties selected to export")
  }
})

$chkRelationships.add_click({
    Populate-Source
})

$chkRemote.add_click({
   Toggle-Visibility -Control $txtRemote
   Populate-Source
})

$txtRemote.add_LostFocus({
   Populate-Source
})

### Multi Tab ###
## Credit to Jonathan Boles for "range" request
$dateTo.add_GotKeyboardFocus({ 
   $radDate.IsChecked = $true
})
$dateFrom.add_GotKeyboardFocus({
   $radDate.IsChecked = $true
})
$txtWIDFrom.add_GotKeyboardFocus({
   $radWID.IsChecked = $true
})
$txtWIDTo.add_GotKeyboardFocus({
   $radWID.IsChecked = $true
})
$txtCSVPath.add_GotKeyboardFocus({
   $radCSV.IsChecked = $true
})

$btnCSVBrowse.add_click({
   $radCSV.IsChecked = $true
   $txtCSVPath.Text = Get-FileName -initialDirectory $env:USERPROFILE\Desktop 
})

$btnExportMulti.add_click({
   if ($radDate.IsChecked) {
      if ($cmbWIType.SelectedIndex -ge 0) {
         if ((get-date $dateTo.SelectedDate) -gt (get-date $dateFrom.SelectedDate)) {
            $dateWIs = Get-SCSMObject -Class (Get-SCSMClass -name system.workitem$) | `
            ?{((get-date $_.CreatedDate) -gt (get-date $dateFrom.SelectedDate)) -and ((get-date $_.CreatedDate) -lt (get-date $dateTo.SelectedDate))} | select id
            if ($dateWIs) {
               foreach ($wi in $dateWIs) {
                  Export-WorkItem -WorkItemID $wi.Id -WorkItemType $cmbWIType.SelectedItem -GetRelationships $true -GetAttachments $true -ExportChildren $true -GetHistory $true -SavePath $env:USERPROFILE\Desktop\MultiExport\$($wi.id)
               }
            }
            else {
               [System.Windows.MessageBox]::Show("No WorkItems in the selected date range")
            }
         }
         else {
            [System.Windows.MessageBox]::show("From Date must be before To Date")
         }
      }
      else {
         [System.Windows.MessageBox]::Show("No Work Item type selected on Single tab")
      }
   }
   elseif ($radWID.IsChecked) {
      if ($cmbWIType.SelectedIndex -ge 0) {
         [int]$rangeStart = $txtWIDFrom.Text
         [int]$rangeEnd = $txtWIDTo.text
         if ($rangeStart -lt $rangeEnd) {
            $WorkItemPrefix = Get-Prefix -WorkItemType $cmbWIType.SelectedItem
            if ($cmbWIType.SelectedItem -eq "Activity") {
               $activityList = @()
               foreach ($p in $WorkItemPrefix) {
                  $workItemsFromRange = GetWorkItems-FromRange -prefix $p -start $rangeStart -end $rangeEnd
                  $activityList += $workItemsFromRange
               }
               if ($activityList.length -gt 0) {
                  foreach ($a in $activityList) {
                     Export-WorkItem -WorkItemID $a.Name -WorkItemType $cmbWIType.SelectedItem -GetRelationships $true -GetAttachments $true -ExportChildren $true -GetHistory $true -SavePath C:\Users\scsm_service\Desktop\MultiExport\$($a.Name)
                  }
               }
               else {
                  [System.Windows.MessageBox]::show("No Work Items in specified range")
               }
            }
            else {
               $workItemsFromRange = GetWorkItems-FromRange -prefix $WorkItemPrefix -start $rangeStart -end $rangeEnd
               if ($workItemsFromRange) {
                  foreach ($w in $workItemsFromRange) {
                     Export-WorkItem -WorkItemID $w.id -WorkItemType $cmbWIType.SelectedItem -GetRelationships $true -GetAttachments $true -ExportChildren $true -GetHistory $true -SavePath $env:USERPROFILE\Desktop\MultiExport\$($w.id)

                  }
               }
               else {
                  [System.Windows.MessageBox]::Show("No Work Items in specified range")
               }
            }
         }
         else {
            [System.Windows.MessageBox]::Show("Start Value must be less than end Value")
         }
      }
      else {
         [System.Windows.MessageBox]::Show("No Work Item type selected on Single tab")
      }
   }
   elseif ($radCSV.IsChecked) {
      if ($txtCSVPath.Text -ne "") {
         if (Test-Path $txtCSVPath.Text) {
            $file = Get-ChildItem $txtCSVPath.Text
            if ($file.Extension -eq ".csv") {
               $csv = Get-Content $txtCSVPath.Text
               foreach ($line in $csv) {
                  $l = $line.split(",")
                  $workItemID = $l[0]
                  $workItemType = $l[1]
                  Export-WorkItem -WorkItemID $workItemID -WorkItemType $workItemType -GetRelationships $true -GetAttachments $true -ExportChildren $true -GetHistory $true -SavePath $env:USERPROFILE\Desktop\MultiExport\$($workItemId)
               }
            }
            else {
               [System.Windows.MessageBox]::Show("Specified file is not a CSV")
            }
         }
         else {
            [System.Windows.MessageBox]::Show("CSV Doesn't appear to be valid")
         }
      }
      else {
         [System.Windows.MessageBox]::Show("No CSV selected")
      }
   }
   else {
      [System.Windows.MessageBox]::Show("No Multi-Export option selected")
   }
})

$win.showdialog() | Out-Null