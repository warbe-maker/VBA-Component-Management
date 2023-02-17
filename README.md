## Table of contents
<sup>[Management of Excel VB-Project Components](#management-of-excel-vb-project-components)  
[Disambiguation](#disambiguation)  
[Services](#services)  
&nbsp;&nbsp;&nbsp;[The _ExportChangedComponents_ service](#the-exportchangedcomponents-service)  
&nbsp;&nbsp;&nbsp;&nbsp;[The _UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service)  
&nbsp;&nbsp;&nbsp;&nbsp;[The Synchronize VB-Project service](#the-synchronize-vb-project-service)  
[Installation](#installation)  
[Configuration](#configuration)  
[Usage](#usage)  
[Other](#other)  
</sup>

## Management of Excel VB-Project Components
- **Export** any _Component_ the code has changed (automated with the Workbook_BeforeSave_ event)
- **Update** outdated _Used&nbsp;[Common Components][7]_
- **Synchronize** the VB-Project of two Workbooks
- **Manage** _Hosted Common&nbsp;Components_
 
Also see the [Programmatic-ally updating Excel VBA code][2] post for this subject.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
| _Component_       | Generic term for any kind of _VB-Project-Component_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_  |
| _Common&nbsp;Component_ | A _VB-Component_ hosted in a Workbook which claims it as the _Raw_ component. This Workbook is dedicated for the development, maintenance and tested of this VB-Component while other Workbooks/VB-Projects are using it as _Clone_ [^1].   |
|_Used&nbsp;Common&nbsp;Component_ | The copy of a _Raw&#8209;Component_ in a _Workbook/VP&#8209;Project_ using it. _Clone-Components_ may be automatically kept up-to-date by the _UpdateOutdatedCommonComponents_ service.  The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_ |
| _Procedure_           | Generic term for any _VB-Component's_ (public or private) _Property_, _Sub_, or _Function_ |
| _Raw&nbsp;Common&nbsp;Component_ | The instance of a _Common&nbsp;Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
| _Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
| _Servicing&nbsp;Workbook_ | The service providing Workbook, either the _[CompMan.xlsb][1]_ Workbook (when it is open) or the _CompMan Add-in_ when it is set up and open. |
| _Serviced&nbsp;Workbook_ | The Workbook prepared for being serviced, provided it is located within the _Serviced&nbsp;Folder_ for the Update and Export service or in the .
|_VB&#8209;Project_    | Used synonymous with Workbook |
| _Sync&#8209;Source&#8209;Workbook_   | A _Workbook/VP-Project_ temporarily copied to the [Serviced _Development and Test Folder_](#serviced-development-and-test-folder) for being modified - and finally synchronized with its origin Workbook.|
| _Sync&#8209;Target&#8209;Workbook_ | A _Workbook/VP-Project_ temporarily moved to a [_Serviced _Synchronization Target Folder_](#serviced-synchronization-target-folder) for being synchronized with this Workbook which had been temporarily copied to the [Serviced _Development and Test Folder_](#serviced-development-and-test-folder) for being modified. |
| _Workbook&#8209;Folder_ | A folder dedicated to a _Workbook/VB-Project_ with all its Export-Files (in a \source sub-folder). When the folder is the equivalent of a _GitHub repo_ it may contain other files like a README and a LICENSE (provided GitHub is used for the project's versioning which not only  recommendable but also pretty easy to use.|

[^1]: The terms _Raw_ and _Clone_ follow [_GitHub_][6] terminology

## Services
### The _ExportChangedComponents_ service
Used with the _Workbook\_BeforeSave_ event, all _VB-Components_ code is compared with its previous _Export&nbsp;File_. When the code had changed the component is re-exported to the configured _Export Folder_ of which the name defaults to _source_. These _Export&nbsp;Files_ not only function as a code  backup in case Excel ends up with a destroyed VB-Project, which may happen every now and then - but only functions as a versioning means (e.g. when [GitHub][6] is used for instance). See also [Common Components](#common-components) which are handled specifically.

### The _UpdateOutdatedCommonComponents_ service
Used with the _Workbook\_Open_ event all  _Used&nbsp;Common&nbsp;Components_ are checked whether they are outdated. In case a dialog is displayed which allows to display the code difference (by means of WinMerge) perform the update or skip it. The update uses the  _Export&nbsp;File_ of the _Raw&nbsp;Common&nbsp;Component_ in the _Common&nbspComponents_ folder. This service is the core service and most critical service provided by CompMan. Excel may every now and then close the serviced Workbook when code is updated. Fortunately the Workbook can be opened again and the update service continues.  

### The _Synchronize VB-Project_ service
The service allows a productive Workbook to remain in use while the VB-Project is developed, modified, maintained, etc., in a copy of it. When all changes had been done the VB-Project of the productive Workbook is synchronized. The benefit: A significant shorter downtime for the productive Workbook. 
> As with the _Export_ and _Update_ service, this service has no user interface! The service is invoked when the _Sync-Target-Workbook_ (i.e. the temporarily moved productive Workbook) is opened from within the configured _Synchronization Target Folder_. Provided the _Sync-Source-Workbook_ (i.e the copy of the productive Workbook) resides in the [_Serviced Development and Test Folder_](#configuration).

## Installation
1. Download and open [CompMan.xlsb][1] (GitHub users may fork the corresponding [public repo][8] and open the CompMan.xlsb Workbook)
2. When the Workbook is opened it will show a _Config_ Worksheet for all required [Configuration](#configuration)
3. [Enable](#enabling-the-services) any service in any Workbook and have it [serviced by CompMan](#serviced-or-not-serviced)

## Configuration
When the [CompMan.xlsb][1] Workbook is opened a _Config_ Worksheet is displayed providing all means for configuring or changing the configuration.
#### The _Serviced Development and Test Folder_
The folder is essential for CompMan's  _Export Changed Components_ and or _Update Outdated Components_ service because the service is only provided for Workbooks with [enabled](#enabling-the-services) services when opened from within a sub-folder of this folder. When no such folder is configured or the configured one is invalid the services will be denied without notice. When the configuration of the folder is terminated, i.e. no folder is selected, the folder becomes 'not configured'. 
#### The Export Folder
Folder within a Workbook's folder into which the _Export_ service exports modified _VB-Components_. This folder is one of the many reasons why a serviced Workbook must be the only Workbook in its parent folder. The name of the _Export Folder_ defaults to _source_ but may be changed at any time. Export folders with an outdated name will be renamed right away or when detected by the _Export Service_. 
#### The Add-in Folder
The folder in which the CompMan Workbook is saved to as Add-in. The folder is obligatory only when CompMan is about to be setup as _Add-in_. When no _Add-in Folder_ is configured the [CompMan.xlsb][1] Workbook cannot be setup as Add-in. When the configuration of the folder is terminated, i.e. no folder is selected, the folder becomes 'not configured'.
#### The Serviced Synchronization Target Folder
Folder into which a Workbook for which the _Synchronize VB-Project_ has been [enabled](#enabling-the-synchronization-service), is temporarily moved and opened from there in order to have its VB-Project synchronized with the corresponding (same named) Workbook residing in the configured [_Serviced Development and Test Folder_](#serviced-development-and-test-folder). When the configuration of the folder is terminated, i.e. no folder is selected, the folder becomes 'not configured'.
#### The Synchronization Archive Folder
A _Synchronization Archive Folder_ is obligatory when the _Synchronize VB-Project_ service is used. The service will archive a _Sync-Target-Workbook_ before it is synchronized with its corresponding _Sync-Source-Workbook_. When none is selected the folder will be reset to 'not configured'. When the configuration of the folder is terminated, i.e. no folder is selected, the folder becomes 'not configured'. 
#### Setup Auto-open for the CompMan Workbook
Once Workbooks with enabled services are frequently opened it may make sense to have the the [CompMan.xlsb][1] Workbook automatically opened when Excel starts. This configuration switches from 'not setup' to 'setup' and vice versa. Note: Also the [CompMan Add-in](#compman-used-as-add-in) will be setup with an Auto-open. And just in case: When the CompMan Workbook is moved to a different location and opened, the setup Auto-open (in the users XLSTART folder) will automatically be updated.
#### Setup/Renew Add-in
Sets up or renews the CompMan Add-in. Once the Add-in is setup/renewed it will automatically be opened when Excel starts.
>The command button is only available/visible when a valid [_Add-in Folder_](#the-add-in-folder) is configured.
#### Pause/Continue Add-in
Pauses/continues the setup Add-in. This configuration option is only used when the [CompMan.xlsb][1] Workbook is maintained in order to enforce its use even when the Add-in is open/available.
>The command button is only available/visible when a valid [_Add-in Folder_](#the-add-in-folder) is configured.
## Usage
### Serviced or not serviced
A Workbook will only be serviced by CompMan provided
- the opened Workbook has a service [enabled](#services-enabling)
- a ***servicing CompMan instance*** (the [CompMan.xlsb][1] Workbook and/or the [CompMan Add-in](#compman-used-as-add-in) is open
- a valid [_Serviced Development and Test Folder_](#serviced-development-and-test-folder) is configured
- the ***to-be-serviced Workbook*** is opened from within a sub-folder of the configured [_Serviced Development and Test Folder_](#configuration), in case of the _Synchronization service_ from within a sub-folder of the configured _Sync-Target-Folder_.
- the ***to-be-serviced Workbook*** is the only Workbook in its parent folder (the parent folder may have sub-folders with Workbooks however)
- WinMerge ([WinMerge English][4], [WinMerge German][5] or any other language version is installed (it is used to display the difference for any components about to be updated by the [_Update_](#enabling-the-update-service)

As a consequence from the above, a productive Workbook must not be used from within the configured [_Serviced Development and Test Folder_](#configuration). When a Workbook with any [enabled](#services-enabling)/prepared service is opened when located elsewhere the user will not be bothered by any means, i.e. will not even recognize CompMan at all - even when open/available.

### Enabling the services
>Even when a Workbook has one or more services enabled, the service is denied without notice when [required (pre)conditions](#serviced-or-not-serviced) are not met.
#### Enabling the _Export_ service
The _Export_ service is performed whenever the Workbook is saved from within the configured [_Serviced Development and Test Folder_](#configuration) and all the [preconditions](#serviced-or-not-serviced) are met..
1. Download and import the _[mCompManClient.bas][3]_ which serves as the link to the CompMan services.
2. Into the Workbook module copy the following:
```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan) had been
    '~~ opened from within a sub-folder of the configured 'Serviced Development and Test Folder'.
    '~~ HOSTED_RAWS is a vbNullString constant when the Workbook not hosts any Common Component. 
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub
```
3. In case the Workbook hosts one or more _Common Components_ copy into the Workbook module:<br>
`Private Const HOSTED_RAWS = <component-name>[,<component-name]...`<br>
else<br>
`Private Const HOSTED_RAWS = vbNullstring`

#### Enabling the _Update_ service
The _Update_ service is performed whenever the Workbook is opened from within the configured [_Serviced Development and Test Folder_](#configuration) folder and all the [preconditions](#serviced-or-not-serviced) are meet.
1. Download and import the _[mCompManClient.bas][3]_ which serves as the link to the CompMan services.
2. Into the Workbook module copy the following:
```vb
Private Sub Workbook_Open()
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan) had been
    '~~ opened from within a sub-folder of the configured 'Serviced Development and Test Folder'.
    '~~ HOSTED_RAWS is a vbNullString constant when the Workbook not hosts any Common Component.
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub
```
3. In case the Workbook hosts one or more _Common Components_ copy into the Workbook module:<br>
`Private Const HOSTED_RAWS = <component-name>[,<component-name]...`<br>
else<br>
`Private Const HOSTED_RAWS = vbNullstring`


#### Enabling the _Synchronization_ service
The _Synchronize VB-Project_ service is performed when the Workbook is opened from within the configured [_Serviced Synchronization Target Folder_](#configuration) and all the [preconditions](#serviced-or-not-serviced) are meet
1. Download and import the _[mCompManClient.bas][3]_ which serves as the link to the CompMan services.
2. In the Workbook module copy the following:
```vb
Private Sub Workbook_Open()
    '~~ The statement is only required for a Workbook which may be synchronized.
    '~~ However, the statement has no effect unless the Workbook is opened from within the configured 'Serviced Synchronization Target Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_SYNCHRONIZE
End Sub
```

### Using the _Synchronization Service_
#### Steps
1. Follow the [Installation](#installation) instructions and [Enabling the _Synchronization_ service](#enabling-the-synchronization-service). 
2. Copy the productive Workbook into a dedicated folder within the configured [_Serviced Development and Test Folder_](#configuration) and modify the VB-Project as intended while the productive Workbook remains in use!.
3. When the development/modification has been finished, ****close the Workbook!**** and proceed with the next steps.
4. Move the productive Workbook to the configured [_Serviced Synchronization Target Folder_](#configuration) and open it. In case this folder has yet not been configured, switch to the open [CompMan.xlsb][1] and use the _CompMan Configuration_ button on the displayed Worksheet. Provided the Workbook had been prepared for it the synchronization will now start.
4. Open the moved Workbook and follow the synchronization steps. The synchronization will be done on a working copy (name with a suffix _Synced_).
6. When the synchronization has finished, save the working copy as the new productive Workbook - e.g. by dropping the _Synced_ suffix from the name and moving it back to the "production location" it originates.
7. When everything has finally turned out perfect the remaining Workbook from step 3 may be removed

#### _Worksheet_ synchronization
- ***New Worksheets***
  - The Sync-Source-Worksheet is cloned to the Sync-Target-Workbook
  - Back-links to the Sync-Source-Workbook are eliminated
  - All concerned Names scope is synchronized
- ***Obsolete Worksheets*** are removed
- ***Worksheets' Name <span style="color:red">****or!****</span> Code-Name change*** are synchronized
- Not yet implemented: ***Worksheets owned by the VB-Project***, that means those protected and without any unlocked (input) cell, are synchronized by default - disregarding any change.

>Attention! <span style="color:red">The _Name_ ***and*** the _CodeName_ of a Worksheet must never both be changed.</span> When a Worksheet's _Name_ ***and*** its _CodeName_ is changed at the same time the concerned sheet will be considered new and the (no longer identifiable as such) corresponding sheet will be considered obsolete - which in such a case is definitely not what was intended.
#### _References_ synchronization
New References are added and obsolete References are removed.
#### _VB-Components_ synchronization
Synchronized are all types of VB-Components:  (_Standard&nbsp;Module_, _Data&nbsp;Module_, _Class&nbsp;Module_, _UserForm_). New components are added, obsolete components are removed, and of changed components the code is updated.

#### _Names_ synchronization
- In order to provide a full transparent synchronization of Names they are synchronized before the Worksheets are [^2] - which makes it a bit more complex. When a Worksheet's name is about to change the Names synchronization has to deal with an old-named Worksheet while the source Name refers already the new Worksheet name. 
- When the referred range of a Name has changed this is not synchronize but either skipped or the synchronization process is interrupted in order to synchronize the change in the source Worksheet's layout in the corresponding target Worksheet. In case the 'change' is caused by the fact that a new row/column has been inserted by the Workbook user while the VB-Project has been maintained this issue can just be ignored, i.e. skipped.
- For a best possible transparent process, multiply named ranges are handled separated by removing all target- and adding all corresponding source-Names.
- ***Synchronization of obsolete Names:*** A Name is regarded obsolete when it only exists in the Sync-Target-Workbook but not (no longer) in the Sync-Source-Workbook
- ***Synchronization of new Names:*** A Name is considered new when it (the Name's 'mere name' [^3]) exists in the Sync-Source-Workbook but not in the Sync-Target-Workbook.
- ***Synchronization of changed Names and/or Scopes:*** A ranges Name and/or its Scope may bee changed in the Sync-Source-Workbook and will accordingly synchronized in the Sync-Target-Workbook. 
- ***Manual pre-synchronization preparation:*** When a synchronization is intentionally terminated (interrupted respectively) this will only be done in order to manually synchronize a design-change. To support this, the Sync-Source-Workbook and the Sync-Target-Workbook's working copy will be closed and the Sync-Source-Workbook re-opened. In the open dialog "manual pre-synchronization" will be chosen and once done the Workbook closed and re-opened with the option "continue with the ongoing synchronization".  

[^2]: When during the Worksheet synchronization a new sheet is cloned all Names are cloned too which obstructs a transparent Names synchronization.
[^3]: A Name objects 'mere name' is one without a sheet-name-prefix

##### Non synchronized 'ambigous names'

#### _Sheet-Shape_ synchronization
Still under construction!
New Shapes (including ActiveX-Controls) are added, obsolete Shapes are removed. The Properties of all Shapes are synchronized. However, though largely covered the properties synchronization may still be incomplete. 

### Other
#### Status of the CompMan Add-in
| Status             | Meaning |
|--------------------|---------|
| configured         | A valid, existing [_Add-in folder_](#add-in-folder) is specified. CompMan's Add-in instance  may be setup/renewed  |
| not configured     | No [_Add-in folder_](#add-in-folder) is specified. CompMan's Add-in instance cannot be setup/renewed (the button is not visible) |
| setup              | CompMan's Add-in instance is setup, i.e. available in the configured [_Add-in folder_](#add-in-folder) |
| not setup          | CompMan's Add-in instance is not setup, i.e. not available in the configured [_Add-in folder_](#add-in-folder). |
| paused             | CompMan's Add-in instance is indicated 'currently paused'. I.e. even when open it will  programmatic-ally be ignored |
| open               | CompMan's Add-in instance is open (it may be paused however) |
| not open           | CompMan's Add-in instance is not open. It will be opened when setup/renew and when Excel is started because the setup/renew establishes auto-open. |
| auto-open setup    | With setup/renew the auto-open has implicitly setup/established. |
| auto&#8209;open&nbsp;not&nbsp;setup| When the [_Add-in folder_](#add-in-folder) is de-configured (no folder is selected when with 'Configure') a setup auto-open is removed |
 
#### Common Components
One of the initial intentions for the development of CompMan was to keep _Common&nbspComponent_ up-to-date in all VB-Projects which use them. Thus the export service handles _Raw&nbsp;Common&nbsp;Components_ in a specific way: It registers hosted _Raw&nbsp;Common&nbsp;Components_, it increases a [_Revision Number_](#the-revision-number) with each export and additionally copies the _Export&nbsp;File_ to a _Common Components_ folder which functions as the source for the [_UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service) (while the hosting Workbook itself is not in charge with this service.  
The service also checks whether a  _Used&nbsp;Common&nbsp;Component_ has been modified within the VB-Project using it - which may happen accidentally or intentionally - and registers a **due modification revert alert** displayed when the Workbook is opened subsequently and the [_UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service) is about to revert the made modifications, allowing to display the code difference (using WinMerge).

#### The _Common Components_ folder
_CompMan_ maintains for each _Raw&nbsp;Common&nbsp;Component_ a copy of the _Export File_ in a _Common&nbsp;Components_ folder. These _Export Files_ are the source for a serviced Workbook's outdated _Used&nbsp;Common&nbsp;Components_. When a _Hosted Raw&nbsp:Common&nbsp;Component_ is modified it is not only exported like any other component but also copied to the _Common&nbsp;Components_ folder thereby increasing a _Revision Number_. 

#### The _Revision Number_
CompMan is pretty much focused on _Common&nbsp;Components_. In order to prevent updates of _Used&nbsp;Common&nbsp;Components_ with outdated raw versions CompMan maintains a _Revision Number_ for them which is increased whenever a new modified version is exported. The _Revision Number_ is maintained in a file _ComCompsHosted.dat_ located in the Workbook folder and kept in sync with the Revision Number_ in a file _ComCompsSaved.dat_ located in [the _Common Components_ folder](#the-common-components-folder).

#### CompMan's specific files

| File                     | Location             | Description               |
|--------------------------|----------------------|---------------------------|
| _ComCompsHosted.dat_     | The serviced Workbook's parent folder | PrivateProfile file for the registration of all _Raw&nbsp;Common&nbsp;Components_ hosted in the corresponding Workbook. Content scheme:<small>  `[component-name]`  `RawExpFileFullName=<file-full-name>`  ` RawRevisionNumber=YYYY-MM-DD.000>` |
| _ComCompsUsed.dat_       | The serviced Workbook's parent folder | Private Profile file for the registration of all _Used&nbsp;Common&nbsp;Components_. Content scheme:<small>  `[component-name]`  `RawRevisionNumber=YYYY-MM-DD.000>` |
| _ComComps&#8209;RawsSaved.dat_ | [_Common&nbsp;Components_ folder](#common-components-folder) | PrivateProfile file for the registration of all known _Raw&nbsp;Common&nbsp;Components_ |
| _CompMan.Service.trc_ | The serviced Workbook's parent folder | Execution trace of the performed CompMan service, available only when the _Conditional Compile Argument_ `ExecTrace = 1` is set in the servicing Workbook which is either the CompMan.xlsb Workbook directly or the CompMan.xlam Add-in instance of it. |
|  _CompMan.Service.log | The serviced Workbook's parent folder | Log file for the executed CompMan services. |

#### Multiple computers involved in VB-Project's development/maintenance
I do use two computers at two locations and prefer not to be bound to one. Some may prefer a network drive others a cloud based service. I prefer GitHub which makes using several computers very comfortable. In any case there is a certain need to prevent updates of used _Common&nbsp;Components_ with outdated hosted/raw versions.

#### CompMan used as Add-in
Life is easy when the [CompMan.xls][1] Workbook is open and that's why it is possible to configure an Auto-open for it. Alternatively, CompMan may be setup as Add-in. This setup will also configure an Auto-open for it. Workbooks with [enabled services](#enabling-the-services) will be served by either of the two depending on which one is open.

##### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open. Pausing the Add-in is only a CompMan development feature. When the Add-in is paused while the [CompMan.xlsb][1] is open CompMan works as if the Add-in were not setup which means the services are directly provided by the open [CompMan.xlsb][1]. When the [CompMan.xlsb][1] Workbook is closed and an Add-in had been setup the Add-in will be _continued_ automatically. This ensures that the Add-in is available for the [CompMan.xlsb][1] Workbook when it is opened again.
> The _CompMan Add-in_ is the only means which allows to update outdated _Used&nbsp;Common&nbsp;Components_ in the [CompMan.xlsb][1] Workbook. I.e. the development instance of the Add-in.

## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatic-ally-updating-Excel-VBA-code][2].


[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/2021/02/06/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
[4]:https://winmerge.org/downloads/?lang=en
[5]:https://winmerge.org/downloads/?lang=de
[6]:https://github.com
[7]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[8]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services