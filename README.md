## Table of contents
<sup>[Management of Excel VB-Project Components](#management-of-excel-vb-project-components)  
[Disambiguation](#disambiguation)  
[Services](#services)  
&nbsp;&nbsp;&nbsp;[The _ExportChangedComponents_ service](#the-exportchangedcomponents-service)  
&nbsp;&nbsp;&nbsp;&nbsp;[The _UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service)  
&nbsp;&nbsp;&nbsp;&nbsp;[The Synchronize VB-Project service](#the-synchronize-vb-project-service)  
[Installation](#installation)  
[Configuration](#configuration)  
&nbsp;&nbsp;&nbsp;[Serviced Development and Test Folder](#serviced-development-and-test-folder)  
&nbsp;&nbsp;&nbsp;[Add-in Folder](#add-in-folder)  
&nbsp;&nbsp;&nbsp;[Name of the _***Export&nbsp;Folder***](#name-of-the-exportfolder)  
&nbsp;&nbsp;&nbsp;[Serviced Synchronization Target Folder](#serviced-synchronization-target-folder)  
[Usage](#usage)  
&nbsp;&nbsp;&nbsp;[Serviced or not serviced](#serviced-or-not-serviced)  
&nbsp;&nbsp;&nbsp;[Enabling the services](#enabling-the-services)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Enabling the _Export_ service](#enabling-the-export-service)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Enabling the _Update_ service](#enabling-the-update-service)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Enabling the _Synchronization_ service](#enabling-the-synchronization-service)  
&nbsp;&nbsp;&nbsp;[Using the Synchronization Service](#using-the-synchronization-service)     
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Steps](#steps)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Worksheet synchronization](#worksheet-synchronization)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[References synchronization](#references-synchronization)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[VB-Components synchronization](#vb-components-synchronization)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Names synchronization](#names-synchronization)  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Synchronization of obsolete Names](#synchronization-of-obsolete-names)   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Synchronization of new Names](#synchronization-of-new-names)   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Manual pre-synchronization preparation](#manual-pre-synchronization-preparation)   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;[Sheet-Shape synchronization](#sheet-shape-synchronization)  
[Other](#other)  
&nbsp;&nbsp;&nbsp;[Common Components](#common-components)  
&nbsp;&nbsp;&nbsp;[The Common Components folder](#the-common-components-folder)  
&nbsp;&nbsp;&nbsp;[The Revision Number](#the-revision-number)  
&nbsp;&nbsp;&nbsp;[Pausing/continuing the CompMan Add-in](#pausingcontinuing-the-compman-add-in)  
&nbsp;&nbsp;&nbsp;[CompMan's specific files](#compmans-specific-files)  
&nbsp;&nbsp;&nbsp;[Multiple computers involved in VB-Project's development/maintenance](#multiple-computers-involved-in-vb-projects-developmentmaintenance)  
[Contribution](#contribution)</sup>

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
| _Servicing&nbsp;Workbook_ | The service providing Workbook, either the _[CompMan.xlsb][1]_ Workbook (when it is open) or the _CompMan Addin_ when it is set up and open. |
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
#### Serviced Development and Test Folder
The folder is essential for CompMan's  _Export Changed Components_ and or _Update Outdated Components_ service because the service is only provided for Workbooks with [enabled](#enabling-the-services) services when opened from within this folder. When no such folder is configured or the configured one is invalid services will be denied without notice for Workbooks even when service is [enabled](#enabling-the-services).
#### Export Folder
Folder within a Workbook's folder into which the _Export_ service exports modified _VB-Components_. This folder is one of the many reasons why a serviced Workbook must be the only Workbook in its parent folder. The name of the _Export Folder_ defaults to _source_ but may be changed at any time. Export folders with an outdated name will be renamed right away or when detected by the _Export Service_. 
#### Addin Folder
The folder in which the CompMan Workbook is saved to as Add-in. The folder is obligatory only when CompMan is about to be setup as _Add-in_. When no _Addin Folder_ is configured the [CompMan.xlsb][1] Workbook cannot be setup as Add-in.
#### _Clear_ the Add-in
When CompMan should not or no longer be available as Add-in this button will clear all resources except the configured Add-in folder physically.

#### Serviced Synchronization Target Folder
Folder into which a Workbook for which the [_Synchronize VB-Project_ service](#enabling-the-synchronization-service) has been enabled, is temporarily moved and opened from there in order to have its VB-Project synchronized with the same named Workbook residing in the configured [_Serviced Development and Test Folder_](#serviced-development-and-test-folder).
#### Locate/specify the _Synchronization Archive Folder_
A _Synchronization Archive Folder_ is obligatory when the _Synchronize VB-Project_ service is used. The service will archive a _Sync-Target-Workbook_ before it is synchronized with its corresponding _Sync-Source-Workbook_. When none is selected the folder will be reset to 'not configured'. 

#### CompMan Workbook auto-open
Since CompMan is a mere VB-Project development aid it makes sense for developers to have it automatically opened when Excel starts. When CompMan is used as Add-in the Add-in will automatically be opened. Once Auto-open is setup the shortcut placed in the users XLSTART folder will automatically be updated once the CompMan Workbook is opened from a different location.
#### Setup/Renew Add-in
Sets up CompMan as Add-in in the configured [_Addin Folder_](#addin-folder)
#### Pause/Continue Add-in
Only used when CompMan is maintained to enforce its use even when Add-in is open.

## Usage
### Serviced or not serviced
A Workbook will only be serviced by CompMan provided
- the ***servicing Workbook*** is open, which is either the [CompMan.xlsb][1] Workbook or the [CompMan Add-in](#compman-used-as-addin)
- a valid [_Serviced Development and Test Folder_](#serviced-development-and-test-folder) is configured
- the Workbook is [enabled](#services-enabling)/prepared for the service
- the serviced Workbook resides in the configured [_Serviced Development and Test Folder_](#configuration) and is opened from within it
- the ***serviced Workbook*** is the only Workbook in its parent folder
- WinMerge ([WinMerge English][4], [WinMerge German][5] or any other language version is installed (it is used to display the difference for any components about to be updated by the [_Update_](#enabling-the-update-service)

As a consequence from the above, a productive Workbook must not be used from within the configured [_Serviced Development and Test Folder_](#configuration). When a Workbook with any [enabled](#services-enabling)/prepared service is opened when located elsewhere the user will not be bothered by any means, i.e. will not even recognize CompMan at all - even when open/available.

### Enabling the services
>Even when a Workbook has a service enabled: When required preconditions are not met the service is denied without notice.
#### Enabling the _Export_ service
The _Export_ service is performed whenever the Workbook is saved from within the configured [_Serviced Development and Test Folder_](#configuration) and all the [preconditions](#serviced-or-not-serviced) are met..
1. Download and import the _[mCompManClient.bas][3]_ which serves as the link to the CompMan services.
2. Into the Workbook module copy the following:
```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan)
    '~~ had been opened from within the configured 'Serviced Development and Test Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub
```

#### Enabling the _Update_ service
The _Update_ service is performed whenever the Workbook is opened from within the configured [_Serviced Development and Test Folder_](#configuration) folder and all the [preconditions](#serviced-or-not-serviced) are meet.
1. Download and import the _[mCompManClient.bas][3]_ which serves as the link to the CompMan services.
2. Into the Workbook module copy the following:
```vb
Private Sub Workbook_Open()
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan)
    '~~ is opened from within the configured 'Serviced Development and Test Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub
```

3. In case the Workbook hosts _Common Components copy:
```vb
Const HOSTED_RAWS = <component-name>[,<component-name]...
```
<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;else<br>
```vb
Const HOSTED_RAWS = vbNullstring
```
The _Update_ service is provided when the Workbook is opened.

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
Synchronized are all types of VB-Components:  (_Standard&nbsp;Module_, _Data&nbsp;Module_, _Class&nbsp;Module_, _UserForm_). New components are added, obsolete components are removed, and of changed components the code is updated. |

#### _Names_ synchronization
!!! still under development !!!
Names synchronization is the most delicate of all synchronizations and requires specific attention.

##### Synchronization of obsolete Names
A Name is regarded obsolete when it is neither used in any Sync-Source-Workbook Worksheets cell's formula nor in any line of vba code. This may occur to Names which do exist in the Sync-Source-Workbook and the Sync-Target-Workbook. An Name detected obsolete in the Sync-Source-Workbook will be exempted from synchronization, an obsolete Name in the Sync-Target-Workbook will be removed. 

##### Synchronization of new Names
A Name is considered new when the Name's 'mere name' [^2] exists in the Sync-Source-Workbook but not in the Sync-Target-Workbook. However, when a new Name is synchronized it may refer-to the wrong range in the Sync-Target-Workbook which is a potentially serious issue:

While the _Sync-Source-Workbook_ is under development, maintenance, etc., the Workbook of which it is a copy remains "productive". The advantage of this approach, a minimized downtime for the productive Workbook comes with the downside that rows and even columns may be added which may affect the range a Name refers-to. On the other side, sheet-design changes in the Sync-Source-Workbook may add or remove cells/ranges as well. Both will results in a synchronization mess impossible to be sorted out.

>The only way out of the dilemma is a [manual pre-synchronization preparation](#manual-pre-synchronization-preparation) flanked by very careful checks before a new Name is added. New inserted or deleted ranges (columns, rows, cells) are not synchronized. When the Workbook's modifications include new and/or inserted ranges these need to be [synchronized manually beforehand](#manual-pre-synchronization-preparation) - which is supported/enabled by the open-decision-dialog displayed when the _Sync-Target-Workbook_ is opened.
>New names with wrong referred range have to be avoided by ****manually establishing the new Name in the Sync-Target-Workbook in a manual pre-synchronization effort****. A corresponding warning is displayed with the synchronization dialog and the pre-synchronization can be made by interrupting the synchronization and continuing it afterwards.

##### Manual pre-synchronization preparation
When a synchronization dialog is terminated without any action the whole synchronization will be interrupted leaving the Sync-Target-Workbook's working copy open. However it is not recommendable to do the manual work in this open Workbook but rather close it by saving the synchronizations already performed and opening the origin Sync-Target-Workbook again by selection the preparation option from the displayed ope-decision dialog. When the Sync-Target-Workbook is closed and re-opened the option ***Continue ongoing synchronization*** will continue synchronizing the outstanding.   

[^2]: A Name objects 'mere name' is one without a sheet-name-prefix

#### _Sheet-Shape_ synchronization
New Shapes (including ActiveX-Controls) are added, obsolete Shapes are removed. The Properties of all Shapes are synchronized. However, though largely covered the properties synchronization may still be incomplete. 

### Other
#### Status of the Addin
| Status             | Meaning |
|--------------------|---------|
| configured | A valid, existing [_Addin folder_](#add-in-folder) is specified |
| paused     | The Addin is currently paused. I.e. even when open it is programmatic-ally deactivated |
| open       | The CompMan-Addin open, it may still be paused however |
| not open   | The [CompMan.xlsb][1] Workbook is setup as Addin but the Addin is yet not open! Excel opens the Addin only when at least one open Workbook's VB-Project has referenced it. |
| setup      | The [CompMan.xlsb][1] Workbook is setup as Addin, i.e. it is available in the configured [_Addin folder_](#add-in-folder)  |
 
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
| _CompMan.Service.trc_ | The serviced Workbook's parent folder | Execution trace of the performed CompMan service, available only when the _Conditional Compile Argument_ `ExecTrace = 1` is set in the servicing Workbook which is either the CompMan.xlsb Workbook directly or the CompMan.xlam Addin instance of it. |
|  _CompMan.Service.log | The serviced Workbook's parent folder | Log file for the executed CompMan services. |

#### Multiple computers involved in VB-Project's development/maintenance
I do use two computers at two locations and prefer not to be bound to one. Some may prefer a network drive others a cloud based service. I prefer GitHub which makes using several computers very comfortable. In any case there is a certain need to prevent updates of used _Common&nbsp;Components_ with outdated hosted/raw versions.

#### CompMan used as Addin
Life is easy when the [CompMan.xls][1] Workbook is open and that's why it is possible to configure an Auto-open for it. Alternatively, CompMan may be setup as Add-in. This setup will also configure an Auto-open for it. Workbooks with [enabled services](#enabling-the-services) will be served by either of the two depending on which one is open.

##### Pausing/continuing the CompMan Addin
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open. Pausing the Addin is only a CompMan development feature. When the Addin is paused while the [CompMan.xlsb][1] is open CompMan works as if the Addin were not setup which means the services are directly provided by the open [CompMan.xlsb][1]. When the [CompMan.xlsb][1] Workbook is closed and an Addin had been setup the Addin will be _continued_ automatically. This ensures that the Addin is available for the [CompMan.xlsb][1] Workbook when it is opened again.
> The _CompMan Addin_ is the only means which allows to update outdated _Used&nbsp;Common&nbsp;Components_ in the [CompMan.xlsb][1] Workbook. I.e. the development instance of the Addin.
## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].


[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/2021/02/06/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
[4]:https://winmerge.org/downloads/?lang=en
[5]:https://winmerge.org/downloads/?lang=de
[6]:https://github.com
[7]:https://warbe-maker.github.io/vba/common/2021/02/19/Common-VBA-Components.html
[8]:https://github.com/warbe-maker/Common-VBA-Excel-Component-Management-Services