## Management of Excel VB-Project Components
- **Export** any _Component_ the code has changed (automated with the Workbook_BeforeSave_ event)
- **Update** outdated _Used&nbsp;Common&nbsp;Components_
- **Synchronize** two Workbooks (the differences between their _VB-Projects_ excluding the data in Worksheets)
- **Manage** _Hosted Common&nbsp;Components_
 
Also see the [Programmatically updating Excel VBA code][2] post for this subject.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
| _Component_       | Generic term for any kind of _VB-Project-Component_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_  |
| _Common&nbsp;Component_ | A _Component_ which is hosted in one (possibly dedicated) Workbook in which it is developed, maintained and tested and used by other  _Workbooks/VB-Projects_. I.e. a _Common-Component_ exists as one raw and many clones (following GitHub terminology)  |
|_Used&nbsp;Common&nbsp;Component_ | The copy of a _Raw&#8209;Component_ in a _Workbook/VP&#8209;Project_ using it. _Clone-Components_ may be automatically kept up-to-date by the _UpdateOutdatedCommonComponents_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_ |
| _Procedure_           | Any public or private _Property_, _Sub_, or _Function_ of a _Component_|
| _Raw&nbsp;Common&nbsp;Component_ | The instance of a _Common&nbsp;Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
| _Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
| _Servicing&nbsp;Workbook_ | The service providing Workbook, either the _[CompMan.xlsb][1]_ Workbook - when it is open or the CompMan Addin when it is set up. |
| _Serviced&nbsp;Workbook_ | The Workbook prepared for being serviced, provided it is located within the _Serviced&nbsp;Folder_.
|_VB&#8209;Project_    | Used synonymous with Workbook |
| _Source&#8209;Workbook/<br>Source&#8209;VB&#8209;Project_   | The temporary copy of productive Workbook which becomes by then the _Target-Workbook/Project for the synchronization.|
| _Target&#8209;Workbook<br>Target&#8209;VB&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
| _Workbook&#8209;Folder_ | A folder dedicated to a _Workbook/VB-Project_ with all its Export-Files (in a \source sub-folder). When the folder is the equivalent of a _GitHub repo_ it may contain other files like a README and a LICENSE (provided GitHub is used for the project's versioning which not only  recommendable but also pretty easy to use.|

## Services
### The _ExportChangedComponents_ service
Used with the _Workbook\_BeforeSave_ event all component's code is compared with their previous _Export&nbsp;File_ and when the code has changed the component is exported again in the configured _Export Folder_ the name defaults to _source_. By the way the _Export&nbsp;Files_ are a perfect backup in case Excel opens a Workbook with a fucked-up VB-Project.

### The _ExportChangedComponents_ service for _Raw Common Components_
The initial intention for the development of CompMan was to keep _Common&nbspComponent_ up-to-date in all VB-Projects using them. While the export service applies to all kinds of components in a VB-Project the handling of _Raw&nbsp;Common&nbsp;Components_ is specific. The service registers all hosted _Raw&nbsp;Common&nbsp;Components_ by increasing a [_Revision Number_](#the-revision-number) with each export and additionally copies the _Export&nbsp;File_ to a _Common Components_ folder. The _Export&nbsp;Files_ in this folder are the source for the [_UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service). This means that the hosting Workbook is not in charge with this service.<br>
The service also checks whether a  _Used&nbsp;Common&nbsp;Component_ has been modified within the VB-Project using it - which may happen accidentally - and registers a **due modification revert alert** displayed when the Workbook is opened subsequently and the [_UpdateOutdatedCommonComponents_ service](#the-updateoutdatedcommoncomponents-service) is about to revert the made modifications, allowing to display the code difference (using WinMerge).

### The _UpdateOutdatedCommonComponents_ service
Used with the _Workbook\_Open_ event all  _Used&nbsp;Common&nbsp;Components_ are checked whether they are outdated. In case a dialog is displayed which allows to display the code difference (by means of WinMerge) perform the update or skip it. The update uses the  _Export&nbsp;File_ of the _Raw&nbsp;Common&nbsp;Component_ in the _Common&nbspComponents_ folder. This service is the core service and most critical service provided by CompMan. Excel may every now and then close the serviced Workbook when code is updated. Fortunately the Workbook can be opened again and the update service continues.  

### The _Synchronize Workbooks_ service

#### General
The service allows a productive Workbook to remain in use while its VB-Project is developed/modified/maintained in a copy of it. When all changed had been done the productive version is synchronized with the changed copy with the benefit of a significantly shorter downtime for the productive Workbook. 
> The service has no user interface. The service is invoked when the _Sync-Target-Workbook_ (prepared for the service!) is opened from within the configured _Synchronization Target Folder_ while the _Sync-Source-Workbook_ (its dedicated Workbook folder respectively) resides in the _Serviced Development & Test Folder_.

#### Synchronization extent

| Synchronized item                  | Synchronization details |
|------------------------------------|---------------------------|
| _References_                       | New References are added and obsolete References are removed  |
| _Components_                       | All types (_Standard&nbsp;Module_, _Data&nbsp;Module_, _Class&nbsp;Module_, _UserForm_). New components are added, obsolete components are removed, and of changed components the code is updated. |
| _Worksheets_                       | New Worksheets are added, obsolete Worksheets are removed, changed Worksheet Names are synchronized, changed Worksheet-Code-Names are synchronized (see [restrictions](#worksheet-synchronization-restrictions) below).|
| _Sheet-Shapes_                     | New Shapes (including ActiveX-Controls) are added, obsolete Shapes are removed, the Properties of all Shapes are synchronized (though largely covered may still be incomplete) |
| _Range&nbsp;Names_                 | New Range-Names are added, obsolete Range-Names are removed. Attention! The synchronization of new Range-Names which concern new columns or rows depend on (manually beforehand) synced new rows and columns!|
| _Named&nbsp;Range&nbsp;Properties_ | Named ranges already in sync (currently implemented)

> ***Worksheet synchronization restriction!***  
Never change both, the _Name_ ***and*** the _CodeName_ of a Worksheet! When a Worksheet's _Name_ ***and*** its _CodeName_ is changed at the same time the concerned sheet will be considered new and the (no longer identifiable as such) corresponding sheet will be considered obsolete - which in such a case is definitely not what was intended.

## Installation, provision of the CompMan services
1. Download and open [CompMan.xlsb][1] <br> When the Workbook is opened for the first time it will show a dialog for the required _Basic Folders Configuration_. Either the open Workbook is used or an Addin instance of it may be setup which then will be available when Excel is started (requires the next step). 

2. Optionally: Use the _Setup/Renew_ button on the displayed Worksheet to establish _CompMan_ as _Addin_ . This Setup/Renew requires to configure or re-confirm the [Basic Folders Configuration](#basic-folders-configuration). Once _CompMan_ had been established as _Addin_ the services will be available when Excel starts - needless to say: unless it is not removed from the [_Addin&nbsp;Folder_](#basic-folders-configuration).

## Usage

> ***Serviced or not serviced?***  
Even when a Workbook is prepared for being serviced by _CompMan_ **nothing at all will happen** when the Workbook resides outside the configured  [_Serviced Development & Test Folder_](#basic-folders-configuration) - or, in case of the _Synchronization-Service_, outside the configured [_Serviced Synchronization Target Folder_](#basic-folders-configuration). Consequently, a productive Workbook should not reside in a serviced folder and should not be opened from therein. 

### Preconditions for Workbooks for being serviced by CompMan

1. The _[mCompManClient.bas][3]_ had been downloaded and imported. <br>The _mCompManClient_ module serves as the link to the CompMan services.
2. For all three services: The Workbook (its dedicated folder respectively) resides in the configured [_Serviced Development & Test Folder_](#basic-folders-configuration) and is opened from there. Note: For the _Synchronization Service_ this Workbook serves as the _Sync-Source-Workbook_.
3. For the _Synchronization_ service specifically: The Workbook resides (had explicitly and temporarily for this purpose moved to) the configured [_Serviced Synchronization Target Folder_](#basic-folders-configuration) and is opened from there
3. The Workbook resides in its dedicated folder (i.e. it is the only Workbook in its parent folder)
4. Either the  _[CompMan.xlsb][1]_ Workbook is open or _CompMan_ had been setup as _Addin_
5. [WinMerge English][4] ([WinMerge German][5]) is installed
6. The Workbook module is prepared as follows<br>
```vb
' In case the Workbook hosts Raw Common Components, else `HOSTED_RAWS = vbNullString`
Private Const HOSTED_RAWS = "xxx,yyy,zzz" ' The hosted 'Raw Common Components' (comma delimited if any)

Private Sub Workbook_Open()
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan)
    '~~ is opened from within the configured 'Serviced Development & Test Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
    
    '~~ The below statement has no effect unless this (the Workbook being synchronized by CompMan)
    '~~ is opened from within the configured 'Serviced Synchronization Target Folder'.
    '~~ The statement is only required for a Workbook which may be synchronized
    mCompManClient.CompManService mCompManClient.SRVC_UPDATE_OUTDATED, HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '~~ The below statement has no effect unless this (the Workbook serviced by CompMan)
    '~~ had been opened from within the configured 'Serviced Development & Test Folder'.
    mCompManClient.CompManService mCompManClient.SRVC_EXPORT_CHANGED, HOSTED_RAWS
End Sub
```

Note that `HOSTED_RAWS = vbNullString` when the Workbook does not host a _Raw&nbsp;Common&nbsp;Component_.

#### Precondition for _CompMan's Synchronize Workbooks_ service

### Other

#### The _Common Components_ folder
_CompMan_ maintains for each _Raw&nbsp:Common&nbsp;Component_ a copy of the _Export File_ in a _Common&nbsp;Components_ folder. These _Export Files_ are the source for a serviced Workbook's outdated _Used&nbsp;Common&nbsp;Components_. When a _Hosted Raw&nbsp:Common&nbsp;Component_ is modified it is not only exported like any other component but also copied to the _Common&nbsp;Components_ folder thereby increasing a _Revision Number_. 

#### The _Revision Number_
CompMan is pretty much focused on _Common&nbsp;Components_. In order to prevent updates of _Used&nbsp;Common&nbsp;Components_ with outdated raw versions CompMan maintains a _Revision Number_ for them which is increased whenever a new modified version is exported. The _Revision Number_ is maintained in a file _ComCompsHosted.dat_ located in the Workbook folder and kept in sync with the Revision Number_ in a file _ComCompsSaved.dat_ located in the [_Common Compüonents_ folder](#common-components-folder).

#### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open. Pausing the Addin is only a CompMan development feature. When the Addin is paused while the [CompMan.xlsb][1] is open CompMan works as if the Addin were not setup which means the services are directly provided by the open [CompMan.xlsb][1]. When the [CompMan.xlsb][1] Workbook is closed and an Addin had been setup the Addin will be _continued_ automatically. This ensures that the Addin is available for the [CompMan.xlsb][1] Workbook when it is opened again.
> The _CompMan Addin_ is the only means which allows to update outdated _Used&nbsp;Common&nbsp;Components_ in the [CompMan.xlsb][1] Workbook. I.e. the development instance of the Addin.

#### Basic Folders Configuration

| Folder | Description |
|--------|-------------|
| _Serviced&nbsp;Development&nbsp;&&nbsp;Test&nbsp;Folder_ | The folder in which a Workbook (the dedicated Workbook-Folder respectively) must be located for the CompMan services _Export Changed Components_ and or _Update Outdated Components_. |
| _Addin&nbsp;Folder_ | Obligatory only when CompMan's Addin-Instance is Setup/Renewed. This folder defaults to the _Application.AltStartupPath_ when one is already specified or in use repectively. The specified or confirmed folder is (or becomes) the _Application.AltStartupPath_. |
| _Export&nbsp;Folder_ | Name of the folder (defaults to _source_) for the _Export Files_ of new or modified components.<br>Note: CompMan only exports new or modified components. |
| _Serviced&nbsp;Synchronization&nbsp;Target&nbsp;Folder_ | The folder in which a _Sync-Target-Workbook_ must reside for the CompMan service _Synchronize Workbooks_ | 

#### Summary of CompMan specific files

| File                     | Location             | Description               |
|--------------------------|----------------------|---------------------------|
| _ComCompsHosted.dat_     | The serviced Workbook's parent folder | PrivateProfile file for the registration of all _Raw&nbsp;Common&nbsp;Components_ hosted in the corresponding Workbook. Content scheme:<small><br>`[component-name]`<br>`RawExpFileFullName=<file-full-name>`<br>` RawRevisionNumber=YYYY-MM-DD.000>` |
| _ComCompsUsed.dat_       | The serviced Workbook's parent folder | Private Profile file for the registration of all _Used&nbsp;Common&nbsp;Components_. Content scheme:<small><br>`[component-name]`<br>`RawRevisionNumber=YYYY-MM-DD.000>` |
| _ComComps&#8209;RawsSaved.dat_ | [_Common&nbsp;Components_ folder](#common-components-folder) | PrivateProfile file for the registration of all known _Raw&nbsp;Common&nbsp;Components_ |
| _CompMan.Service.trc_ | The serviced Workbook's parent folder | Execution trace of the performed CompMan service, available only when the _Conditional Compile Argument_ `ExecTrace = 1` is set in the servicing Workbook which is either the CompMan.xlsb Workbook directly or the CompMan.xlam Addin instance of it. |
|  _CompMan.Service.log | The serviced Workbook's parent folder | Log file for the executed CompMan services. |

#### Multiple computers involved in VB-Project's development/maintenance
I do use two computers at two locations and prefer not to be bound to one. Some may prefer a network drive others a cloud based service. I prefer GitHub which makes using several computers very comfortable. In any case there is a certain need to prevent updates of used _Common&nbsp;Components_ with outdated hosted/raw versions.

## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].

[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/vba/excel/code/component/management/2021/03/22/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
[4]:https://winmerge.org/downloads/?lang=en
[5]:https://winmerge.org/downloads/?lang=de
