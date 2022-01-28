## Management of Excel VB-Project Components
- **Export** any _Component_ of which the code has changed (with the Workbook_BeforeSave_ event)
- **Update** outdated _Used Common Components_
- **Synchronize** a maximum of differences between _VB-Projects_ (excluding the data in Worksheets)
- **Manage** _Hosted Common Components_
 
Also see the [Programmatically updating Excel VBA code][2] post for more details.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic term for any kind of _VB-Project-Component_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_  |
|_Common&#8209;Component_ | A _Component_ which is hosted in one (possibly dedicated) Workbook in which it is developed, maintained and tested and used by other  _Workbooks/VB-Projects_. I.e. a _Common-Component_ exists as one raw and many clones (following GitHub terminology)  |
|_Used&nbsp;Common&nbsp;Component_ | The copy of a _Raw&#8209;Component_ in a _Workbook/VP&#8209;Project_ using it. _Clone-Components_ may be automatically kept up-to-date by the _UpdateUsedCommonComponents_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_ |
|_Procedure_           | Any public or private _Property_, _Sub_, or _Function_ of a _Component_|
|_Raw&nbsp;Common&nbsp;Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
|_Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
|_VB&#8209;Project_    | Used synonymous with Workbook |
|_Source&#8209;Workbook/<br>Source&#8209;VB&#8209;Project_   | The temporary copy of productive Workbook which becomes by then the _Target-Workbook/Project for the synchronization.|
|_Target&#8209;Workbook<br>Target&#8209;VB&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
| _Workbook&#8209;Folder_ | A folder dedicated to a _Workbook/VB-Project_ with all its Export-Files (in a \source sub-folder). When the folder is the equivalent of a GitHub repo it may contain other files like a README and a LICENSE (provided GitHub is used for the project's versioning which not only  recommendable but also pretty easy to use.|

## Services
### Export service (_ExportChangedComponents_)
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export-File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export-File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

### Update service (_UpdateUsedCommonComponents_)
Used with the _Workbook\_Open_ event, checks each _Component_ in the VB-Project for being known/registered as _Raw-Component_ hosted by another _VB-Project_ by comparing the Export-Files. When they differ, the raw's _Export-File_ is used to 'renew' the _Clone-Component_.

### Synchronization service (_SyncVBProjects_)
The service is meant for a productive Workbook which is temporarily copied for the modification of the VB-Project. The service allows the productive Workbook to be continuously used while the modification is done. When the modification is finished, only a significantly shorter downtime is used to synchronize all made modifications with the productive version. The synchronization is done to  the following extent:

| Item                   | Extent of synchronization |
| ---------------------- | ------------------------- |
|_References_            | Complete (new References are added, obsolete References are removed  |
| _Component&nbsp;types:<br>-&nbsp;Standard&nbsp;Module_<br>-&nbsp;_Class&nbsp;Modules_<br>\-&nbsp;_UserForm_| Complete (new components are added, obsolete components are deleted, and changed components are replaced |
|_Data&nbsp;Modules_    |**Workbook**: Code change<br>**Worksheet**: New, obsolete, and code changes |
|_Shapes_                | New, obsolete, properties (though largely covered they may still be incomplete) |
|_ActiveX&nbsp;Controls_      | New, obsolete, properties |
|_Range&nbsp;Names_           | Obsolete, new only when it is asserted that all relevant columns and rows are synched beforehand (pending implementation, see Worksheet synchronization below) |
|_Named&nbsp;Range&nbsp;Properties_ | Named ranges already in sync (currently implemented)

### Worksheet synchronization
While the code of a sheet is fully synchronized, design changes such like insertion of new columns/rows and formatting changes remain a manual task.\ A Worksheet's _Name_ may be changed as well as its _CodeName_. However, when both are changed this would be interpreted as new and the old sheet as an obsolete one - which is definitely not what was intended. In order to avoid a disastrous synchronization error such a change demands the explicit assertion that only one of the two but never both are changed at a time.

## Installation
1. Download and open [CompMan.xlsb][1]<br>When the Workbook is opened for the first time it will show a dialog for the reqired _Basic Configuration_. That's all unless one prefers to have the services available as Addin which requires the second step. 

2. Optionally use the _Setup/Renew_ button to establish the [_Component Management_ as Addin](#component-management-as-addin). The service requires to re-confirm the [Component Management configuration](#component-management-configuration)
   - a folder for the Addin which will become the _Application.AltStartupPath_ and therefore defaults to it when already in use
   - a _Serviced-Folder_ in which a Workbook must be located (in its dedicated folder by the way) in order to get serviced by _CompMan_
 
Once the _CompMan-Addin_ is established it will automatically be opened and available when Excel starts - unless it is not removed from the Addin-Folder.

## Usage

> #### Serviced or not serviced
> Even when a Workbook is prepared for being serviced by _CompMan_ **nothing at all will happen** when the Workbook resides outside the _Serviced Folder_. In other words, when the Workbook becomes productive it should be copied to another location in order not to be bothered by CompMan services. 

### Preconditions for any Workbook for being serviced by _CompMan_

For a Workbook the following preconditions must be met in order to get is serviced by CompMan:
1. The Workbook must be located in the configured [Serviced Folder](#basic configuration) 
2. The Workbook must resides exclusively in its dedicated folder, in other words it is not the only Workbook in its parent folder
4. Either the  _[CompMan.xlsb][1]_ Workbook is open or the CompMan-Addin is setup
5. WinMerge is installed
6. The _[mCompManClient.bas][3]_ had been downloaded and imported. This component is the link to the services either provided by the _CompMan-Addin_ or directly by the _[CompMan.xlsb][1]_ Workbook
7. The Workbook component is prepared as follows
```vb
Private Const HOSTED_RAWS = "mFile" ' The hosted 'Raw Common Component' (if any)

Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub
```

### Other

#### Synchronization (_SyncVBProjects_) service
> There is no user interface for this service. As it is a mere development task it is to be initiated from the _Immediate Window_ of the VB Editor when the _[CompMan.xlsb][1]_ Workbook is open.

A dialog will open for the selection of the _Source-Workbook_ and the _Target Workbook_ - which may be already open however. To avoid a possible irritation, opening them beforehand may be appropriate. In case there are some _Outdated Used Common-Components_ the update service will run and display a confirmation dialog.

#### _Common Components_ folder
_CompMan_ maintains for each _Raw Common Component_ a copy of the _Export File_ in a _Common Components_ folder. These _Export Files_ are the source for a service Workbook's _Outdated Used Common Components_. When a _Hosted Raw Common Component_ is modified it is not only exported like any other component but also copied to the _Common Components_ folder thereby increasing a _Revision Number_. 

#### _Revision Number_
CompMan is pretty much focused on _Common Components_. In order to prevent updates of _Used Common Components_ with outdated raw versions CompMan maintains a _Revision Number_ for them which is increased whenever a new modified version is exported. The _Revision Number_ is maintained in a file _ComCompsHosted.dat_ located in the Workbook folder and kept in sync with the Revision Number_ in a file _ComCompsSaved.dat_ located in the [_Common Compüonents_ folder](#common-components-folder).

#### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open. Pausing the Addin is only a CompMan development feature. When the Addin is paused while the [CompMan.xlsb][1] is open CompMan works as if the Addin were not setup which means the services are directly provided by the open [CompMan.xlsb][1]. When the [CompMan.xlsb][1] Workbook is closed and an Addin had been setup the Addin will be _continued_ automatically. This ensures that the Addin is available for the [CompMan.xlsb][1] Workbook when it is opened again.
> The _CompMan Addin_ is the only means which allows to update _Outdated Used Common Components_ in the [CompMan.xlsb][1].

#### Basic configuration 
| Item | Description |
|------|-------------|
| _Serviced&nbsp;Folder_ | The folder in which a Workbook (in its dedicated folder) must be located in order to get serviced by _CompMan_ |
| _Addin&nbsp;Folder_ | Obligatory only when CompMan is Setup/Renewed. This folder defaults to the _Application.AltStartupPath_ when one is already specified/in use. The specified or confirmed folder is (or becomes) the _Application.AltStartupPath_. |
| _Export&nbsp;Folder_ | Name of the folder (defaults to _source_) for the _Export Files_ of new or modified components.<br>Note: CompMan only exports new or modified components. |

#### CompMan specific PrivateProfile files

| File                     | Location             | Purpose               |
|--------------------------|----------------------|-----------------------|
| _ComCompsHosted.dat_     | Dedicated folder of a Workbook which hosts a _Raw&nbsp;Common&nbsp;Component_ | Properties of the Raw component |
| _ComCompsUsed.dat_       | Dedicated folder of a Workbook which uses a _Common&nbsp;Component_  | Properties of all _Used&nbsp;Common&nbsp;Components_ specifically the current _Revision&nbsp;Number_ |
| _ComComps-RawsSaved.dat_ | _Common&nbsp;Components_ folder | Properties of all known Raw_Common Components_, source for the _UpdateOutdatedCommonComponents_ service. |

#### Multiple computers involved in VB-Project's development/maintenance
I do use two computers at two locations and prefer not to be bound to one. Some may prefer a network drive others a cloud based service. I prefer GitHub which makes using several computers very comfortable. In any case there is a certain need to prevent updates of used Common Components with outdated hosted/raw versions.

## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].

[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/vba/excel/code/component/management/2021/03/22/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
