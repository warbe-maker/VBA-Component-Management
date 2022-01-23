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
1. Download and open [CompMan.xlsb][1]

2. Optionally use the _Setup/Renew_ button to establish the [_Component Management_ as Addin](#component-management-as-addin). The service requires to re-confirm the [Component Management configuration](#component-management-configuration)
   - a folder for the Addin which will become the _Application.AltStartupPath_ and therefore defaults to it when already used
   - a _Serviced-Root-Folder_ which is used to serve only Workbooks under this root but not when they are located elsewhere outside
 
Once the _CompMan-Addin_ is established it will automatically be opened when Excel starts.

## Usage

### Preconditions
First of all: Even when a Workbook is prepared for being serviced by _CompMan_ nothing at all happens when the the Workbook resides outside the _Serviced Folder_ .

For a Workbook the [Common requirements](#common-requirements) are met the services will still be denied under one of the following conditions:
1. The basic configuration is incomplete or invalid
2. The Workbook not resides exclusively in its dedicated folder, in other words it is not the only Workbook in its parent folder
4. The CompMan-Addin is not setup or _Paused_ and the _CompMan.xlsb_ Workbook which alternatively can provide the services is not open
5. WinMerge is not installed

### Preparing a Workbook for being serviced
1. Download and import _[mCompManClient.bas][3]_ from the open _[CompMan.xlsb][1]_ Workbook into the Workbook (drag and drop in the VBE)
2. Copy the below code into the Workbook module:
```vb
Private Const HOSTED_RAWS = "mFile" ' The hosted 'Raw Common Component' (if any)

Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateOutdatedCommonComponents", HOSTED_RAWS
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub
```

### Using the Export and Update service

### Synchronization (_SyncVBProjects_) service
There is no user interface for this service. As it is a mere development task it is to be initiated from the _Imediate Window_ of the VB Editor.
When either CompMan's development instance Workbook (_[CompMan.xlsb][1]_) or the corresponding Addin instance is loaded, in the _Immediate Window_ enter<br>
`mService.SyncVBProjects`<br>
A dialog will open for the selection of the source and the target Workbook through their file names regardless the are already open. To avoid a possible irritation, opening them beforehand may be appropriate. In case there are some not yet up-to-date used _Common-Components_ the update service will run and display a confirmation dialog. For more details using this service see th post [Programmatically-updating-Excel-VBA-code][2].

### Multiple computers involved in VB-Projectsd development/maintenance
I do use two computers at two locations and prefer not to be bound to one. Some may prefer a network drive others a cloud based storage. I prefer Github which provides a versioning control and development spread over several computer is just a matter of maintained clone/fork repositories. In any case there is a certain need to prevent updates of used Common Components with outdated hosted/raw versions.

#### _Common Components_ folder
Whenever a hosted Common Component had been modified it is copied to the _Common Components_ folder located in the _Serviced Root Folder_. Thereby, Common Components are available for being imported into other VB-Projects even when the Workbook it is hosted is not available. This is specifically to support when more than one computer (of the same or different users) is in charge of the development of VB-Projects. 

#### _Revision Number_
CompMan is pretty much focused on _Common Components_. In order to prevent updates of _Used Common Components_ with outdated raw versions CompMan maintains a _Revision Number_ for them which is increased whenever a new modified version is exported. The _Revision Number_ is maintained in a file _ComCompsHosted.dat_ located in the Workbook folder and kept in sync with the Revision Number_ in a file _ComCompsSaved.dat_ located in the [_Common Compüonents_ folder](#common-components-folder).



### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open.

### Component Management configuration 
#### _Serviced Root Folder_
CompMan only serves Workbooks in dedicated folders and when these folders arer located in/under the configured _Serviced Root Folder_.

- The Workbook has the _Common Component_ _**CompManClient**_ installed and the Workbook_Open and the Workbook_BeforeSave event is prepared as described above
- CompMan services are available

#### _Addin Folder_
When the Component Management is established as Addin this requires a dedicated folder. This folder defaults to the _Application.AltStartupPath_ when one is already specified/in use. The specified or confirmed folder is (or becomes) the _Application.AltStartupPath_.

## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].

[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/vba/excel/code/component/management/2021/03/22/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
