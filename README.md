# Management of Excel VB-Project Components
- **Export** any _Components_ of which the code has changed - when the Workbook save
- **Manage** used _Common Components_
  - **Maintain** a _Revision Number_ when the code is modified in it's hosting Workbook
  - **Update** _Common Components_ used in Workbooks when opened
- **Synchronize** a maximum of differences between _VB-Projects_ (excluding the data in Worksheets)
 
Also see the [Programmatically updating Excel VBA code][2] post for more details.

## Disambiguation

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic term for any kind of _VB-Project-Component_ (_Class Module_,  _Data Module_, _Standard Module_, or _UserForm_  |
|_Common&#8209;Component_ | A _Component_ which is hosted in one (possibly dedicated) Workbook in which it is developed, maintained and tested and used by other  _Workbooks/VB-Projects_. I.e. a _Common-Component_ exists as one raw and many clones (following GitHub terminology)  |
|_Clone&#8209;Component_ <br> | The copy of a _Raw&#8209;Component_ in a _Workbook/VP&#8209;Project_ using it. _Clone-Components_ may be automatically kept up-to-date by the _UpdateRawClones_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_ |
|_Procedure_           | Any public or private _Property_, _Sub_, or _Function_ of a _Component_|
|_Raw&#8209;Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
|_Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_             | Generic term for any _Public Property_, _Public Sub_, or _Public Function_ of a _Component_ |
|_VB&#8209;Project_    | Used synonymous with Workbook |
|_Source&#8209;Workbook/<br>Source&#8209;VB&#8209;Project_   | The temporary copy of productive Workbook which becomes by then the _Target-Workbook/Project for the synchronization.|
|_Target&#8209;Workbook<br>Target&#8209;VB&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
| _Workbook&#8209;Folder_ | A folder dedicated to a _Workbook/VB-Project_ with all its Export-Files (in a \source sub-folder). When the folder is the equivalent of a GitHub repo it may contain other files like a README and a LICENSE (provided GitHub is used for the project's versioning which not only  recommendable but also pretty easy to use.|

# Services
## Export service (_ExportChangedComponents_)
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export-File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export-File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## Update service (_UpdateRawClones_)
Used with the _Workbook\_Open_ event, checks each _Component_ in the VB-Project for being known/registered as _Raw-Component_ hosted by another _VB-Project_ by comparing the Export-Files. When they differ, the raw's _Export-File_ is used to 'renew' the _Clone-Component_.

## Synchronization service (_SyncVBProjects_)
### Aim, Purpose
The service is meant for a productive Workbook which is temporarily copied for the modification of the VB-Project. Because the productive Workbook remains in use until the modification is finished not only the down time but also the time stress is minimized.

### Coverage, synchronization extent

| Item                   | Extent of synchronization |
| ---------------------- | ------------------------- |
|_References_            | New, obsolete             |
|_Standard&#8209;Modules_<br>_Class&#8209;Modules_<br>_UserForms_| New, obsolete, code change |
|_Data&#8209;Modules_           |**Workbook**: Code change<br>**Worksheet**: New, obsolete, code change |
|_Shapes_                | New, obsolete, properties (largely covered, may still be incomplete) |
|_ActiveX&#8209;Controls_      | New, obsolete, properties |
|_Range&#8209;Names_           | Obsolete, new only when it is asserted that all relevant columns and rows are synched beforehand (pending implementation, see Worksheet synchronization below) |
|_Named&#8209;Range&#8209;Properties_ | Named ranges already in sync (currently implemented)

### Worksheet synchronization
While the code of a sheet is fully synchronized design changes such like insertion of new columns/rows or cell formatting remain a manual task. Because a Worksheet's Name and its CodeName may be changed this would be interpreted either as new or obsolete sheets. It is therefore explicitly required to assert that only one of the two is changed but never both at once.

# Installation
1. Download and open [CompMan.xlsb][1]

2. Optionally use the _Setup/Renew_ button to establish a CompMan-Addin. The service asks for two required basic configurations
   - a folder for the Addin which will become the _Application.AltStartupPath_ and therefore defaults to it when already used
   - a _Serviced-Root-Folder_ which is used to serve only Workbooks under this root but not when they are located elsewhere outside
 
Once the _CompMan-Addin_ is established it will automatically be opened when Excel starts.

# Usage
## Export (_ExportChangedComponents_) and Update (_UpdateRawClones_) service 
### Preconditions
The services will be denied when any of the following preconditions is not met:
1. The basic configuration - confirmed with each Setup/Renew is complete and valid
2. The serviced Workbook resides in a sub-folder of the configured _ServicedRootFolder_
3. The serviced Workbook is the only Workbook in its parent folder
4. The CompMan services are not _Paused_
5. WinMerge is installed

### Common usage requirements
1. In any Workbook (specifically in those which host a _Common-Component_) copy the module _[mCompManClient][3]_ from the open _[CompMan.xlsb][1]_  Workbook into the Workbook (drag and drop in the VBE)
2. For the **Export** service (_ExportChangedComponents_)<br>Crucial for all Workbooks which either **host** a _Common-Component_ or may be copied for synchronization one time (which will rely on up-to-date Export-Files).<br>In the concerned Workbook's Workbook-Component copy:
```vb

Private Const HOSTED_RAWS = ""      ' Comma delimited names of Common Components hosted, developed,
                                    ' tested, and provided by this Workbook - if any

```

```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub
```

3. For the **Update** service (_UpdateRawClones_)<br>
Essential for the update of any _Clone-Component_ of which the _Raw-Component_ had changed, copy the following into the _Workbook\_Open_ event procedure:
```vb
Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
End Sub
```

## Synchronization (_SyncVBProjects_) service 
When either the _[CompMan.xlsb][1]_ Workbook or the corresponding _CompMan-Addin_ is open, in the _Immediate Window_ enter<br>
`mService.SyncVBProjects`<br>
A dialog will open for the selection of the source and the target Workbook through their file names regardless the are already open. To avoid a possible irritation, opening them beforehand may be appropriate. In case there are some not yet up-to-date used _Common-Components_ the update service will run and display a confirmation dialog. For more details using this service see th post [Programmatically-updating-Excel-VBA-code][2]. 

### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the Workbook [CompMan.xlsb][1] is open.
  
## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].

[1]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/vba/excel/code/component/management/2021/03/22/Programatically-updating-Excel-VBA-code.html
[3]:https://gitcdn.link/cdn/warbe-maker/VBA-Components-Management-Services/master/source/mCompManClient.bas
