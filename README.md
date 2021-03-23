# Management of Excel VB-Project Components
Installation, export, update, and synchronization services dealing with Excel-VBComponents. 
See also [Programatically updating Excel VBA code][2]


The services are available when the [CompMan.xlsb][1] Workbook is downloaded and opened and this Workbook also offers to establish (or renew when modified) an Addin-Workbook by a _Renew_ service.<br>
See also [Programmatically updating Excel VBA code][2]

## Disambiguation
Terms used in this VB-Project and all posts related to the matter.

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB&#8209;Project_ term for a _VB-Project-Component_ which may be a _Class Module_ , a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_ | A _VB-Component_ which is hosted in one and commonly used by two or more _VB-Projects_ |
|_Clone&#8209;Component_ <br> | The copy of a _Raw&#8209;Component_ in a _VP&#8209;Project_ using it. _Clone-Components_ are kept up-to-date by an _UpdateRawClones_ service.<br>The term _clone_ is borrowed from GitHub but has a slightly different meaning because the clone is usually not maintained but the _raw_. |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw&#8209;Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw&#8209;Host_ Workbook. The term _raw_ is borrowed from GitHub and indicates the original version of something |
|_Raw&#8209;Host_      | The Workbook/_VB-Project_ which hosts the _Raw-Component_ |
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB&#8209;Clone&#8209;Project_ | A _VP-Project_ which is a copy (i.e regarding the VB-Project code a clone) of a corresponding  _VB&#8209;Raw&#8209;Project_. The code of the clone project is kept up-to-date by means of a code synchronization service. |
|_VBProject_     | In the present case this term is used synonymously with Workbook |
|_Source&#8209;Workbook/VBProject_   | The temporary copy of productive Workbook which becomes by then the _Target-Workbook/Project for the syncronization.|
| _Workbook-_, or<br>_VB&#8209;Project&#8209;Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export-Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable.|

# Services
## _ExportChangedComponents_
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export-File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export-File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## _UpdateRawClones_
Used with the _Workbook\_Open_ event, checks each _Component_ in the VB-Project for being known/registered as _Raw-Component_ hosted by another _VB-Project_ by comparing the Export-Files. When they differ, the raw's _Export-File_ is used to 'renew' the _Clone-Component_.

## _SyncVBProject_
Synchronizes the code of a _Target-Workbook/VBroject_ with a corresponding _Source-Workbook/VBProject_ with the following covered:
- _Standard Modules_ new, obsolete, code change
- _Class Modules_ new, obsolete, code change
- _UserForms_ new, obsolete, code/design change
- _Worksheets_ new, obsolete, code change
- _Workbook_ (Document-Module): code change
- _Shapes_ new, obsolete, properties
- _References_ new, obsolete

## _UpdateRawClones_
The service is used with the _Workbook\_Open_ event. It checks each _Component_ for being known/registered as _Raw-Component_,  _hosted_ by another _VB-Project_ - which means it is a _Clobmne-Component_. If yes their code is compared and suggested for being updated if different.

## _SyncTargetWithSourceWb_

**pending implementation**<br>

### Aim, Purpose
Service for temporarily copied productive Workbooks for modifying the VB-Project while the productive Workbook remains in use. By this minimizing the down time of the productive Workbook to the time required for the "back-syncronization" of the modified VB-Project.

### Coverage, synchronization extent

| Item        | Extent of synchronization |
| ----------- | ------------------------- |
|_References_ | New, obsolete |
|_Standard Modules_<br>_Class Modules_<br>_UserForms_| New, obsolete, code change |
|_Data Module_|**Workbook**: Code change<br>**Worksheet**: New, obsolete, code change (see [Worksheet synchronization](#worksheet-synchronization) and [Planning the release of a VB-Project modification](#planning-the-release-of-a-vb-project-modification))|
|_Shapes_ | New, obsolete, properties (may still be incomplete) |
|_ActiveX-Controls_| None. May be added in future |

### Worksheet synchronization
While the code of a sheet is fully synchronized design changes such like insertion of new columns/rows or cell formatting remain a manual task. Because a Worksheet's Name and its CodeName may be changed this would be interpreted either as new or obsolete sheets. It is therefore explicitly required to assert that only one of the two is changed but never both at once.

# Installation
The _Component Management Services_ are available when the _[development instance Workbook][1]_ is downloaded and opened (see [Usage without Addin instance](#usage-without-addin-instance)). Alternatively the services can be made available as an Addin-Workbook.

1. Download and open [CompMan.xlsb][1]

2. Optionally use the _Setup/Renew_ button to establish a CompMan-Addin. The service asks for two required basic configurations
  - a dedicated Addin-folder for the Addin-Workbook - preferably a dedicated folder like ../CompMan/Addin
  - a _Serviced-Root-Folder_ which is used to serve only Workbooks under this root but not when they are located elsewhere outside
 
Once the Addin is established it will automatically be loaded with the first Workbook opened which ha a VBProject with a _Reference_ to it. When no Workbook refers to it, the Addin may be made available at any time via the CompMan-Development-Instance-Workbook.

## Usage
### Common preconditions
Every service will be denied unless the following preconditions are met:
1. The basic configuration - confirmed with each Setup/Renew of the _CompMan-Addin-Folder_ - is complete and valid
2. The serviced Workbook resides in a sub-folder of the configured _ServicedRootFolder_
3. The serviced Workbook is the only Workbook in its parent folder
4. The CompMan services are not _Paused_
5. WinMerge is installed

### Common usage requirement
In any Workbook either using the _ExportChangedComponents_ and/or the _UpdateChangedRawClones_ service copy the following in a Standard-Module called _mCompManClient_:
```vb
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mCompManClient, optionally used by any Workbook to:
' - update used 'Common-Components' (hosted, developed, tested,
'   and provided, by another Workbook) with the Workbook_open event
' - export any changed VBComponent with the Workbook_Before_Save event.
'
' W. Rauschenberger, Berlin March 2021
'
' See also Github repo:
' https://github.com/warbe-maker/Excel-VB-Components-Management-Services
' ----------------------------------------------------------------------

Public Sub CompManService(ByVal cm_service As String, _
                          ByVal hosted As String)
' -----------------------------------------------------
' Execution of the CompMan service (cm_service) pre-
' ferably via the CompMan-Addin or when not available
' alternatively via the CompMan.xlsb Workbook.
' -----------------------------------------------------
    Const COMPMAN_BY_ADDIN = "CompMan.xlam!mCompMan."
    Const COMPMAN_BY_DEVLP = "CompMan.xlsb!mCompMan."
    
    On Error Resume Next
    Application.Run COMPMAN_BY_ADDIN & cm_service, ThisWorkbook, hosted
    If Err.Number = 1004 Then
        On Error Resume Next
        Application.Run COMPMAN_BY_DEVLP & cm_service, ThisWorkbook, hosted
        If Err.Number = 1004 Then
            Application.StatusBar = "'" & cm_service & "' neither available by '" & COMPMAN_BY_ADDIN & "' nor by '" & COMPMAN_BY_DEVLP & "'!"
        End If
    End If
End Sub
```

### Using the _ExportChangedComponents_ service
This service is crucial for all Workbooks which either host a commonly used component or which may become the source for a synchronization because both rely on up-to-date Export-Files.

In the concerned Workbook's Workbook-Component copy:
```vb
                                    ' -------------------------------------------------------------
Private Const HOSTED_RAWS = ""      ' Comma delimited names of Common Components hosted, developed,
                                    ' tested, and provided by this Workbook - if any
                                    ' -------------------------------------------------------------
```

and in the concerned Workbook's Workbook_BerforeSave event procedure copy:
```vb
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
End Sub
```

### Using the _UpdateRawClones_ service
In the concerned Workbook's Workbook_Open event procedure copy:
```vb
Private Sub Workbook_Open()
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
End Sub

```

### Using the _SyncTargetWithSource_ service
When either the [CompMan.xlsb][1] Workbook or the corresponding CompMan-Addin is open, in the _Immediate Window_ enter<br> `mService.SyncTargetWithSource`<br>A dialog will open for the selection of the source and the target Workbook. They are selected by their files even when already open. To avoid a possible irritation, opening them beforehand may be appropriate in case there are some not yet up-to-date used _Common-Components_.


### Pausing/continuing the CompMan Add-in
Use the corresponding command buttons when the [CompMan.xlsb][1] Workbook is open.

  
## Contribution
Contribution of any kind is welcome raising issues or by commenting the corresponding post [Programmatically-updating-Excel-VBA-code][2].


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompMan.xlsb
[2]:https://warbe-maker.github.io/vba/excel/code/component/management/2021/03/02/Programatically-updating-Excel-VBA-code.html
