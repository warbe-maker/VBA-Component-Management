# Excel Common Components Management
Services for the management of
- individual Common Raw VBA Components hosted in one and used in many other Workbooks/VB-Projects
- VB-Raw-Projects and their corresponding VB-Clone-Project, i.e. synchronizing the code of the clone project with the code in its corresponding raw project

Whereby the services are available via an Addin-Workbook which is setup/established/renewed by means of a _Development-Instance_ (this repo!) which provides the _Renew_ service.<br>
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
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
|_VB&#8209;Raw&#8209;Project_   | A code-only _VP-Project_ of which all components are regarded _Raw-Components_. A _Raw-Project_ is kind of a template for the productive version of it. In contrast to a classic template it is the life-time raw code base for the productive _Clone-Project_.  The service and process of 'synchronizing' the productive (clone) code with the raw is part of the _Component Management_.|
| _Workbook-_, or<br>_VB&#8209;Project&#8209;Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export-Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |



# Services
## _ExportChangedComponents_
Used with the _Workbook_Before_Save_ event it compares the code of any component in a _VB-Project_ with its last _Export-File_ and re-exports it when different. The service is essential for _VB-Projects_ which host _Raw-Components_ in order to get them registered as available for other _VB-Projects_. Usage by any _VB-Project_ in a development status is appropriate as it is not only a code backup but also perfectly serves versioning - even when using [GitHub][]. Any _Component_ indicated a _hosted Raw-Component is registered as such with its _Export-File_ as the main property.<br>
The service also checks a _Clone-Component_ modified within the VB-Project using it a offers updating the _Raw-Component_ in order to make the modification permanent. Testing the modification will be a task performed with the raw hosting project.

## _UpdateRawClones_
Used with the _Workbook\_Open_ event, checks each _Component_ in the VB-Project for being known/registered as _Raw-Component_ hosted by another _VB-Project_ by comparing the Export-Files. When they differ, the raw's _Export-File_ is used to 'renew' the _Clone-Component_.

## _SyncVBProject_
Synchronizes the code of a _VB-Clone-Poject_ with the code of a corresponding _VB-Raw-Project_. _Standard Modules_ , _Class Modules_ , and _UserForms_ are synchronized by re-importing the raw's _Export-File_. _Data Modules_ indicating a Worksheet are synchronized by re-writing the code line by line taken from the raw's _Export-File_. A yet not existing Worksheet is created - which one of the circumstances the raw Workbook will be opened for in order to obtain the Worksheets name.<br>
**Constraint:** The _Data-Module_ indicating the Workbook object cannot be synchronized!

## Installation
### Installation of the _CompMan_ Add-in
1. Download and open [CompManDev.xlsb][1]
2. Follow the instructions to identify a location for the Add-in - preferably a dedicated folder like ../CompMan/Add-in. The folder will hold the following files:
   - CompMan.cfg    ' the basic configuration
   - CompMan.xlam   ' the Add-in
   - HostedRaws.dat ' the specified raws hosted in any Workbook
   - RawHost.dat    ' the Workbooks which claim raws hosted
   
3. Follow the instructions to identify a 'serviced root'
4. Use the built-in Command button to run the _Renew_ service. It will:
   - ask to confirm or change the basic configuration
   - initially setup or subsequently renew the CompMan Add-in by saving a copy  of the development instance as Add-in (mind the fact that this is a multi-step process which may take some seconds)

Once the Add-in is established it will automatically be loaded with the first Workbook opened having it referenced. See the Usage below for further required preconditions.

### Installation for Workbooks/VB-Projects hosting raws or using raw clones
1. Copy the following into the Workbook component
```vb
Option Explicit

Private Const HOSTED_RAWS = ""

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateRawClones uc_wb:=ThisWorkbook _
                           , uc_hosted:=HOSTED_RAWS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents ec_wb:=ThisWorkbook _
                                   , ec_hosted:=HOSTED_RAWS
#End If
End Sub
```
2. For a Workbook which hosts _Raw_Components_ specify them in the HOSTED_RAWS constant. If its more then one, have the component's names delimited with commas.

> ++**Be aware:**++ The Workbook component will be one of which the code cannot be updated by any means because it contains the code executed to perform the update. Thought this will only be relevant for Raw/Clone-VB-Projects which are yet not supported. However, as a consequence only calls to procedures provided with all arguments will remain in the Workbook component code and all the rest will be in a dedicated mWorkbook component.


[1]:https://gitcdn.link/repo/warbe-maker/VBA-Components-Management-Services/master/CompManDev.xlsb
[2]:https://warbe-maker.github.io/warbe-maker.github.io/vba/excel/code/component/management/2021/02/05/Programatically-updating-Excel-VBA-code.html
