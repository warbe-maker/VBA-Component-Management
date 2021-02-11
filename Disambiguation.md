## Disambiguation
The terms below are not only those used in this post but also used with the implementation of the _Component Management_.

| Term             | Meaning                  |
|------------------|------------------------- |
|_Component_       | Generic _VB-Project_ term for a _Class Module_, a  _Data Module_, a _Standard Module_, or a _UserForm_  |
|_Common Component_| A _Component_ which is used by two or more VB-Projects |
| _Raw_,<br>_Raw-Component_ | The instance of a _Common Component_ which is regarded the developed, maintained and tested 'original', hosted in a dedicated _Raw-Host_ Workbook. |
| _Clone_,<br>_Clone-Component_,<br>_Raw-Clone_ | The copy of a _Raw- Component_ in a _VP-Project_ using it |
|_Clone-Project_ | A _VP-Project_ derived from a _Raw-Project_ |
|_Procedure_     | Any - Public or Private _Property_, _Sub_, or _Funtion_ of a _Component_. See also _Service_.
|_Raw-Host_.     | The Workbook/_VP-Project_ which hosts the _Raw-Component_ |
|_Raw-Project_   | A code-only _VP-Project_ of which all components are regarded _Raw-Components_. A _Raw-Project_ is kind of a template for the productive version of it. In contrast to a classic template it is the life-time raw code base for the productive _Clone-Project_.  The service and process of 'synchronizing' the productive (clone) code with the raw is part of the _Component Management_.|
|_Service_       | Generic term for any _Public Property_, _Public Sub_, or _Public Funtion_ of a _Component_ |
|_VB-Project_     | In the present case this term is used synonymously with Workbook |
| _Workbook-_, or<br>_VB-Project-Folder_ | A folder dedicated to a Workbook/VB-Project with all its Export Files and other project specific means. Such a folder is the equivalent of a Git-Repo-Clone (provided Git is used for the project's versioning which is recommendable |
