Attribute VB_Name = "mAppData"
Option Explicit
' ---------------------------------------------------------------------
' Standard Module mAppData
'          Declarations for the use of the Class Module clsAppData.
' ---------------------------------------------------------------------
Public Enum enRegDataType
    spp_REG_NONE = 0                    ' No defined value type.
    spp_REG_SZ = 1                      ' A null-terminated string. This will be either a Unicode or an ANSI string,
                                        ' depending on whether you use the Unicode or ANSI functions.
    spp_REG_EXPAND_SZ = 2               ' A null-terminated string that contains unexpanded references to environment variables
                                        '(for example, "%PATH%"). It will be a Unicode or ANSI string depending on whether
                                        ' you use the Unicode or ANSI functions. To expand the environment variable references,
                                        ' use the ExpandEnvironmentStrings function.
    spp_REG_BINARY = 3                  ' Binary data in any form.
    spp_REG_DWORD = 4                   ' A 32-bit number.
    spp_REG_DWORD_LITTLE_ENDIAN = 4     ' A 32-bit number in little-endian format.
                                        ' Windows is designed to run on little-endian computer architectures.
                                        ' Therefore, this value is defined as REG_DWORD in the Windows header files.
    spp_REG_DWORD_BIG_ENDIAN = 5        ' A 32-bit number in big-endian format. Some UNIX systems support big-endian architectures.
    spp_REG_LINK = 6                    ' A null-terminated Unicode string that contains the target path of a symbolic link that
                                        ' was created by calling the RegCreateKeyEx function with REG_OPTION_CREATE_LINK.
    spp_REG_MULTI_SZ = 7                ' A sequence of null-terminated strings, terminated by an empty string (\0).
                                        ' The following is an example:
                                        ' String1\0String2\0String3\0LastString\0\0
                                        ' The first \0 terminates the first string, the second to the last \0
                                        ' terminates the last string, and the final \0 terminates the sequence.
                                        ' Note that the final terminator must be factored into the length of the string.
    spp_REG_RESOURCE_LIST = 8
    spp_FullResourceDescriptor = 9      ' Resource list in the hardware description
    spp_ResourceRequirementsList = 10
    spp_REG_QWORD = 11                  ' A 64-bit number.
    spp_REG_QWORD_LITTLE_ENDIAN = 11    ' A 64-bit number in little-endian format.
                                        ' Windows is designed to run on little-endian computer architectures.
                                        ' Therefore, this value is defined as REG_QWORD in the Windows header files.
End Enum

Public Enum enLocation                  ' Enumerated locations (saves error handling by enforcing valid values)
    spp_File = 1                        ' "Private Profile" file
    spp_Registry = 2
End Enum

Public Enum enHKey                      ' Enumerated HKey (saves error handling by enforcing a valid values)
    spp_HKEY_CLASSES_ROOT = 1
    spp_HKEY_CURRENT_CONFIG = 2
    spp_HKEY_CURRENT_USER = 3
    spp_HKEY_LOCAL_MACHINE = 4
    spp_HKEY_USERS = 5
End Enum

Public Enum enItem                      ' Enumerated items of a session persistent property
    spp_Subjct
    spp_Aspect
    spp_VaName
End Enum

Public Const REG_MAX_VALUE_LENGTH   As Long = &H100
Public Const COMMON_                As String = "Common"
