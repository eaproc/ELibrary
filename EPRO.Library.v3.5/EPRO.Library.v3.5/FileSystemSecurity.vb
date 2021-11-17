Imports System.IO
Imports System.Security

Imports EPRO.Library.v3._5.MicrosoftOS

Public Class FileSystemSecurity


    ''' <summary>
    ''' Using Installed OSCulture
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFolderPath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveAccountFromFolder(pFolderPath As String,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        RemoveAccountFromFolder(pFolderPath,
                           New PCAccountNames(
                               OSCulture.getOSCultureInstalledOnPC()
                               ), pAccount
                           )

    End Sub

    ''' <summary>
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFolderPath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveAccountFromFolder(pFolderPath As String,
                                          PCAccountHandler As PCAccountNames,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        Dim di As New DirectoryInfo(pFolderPath)
        Dim AccessRule As New AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount),
                                                                      AccessControl.FileSystemRights.FullControl,
                                                                      AccessControl.AccessControlType.Allow
                                                                      )

        Dim dSecurity As AccessControl.DirectorySecurity = di.GetAccessControl()
        dSecurity.RemoveAccessRule(AccessRule)
        di.SetAccessControl(dSecurity)

    End Sub



    ''' <summary>
    ''' Using Installed OSCulture
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFolderPath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub AddAccountToFolder(pFolderPath As String,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        AddAccountToFolder(pFolderPath,
                          New PCAccountNames(
                              OSCulture.getOSCultureInstalledOnPC()
                              ), pAccount
                          )
    End Sub

    ''' <summary>
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFolderPath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub AddAccountToFolder(pFolderPath As String,
                                          PCAccountHandler As PCAccountNames,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        Dim di As New DirectoryInfo(pFolderPath)
        Dim AccessRule As New AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount),
                                                                      AccessControl.FileSystemRights.FullControl,
                                                                      AccessControl.AccessControlType.Allow
                                                                      )

        Dim dSecurity As AccessControl.DirectorySecurity = di.GetAccessControl()
        dSecurity.AddAccessRule(AccessRule)
        di.SetAccessControl(dSecurity)

    End Sub



    ''' <summary>
    ''' Using Installed OSCulture
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub AddAccountToFile(pFilePath As String,
                                         pAccount As PCAccountNames.AccountNames
                                         )

        AddAccountToFile(pFilePath,
                         New PCAccountNames(
                             OSCulture.getOSCultureInstalledOnPC()
                             ), pAccount
                         )

    End Sub


    ''' <summary>
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub AddAccountToFile(pFilePath As String,
                                         PCAccountHandler As PCAccountNames,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        Dim di As New FileInfo(pFilePath)
        Dim AccessRule As New AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount),
                                                                        AccessControl.FileSystemRights.FullControl,
                                                                        AccessControl.AccessControlType.Allow
                                                                        )

        Dim fi As New FileInfo(pFilePath)
        Dim fSecurity As AccessControl.FileSecurity = fi.GetAccessControl()
        fSecurity.AddAccessRule(AccessRule)
        fi.SetAccessControl(fSecurity)

    End Sub


    ''' <summary>
    ''' Using Installed OSCulture
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveAccountFromFile(pFilePath As String,
                                         pAccount As PCAccountNames.AccountNames
                                         )

        RemoveAccountFromFile(pFilePath,
                         New PCAccountNames(
                             OSCulture.getOSCultureInstalledOnPC()
                             ), pAccount
                         )

    End Sub

    ''' <summary>
    ''' NB: This does not handle errors like File or Folder NOT Found
    ''' </summary>
    ''' <param name="pFilePath"></param>
    ''' <param name="pAccount"></param>
    ''' <remarks></remarks>
    Public Shared Sub RemoveAccountFromFile(pFilePath As String,
                                         PCAccountHandler As PCAccountNames,
                                         pAccount As PCAccountNames.AccountNames
                                         )
        Dim di As New FileInfo(pFilePath)
        Dim AccessRule As New AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount),
                                                                        AccessControl.FileSystemRights.FullControl,
                                                                        AccessControl.AccessControlType.Allow
                                                                        )

        Dim fi As New FileInfo(pFilePath)
        Dim fSecurity As AccessControl.FileSecurity = fi.GetAccessControl()
        fSecurity.RemoveAccessRule(AccessRule)
        fi.SetAccessControl(fSecurity)

    End Sub



End Class
