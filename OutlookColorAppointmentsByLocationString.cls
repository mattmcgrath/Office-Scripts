VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public WithEvents olItems As Outlook.Items
Attribute olItems.VB_VarHelpID = -1

Sub Application_Startup()
    Set olItems = Session.GetDefaultFolder(olFolderCalendar).Items
End Sub

Private Sub olItems_ItemAdd(ByVal Item As Object)
    Dim NewAppt As Outlook.AppointmentItem
 
    If Item.Class = olAppointment Then
       Set NewAppt = Item
    'Specify a color category based on appointment subject
    'You can replace "Test" with your desired string in subject line
        If InStr(Item.Location, "CR-") > 0 Then
            Item.Categories = AddCategory(Item.Categories, "In-Person Meeting")
            Item.Save
        End If
        If InStr(Item.Location, "Skype") > 0 Then
            Item.Categories = AddCategory(Item.Categories, "Phone Meeting")
            Item.Save
        End If
    End If

    Set NewAppt = Nothing
End Sub

Sub ChangeExistingAppointments()
    Dim oOL As New Outlook.Application
    Dim oNS As Outlook.NameSpace
    Dim oAppointments As Object
    Dim oAppointmentItem As Outlook.AppointmentItem

    Set oNS = oOL.GetNamespace("MAPI")
    Set oAppointments = oNS.GetDefaultFolder(olFolderCalendar)
    
    
    For Each oAppointmentItem In oAppointments.Items
        If InStr(oAppointmentItem.Location, "CR-") > 0 Then
            oAppointmentItem.Categories = AddCategory(oAppointmentItem.Categories, "In-Person Meeting")
            oAppointmentItem.Save
        End If
        If InStr(oAppointmentItem.Location, "Skype") > 0 Then
            oAppointmentItem.Categories = AddCategory(oAppointmentItem.Categories, "Phone Meeting")
            oAppointmentItem.Save
        End If
    Next

    Set oAppointmentItem = Nothing
    Set oAppointments = Nothing
    Set oNS = Nothing
    Set oOL = Nothing
    
End Sub

Function AddCategory(CurrentCategories As String, NewCategory As String) As String
If InStr(CurrentCategories, NewCategory) Then
    AddCategory = CurrentCategories
Else
    AddCategory = CurrentCategories & ", " & NewCategory
End If
End Function

Function RemoveCategory(CurrentCategories As String, OldCategory As String) As String
If InStr(CurrentCategories, OldCategory) Then
    RemoveCategory = Replace(CurrentCategories, ", " & OldCategory, "")
    RemoveCategory = Replace(CurrentCategories, OldCategory, "")
Else
    RemoveCategory = CurrentCategories
End If
End Function

