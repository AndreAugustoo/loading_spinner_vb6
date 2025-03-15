Attribute VB_Name = "ModuloLoadingSpinner"
Public Sub IsLoading(State As Boolean)

   If State Then
      frm_menu_principal.Enabled = False
      Load frm_loading_spinner
      frm_loading_spinner.Show vbModeless
      DoEvents
   Else
      frm_menu_principal.Enabled = True
      Unload frm_loading_spinner
      frm_loading_spinner.Hide
      DoEvents
   End If
   
End Sub
