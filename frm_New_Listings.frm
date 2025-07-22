VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_New_Listings 
   Caption         =   "Main Form"
   ClientHeight    =   19920
   ClientLeft      =   -18900
   ClientTop       =   -76020.01
   ClientWidth     =   17380
   OleObjectBlob   =   "frm_New_Listings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_New_Listings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

   
   






Private Sub cbAcoustic_Click()
    
    If ActiveSheet.Name <> "Acoustic" Then
        Worksheets("Acoustic").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If
End Sub

Private Sub cbAguilar_Click()

    If ActiveSheet.Name <> "Aguilar" Then
        Worksheets("Aguilar").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If

End Sub

Private Sub cbAkai_Click()

    If ActiveSheet.Name <> "AKAI" Then
        Worksheets("AKAI").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If

End Sub

Private Sub cbAlesis_Click()
    
    If ActiveSheet.Name <> "Alesis" Then
        Worksheets("Alesis").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If
    
End Sub

Private Sub cbAmazonListingSheet_Click()

        Worksheets("Amazon").Activate
        frm_Listings.Hide

End Sub

Private Sub cbAmpeg_Click()
    
    If ActiveSheet.Name <> "Ampeg" Then
        Worksheets("Ampeg").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If
    
End Sub

Private Sub cbArturia_Click()

    If ActiveSheet.Name <> "Arturia" Then
        Worksheets("Arturia").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If
    
End Sub

Private Sub cbBehringer_Click()

    If ActiveSheet.Name <> "Behringer" Then
        Worksheets("Behringer").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If
    
End Sub

Private Sub cbBlackstar_Click()

    If ActiveSheet.Name <> "Blackstar" Then
        Worksheets("Blackstar").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If


End Sub

Private Sub cbBogner_Click()

    If ActiveSheet.Name <> "Bogner" Then
        Worksheets("Bogner").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If


    
End Sub

Private Sub cbBoss_Click()

    If ActiveSheet.Name <> "Boss" Then
        Worksheets("Boss").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If


    
End Sub

Private Sub cbBugera_Click()

    If ActiveSheet.Name <> "Bugera" Then
        Worksheets("Bugera").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbDexibell_Click()

    If ActiveSheet.Name <> "Dexibell" Then
        Worksheets("Dexibell").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If


    
End Sub

Private Sub cbDrZ_Click()

    If ActiveSheet.Name <> "Dr-Z" Then
        Worksheets("Dr-Z").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbDVMark_Click()

    If ActiveSheet.Name <> "DVMark" Then
        Worksheets("DVMark").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbEden_Click()

    If ActiveSheet.Name <> "Eden" Then
        Worksheets("Eden").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If


    
End Sub

Private Sub cbElliott_Click()

    If ActiveSheet.Name <> "Elliott" Then
        Worksheets("Elliott").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
    
End Sub

Private Sub cbEminence_Click()

    If ActiveSheet.Name <> "Eminence" Then
        Worksheets("Eminence").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbEngl_Click()

    If ActiveSheet.Name <> "Engl" Then
        Worksheets("Engl").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbFender_Click()

    If ActiveSheet.Name <> "Fender" Then
        Worksheets("Fender").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbFishman_Click()

    If ActiveSheet.Name <> "Fishman" Then
        Worksheets("Fishman").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbFriedman_Click()

    If ActiveSheet.Name <> "Friedman" Then
        Worksheets("Friedman").Activate
        frm_Listings.Hide
    Else
        Worksheets("Input").Activate
    End If



    
End Sub

Private Sub cbGRBass_Click()
    Worksheets("GR-Bass").Activate
    frm_Listings.Hide
End Sub

Private Sub cbHammond_Click()
    Worksheets("Hammond").Activate
    frm_Listings.Hide
End Sub

Private Sub cbHarmony_Click()
    Worksheets("Harmony").Activate
    frm_Listings.Hide
End Sub

Private Sub cbHartke_Click()
    Worksheets("Hartke").Activate
    frm_Listings.Hide
End Sub

Private Sub cbHughesKettner_Click()
    Worksheets("Hughes-Kettner").Activate
    frm_Listings.Hide
End Sub

Private Sub cbInsertNewManufacturer_Click()
    modMaintenanceProcedures.insertManufacturerName
End Sub

Private Sub cbKawai_Click()
    Worksheets("Kawai").Activate
    frm_Listings.Hide
End Sub

Private Sub cbKorg_Click()
    Worksheets("Korg").Activate
    frm_Listings.Hide
End Sub

Private Sub cbKurzweil_Click()
    Worksheets("Kurzweil").Activate
    frm_Listings.Hide
End Sub

Private Sub cbLegend_Click()
    Worksheets("Legend").Activate
    frm_Listings.Hide
End Sub

Private Sub cbLine6_Click()
    Worksheets("Line6").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMackie_Click()
    Worksheets("Mackie").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMagnatone_Click()
    Worksheets("Magnatone").Activate
    frm_Listings.Hide
End Sub

Private Sub cbManufacturersName_Click()
    frm_Listings.Hide
    frmUpdateManufacturerNameList.Show
    Worksheets("ManufacturerNames").Activate
    
End Sub

Private Sub cbMarshall_Click()
    Worksheets("Marshall").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMAudio_Click()
    Worksheets("MAudio").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMemphisBlues_Click()
    Worksheets("Memphis-Blues").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMesaBoogie_Click()
    Worksheets("Mesa-Boogie").Activate
    frm_Listings.Hide
End Sub

Private Sub cbMonoprice_Click()
    Worksheets("Monoprice").Activate
    frm_Listings.Hide
End Sub

Private Sub cbNektar_Click()
    Worksheets("Nektar").Activate
    frm_Listings.Hide
End Sub

Private Sub cbNord_Click()
    Worksheets("Nord").Activate
    frm_Listings.Hide
End Sub

Private Sub cbNovation_Click()
    Worksheets("Novation").Activate
    frm_Listings.Hide
End Sub

Private Sub cbOrange_Click()
    Worksheets("Orange").Activate
    frm_Listings.Hide
End Sub

Private Sub cboRetrievePreviousListing_Click()
'    strSheetName = frm_Listings.cboManufacturerNames.Value
'    strBrandName = frm_Listings.cboManufacturerNames.Value
'    strSeriesName = frm_Listings.cboSeriesName.Value
    modMaintenanceProcedures.clearOmpPages
    modEnterModelInformation.retrieveModelInformation

       ' modEnterModelInformation.retrieveEquipmentTypeListingInformation
    'modQuickProcedures.colorCodeInputRows
        'mod_perform_Calculations.calculateMeasurementMaterials

    
End Sub



Private Sub cboUpdateLinks_Change()
    modMaintenanceProcedures.updateTemplate
End Sub


Private Sub cbPaulReedSmith_Click()
    modOpenFormsNSheets.openPaulReedSmithSheet
End Sub

Private Sub cbPeavey_Click()
    Worksheets("Peavey").Activate
    frm_Listings.Hide
End Sub

Private Sub cbPositiveGrid_Click()
    Worksheets("Positive-Grid").Activate
    frm_Listings.Hide
End Sub

Private Sub cbPreSonus_Click()
    Worksheets("PreSonus").Activate
    frm_Listings.Hide
End Sub

Private Sub cbProcessAllSeries_Click()
    boo_All_Series = True
    mod_Listings_Create.process_Listings
End Sub

Private Sub cbRandall_Click()
    Worksheets("Randall").Activate
    frm_Listings.Hide
End Sub

Private Sub cbRoland_Click()
    Worksheets("Roland").Activate
    frm_Listings.Hide
End Sub

Private Sub cbRoli_Click()
    Worksheets("Roli").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSeismic_Click()
    Worksheets("Seismic").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSequential_Click()
    Worksheets("Sequential").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSilktone_Click()
    Worksheets("Silktone").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSoldano_Click()
    Worksheets("Soldano").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSoundCraft_Click()
    Worksheets("Soundcraft").Activate
    frm_Listings.Hide
End Sub

Private Sub cbStudioLogic_Click()
    Worksheets("Studiologic").Activate
    frm_Listings.Hide
End Sub

Private Sub cbSupro_Click()
    Worksheets("Supro").Activate
    frm_Listings.Hide
End Sub

Private Sub cbTraceElliot_Click()
    Worksheets("Trace-Elliot").Activate
    frm_Listings.Hide
End Sub

Private Sub cbVictoria_Click()
    Worksheets("Victoria").Activate
    frm_Listings.Hide
End Sub

Private Sub cbVox_Click()
    Worksheets("Vox").Activate
    frm_Listings.Hide
End Sub

Private Sub cbWaldorf_Click()
    Worksheets("Waldorf").Activate
    frm_Listings.Hide
End Sub

Private Sub cbWarwick_Click()
    Worksheets("Warwick").Activate
    frm_Listings.Hide
End Sub

Private Sub cbWilliams_Click()
    Worksheets("Williams").Activate
    frm_Listings.Hide
End Sub

Private Sub cbx_Free_Shipping_Amazon_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbx_Free_Shipping_eBay_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbx_Free_Shipping_Reverb_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbx_Paid_Shipping_Amazon_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbx_Paid_Shipping_eBay_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbx_Paid_Shipping_Reverb_Click()
    get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub cbYamaha_Click()
    Worksheets("Yamaha").Activate
    frm_Listings.Hide
End Sub






Private Sub chk_Hidden_Dictionaries_Click()
    Call mod_Form_Load_Controls.handle_chk_Hidden_Dictionaries
End Sub

Private Sub lst_Manufacturer_Names_AfterUpdate()
   ' Call mod_Form_Load_Controls.On_Manufacturer_User_Selection
End Sub

Private Sub lst_Manufacturer_Names_Click()
    
    Call On_Manufacturer_User_Selection(frm_New_Listings)


End Sub





Private Sub lst_Series_Name_Click()
Call mod_Form_Load_Controls.On_Series_Name_User_Selection(Me)
 

End Sub
Private Sub tbEquipmentType_Change()

End Sub

Private Sub MultiPage1_Change()

End Sub


Private Sub opt_Free_Shipping_Reverb_Click()
    boo_Reverb_Shipping_Profile = True
    'get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub opt_Paid_Shipping_Amazon_Click()
    boo_Amazon_Shipping_Profile = False
   'get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub opt_Paid_Shipping_eBay_Click()
    boo_eBay_Shipping_Profile = False
   ' get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub opt_Paid_Shipping_Reverb_Click()
    boo_Reverb_Shipping_Profile = False
    'get_Set_Pricing_Variables.free_Or_Paid_Shipping
End Sub

Private Sub opt1_Current_Pricing_Click()
    mod_Get_Variables.get_Pricing_Options
End Sub

Private Sub opt2_Previous_Pricing_Click()
    mod_Get_Variables.get_Pricing_Options
End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Activate()


'
'boo_Listings_Initializing = True
'
'Dim ws_Row_Source As Worksheet
'
'
'Set ws_Row_Source = Worksheets("Range_Lists")
'
'
'    Me.lst_Brand_Names.List = ws_Row_Source.Range("A1:A67").Value
'
'boo_Listings_Initializing = False

End Sub

Private Sub UserForm_Initialize()

    Set obj_Form = Me
    Call mod_Form_Load_Controls.init_form_New_Listings
    
End Sub
Sub saveFileTemplates()
    
    Dim myCSVFileName As String
    Dim myUnicodeFileName As String
    Dim tempWB As Workbook
    Dim strSaveLine As String
    brandName = ThisWorkbook.Sheets("input").Range("A2")
    series = ThisWorkbook.Sheets("input").Range("A4")
    strSaveLine = ThisWorkbook.Sheets("input").Range("A6")
    

   Dim myFileName As String
   Dim mySavePath As Variant
   
    
    Dim saveDirectory As String
    
    
    amazon.saveAmazon
    woo.saveWooCsv
    eBay.saveEbayCsv
    
'    myYear = Year(Now)
'    myMonth = Month(Now)
'    myDay = Day(Now)
    
'    Select Case strSaveLine
'        Case Is = "Printer Cover"
'            mySavePath = "C:\Users\SMartz\Dust Covers For You\Dust-Covers-For-You - Documents\Listings\" & strSaveLine & "\" & brandName & "\" & series & "\" & myYear & "-" & myMonth & "-" & myDay & "-" & brandName & "-" & series & ".xlsm"
'        Case Is = "Music Keyboard Cover"
'            mySavePath = "C:\Users\SMartz\Dust Covers For You\" & strSaveLine '& "\" & brandName & "\" & series & "\" & myYear & "-" & myMonth & "-" & myDay
'
'    End Select
'
'   ActiveWorkbook.SaveAs Filename:=mySavePath


End Sub
Sub uploadToEbay()
    
    Dim ebayURL As String
    ebayURL = "https://www.ebay.com/sh/reports/"
    ActiveWorkbook.FollowHyperlink ebayURL

End Sub

Sub uploadToWoo()

    Dim wooCommerceURL As String
    wooCommerceURL = "https://dustcoversforyou.store/wp-admin/edit.php?post_type=product"
    ActiveWorkbook.FollowHyperlink wooCommerceURL

End Sub

Sub uploadToAmazon()

    Dim amazonURL As String
    amazonURL = "https://sellercentral.amazon.com/listing/upload?ref_=xx_upload_tnav_status"
    ActiveWorkbook.FollowHyperlink amazonURL
    
    Dim amazonOptionsURL As String
    amazonOptionsURL = "https://sellercentral.amazon.com/gestalt/sellertemplate/index.html"
    ActiveWorkbook.FollowHyperlink amazonOptionsURL

End Sub

Sub uploadToMarketPlaces()

    Dim ebayURL As String
    ebayURL = "https://www.ebay.com/sh/reports/"
    ActiveWorkbook.FollowHyperlink ebayURL
    
    Dim wooCommerceURL As String
    wooCommerceURL = "https://dustcoversforyou.store/wp-admin/edit.php?post_type=product"
    ActiveWorkbook.FollowHyperlink wooCommerceURL
    
    Dim amazonURL As String
    amazonURL = "https://sellercentral.amazon.com/listing/upload?ref_=xx_upload_tnav_status"
    ActiveWorkbook.FollowHyperlink amazonURL
    
    Dim amazonOptionsURL As String
    amazonOptionsURL = "https://sellercentral.amazon.com/gestalt/sellertemplate/index.html"
    ActiveWorkbook.FollowHyperlink amazonOptionsURL

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

   Unload Me
End Sub
