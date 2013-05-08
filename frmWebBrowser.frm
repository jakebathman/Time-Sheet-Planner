VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmWebBrowser 
   Caption         =   "PeopleSoft Web Browser"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11700
   OleObjectBlob   =   "frmWebBrowser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmWebBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
Dim strURLtoPeopleSoft$

strURLtoPeopleSoft = "https://hrportal.co.collin.tx.us:15443/psp/ps/EMPLOYEE/HRMS/c/ROLE_EMPLOYEE.TL_MSS_EE_SRCH_PRD.GBL?PORTALPARAM_PTCNAV=HC_TL_SS_JOB_SRCH_EE_GBL&EOPP.SCNode=EMPL&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=PT_PTPP_PORTAL_ROOT&EOPP.SCLabel=Report Time&EOPP.SCFName=HC_RECORD_TIME&EOPP.SCSecondary=true&EOPP.SCPTfname=HC_RECORD_TIME&FolderPath=PORTAL_ROOT_OBJECT.CCG_HCM_MAIN.CO_EMPLOYEE_SELF_SERVICE.HC_TIME_REPORTING.HC_RECORD_TIME.HC_TL_SS_JOB_SRCH_EE_GBL&IsFolder=false"
    WebBrowser1.Navigate strURLtoPeopleSoft
End Sub



Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    frmWebBrowser.WebBrowser1
End Sub
