VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgressBar 
   Caption         =   "UserForm1"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   OleObjectBlob   =   "frmProgressBar.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub UserForm_Initialize()
    WebBrowser1.AddressBar = True
    WebBrowser1.Navigate "https://employees.co.collin.tx.us/psp/EMPSS/EMPLOYEE/HRMS/c/ROLE_EMPLOYEE.TL_MSS_EE_SRCH_PRD.GBL?PORTALPARAM_PTCNAV=HC_TL_SS_JOB_SRCH_EE_GBL&EOPP.SCNode=HRMS&EOPP.SCPortal=EMPLOYEE&EOPP.SCName=CO_EMPLOYEE_SELF_SERVICE&EOPP.SCLabel=Report Time&EOPP.SCFName=HC_RECORD_TIME&EOPP.SCSecondary=true&EOPP.SCPTfname=HC_RECORD_TIME&FolderPath=PORTAL_ROOT_OBJECT.CO_EMPLOYEE_SELF_SERVICE.HC_TIME_REPORTING.HC_RECORD_TIME.HC_TL_SS_JOB_SRCH_EE_GBL&IsFolder=false"
    WebBrowser1.GetProperty


End Sub
