Attribute VB_Name = "ServerMain"
Option Explicit

Enum BOOKMODEL
  gintSCHOOL_BOOK_MODEL
  gintPOP_NOVEL_MODEL
  gintCELEBRITY_BIOGRAPHY_MODEL
End Enum

Global Const gRoyalty = 0.15
Global gintInstanceCount As Integer

Global gDB As Database
Global gWkspc As Workspace
Global gDBName As String

Global gintSalesModel As Integer
Global glngUnitsPerMonth As Long
Global gcurCostPerUnit As Currency
Global gcurPromoCost As Currency
Global gsngIncomeTaxRate As Single
Global gintSalesPeriod As Integer
Global gbolIsDiscount As Boolean
Global gintPromoStart As Integer
Global gintPromoDiscPeriod As Integer
Global gsngPromoDiscRate As Single
Global gintPubProfit As Integer
Global gcurAdvertisingCost As Currency

Sub Main()

End Sub

Sub ServerMsg(rstrMessage As String, rintButtons As Integer, rstrTitle As String)
  #If gbDEBUG Then
    MsgBox rstrMessage, rintButtons, rstrTitle
  #End If
End Sub

