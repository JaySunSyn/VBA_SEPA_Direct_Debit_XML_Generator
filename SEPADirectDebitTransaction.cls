VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SEPADirectDebitTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'
'/**
'   * Code allocated to a currency, by a maintenance agency, under an international identification
'   * scheme as described in the latest edition of the international standard ISO 4217.
'   */
Private varCURRENCY_CODE As String

' End Property
'   * Specifies which party/parties will bear the charges associated with the
'   * processing of the payment transaction. (Tag: <ChrgBr>)
'   *
'   * @var string
'   */
Private varCHARGE_BEARER As String

' End Property
'   * The instruction identification is a point to point reference that can be used between
'   * the instructing party and the instructed party to refer to the individual instruction.
'   * It can be included in several messages related to the instruction. (Tag: <DrctDbtTxInf->PmtId->InstrId>)
'   *
'   * @var string
'   */
Private varInstructionIdentification As String

' End Property
'   * The end-to-end identification can be used for reconciliation or to link tasks relating to the transaction.
'   * It can be included in several messages related to the transaction. (Tag: <DrctDbtTxInf->PmtId->EndToEndId>)
'   *
'   * @var string
'   */
Private varEndToEndIdentification As String

' End Property
'   * Amount of money to be moved between the debtor and creditor, before deduction of charges,
'   * expressed in the currency as ordered by the initiating party. (Tag: <DrctDbtTxInf->InstdAmt>)
'   *
'   * @var float
'   */
Private varInstructedAmount As Currency

' End Property
'   * Unique identification, as assigned by the creditor,
'   * to unambiguously identify the mandate. (Tag: <DrctDbtTxInf->DrctDbtTx->MndtRltdInf->MndtId>)
'   *
'   * @var string
'   */
Private varMandateIdentification As String

' End Property
'   * Date on which the direct debit mandate has been signed
'   * by the debtor. (Tag: <DrctDbtTxInf->DrctDbtTx->MndtRltdInf->DtOfSgntr>)
'   *
'   * @var string
'   */
Private varDateOfSignature As String

' End Property
'   * Indicator notifying whether the underlying mandate is amended or not.
'   * (Tag: <DrctDbtTxInf->DrctDbtTx->MndtRltdInf->AmdmntInd>)
'   *
'   * @var string
'   */
Private varAmendmentIndicator As String

' End Property
'   * Financial institution servicing an account for the debtor. (Tag: <DbtrAgt->FinInstnId->BIC>)
'   *
'   * @var string
'   */
Private varDebtorAgentBIC As String

' End Property
'   * Party that owes an amount of money to the (ultimate) creditor. (Tag: <Dbtr->Nm>)
'   *
'   * @var string
'   */
Private varDebtorName As String

' End Property
'   * Unambiguous identification of the account of the debtor to which a debit entry will
'   * be made as a result of the transaction. (Tag: <DbtrAcct->Id->IBAN>)
'   *
'   * @var string
'   */
Private varDebtorIBAN As String

' End Property
'   * Information supplied to enable the matching/reconciliation of an entry with the items that the
'   * payment is intended to settle, such as commercial invoices in an accounts' receivable system.
'   * (Tag: <RmtInf->Ustrd>)
'   *
'   * @var string
'   */
Private varRemittanceInformation As String

'Constructor
Private Sub Class_Initialize()

    'Constants
    varCURRENCY_CODE = "EUR"
    varCHARGE_BEARER = "SLEV"
    
    'Variable initialization
    varInstructionIdentification = ""
    varEndToEndIdentification = ""
    varInstructedAmount = 0#
    varMandateIdentification = ""
    varDateOfSignature = ""
    varAmendmentIndicator = "false"
    varDebtorAgentBIC = ""
    varDebtorName = ""
    varDebtorIBAN = ""
    varRemittanceInformation = ""

End Sub


'/**
'   * Getter for Instruction Identification (DrctDbtTxInf->PmtId->InstrId)
'   *
'   * @return string
'   */
Public Property Get InstructionIdentification() As String
    
    If varInstructionIdentification = "" Then
        InstructionIdentification = EndToEndIdentification
    Else
        InstructionIdentification = varInstructionIdentification
    End If

End Property


'   * Setter for Instruction Identification (DrctDbtTxInf->PmtId->InstrId)
'   *
'   * @param string $instrId
'   * @throws SEPAException
'   */
Public Property Let InstructionIdentification(ByVal instrId As String)
    
    '$instrId = URLify::downcode($instrId, "de");
     
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([\-A-Za-z0-9\+\/\?:\(\)\., ]){0,35}$"
    
    If re.Test(instrId) Then
        varInstructionIdentification = instrId
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid InstructionIdentification (max. 35).")
    
    End If

End Property

'   * Getter for EndToEnd Identification (DrctDbtTxInf->PmtId->EndToEndId)
'   *
'   * @return string
'   */
Public Property Get EndToEndIdentification() As String
  
  EndToEndIdentification = varEndToEndIdentification

End Property
   
'   * Setter for EndToEnd Identification (DrctDbtTxInf->PmtId->EndToEndId)
'   *
'   * @param string $endToEndId
'   * @throws SEPAException
'   */
Public Property Let EndToEndIdentification(ByVal endToEndId As String)
  
   '$endToEndId = URLify::downcode($endToEndId, "de");

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([\-A-Za-z0-9\+\/\?:\(\)\., ]){1,35}$"
    
    If re.Test(endToEndId) Then
        varEndToEndIdentification = endToEndId
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid EndToEndIdentification (max. 35).")
    
    End If

End Property
   
'   * Getter for Instructed Amount (DrctDbtTxInf->InstdAmt)
'   *
'   * @return float
'   */
Public Property Get InstructedAmount() As Currency
  
    InstructedAmount = varInstructedAmount
    
End Property

'   * Setter for Instructed Amount (DrctDbtTxInf->InstdAmt)
'   *
'   * @param float $instdAmt
'   */
Public Property Let InstructedAmount(ByVal instdAmt As Currency)
  
  varInstructedAmount = instdAmt

End Property

'   * Getter for Mandate Identification (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->MndtId)
'   *
'   * @return string
'   */
Public Property Get MandateIdentification() As String
  
  MandateIdentification = varMandateIdentification

End Property

'   * Setter for Mandate Identification (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->MndtId)
'   *
'   * @param string $mndtId
'   * @throws SEPAException
'   */
Public Property Let MandateIdentification(ByVal mndtId As String)
  
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([A-Za-z0-9]|[\+|\?|\/|\-|:|\(|\)|\.|,|']){1,35}$"
    
    If re.Test(mndtId) Then
        varMandateIdentification = mndtId
    Else
        Dim ex As New SEPAException
        ex.throw ("MndtId empty, contains invalid characters or too long (max. 35).")
        
    End If

End Property

'   * Getter for Mandate Identification (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->DtOfSgntr)
'   *
'   * @return string
'   */
Public Property Get DateOfSignature() As String
  
    DateOfSignature = varDateOfSignature

End Property
   
'   * Setter for Mandate Identification (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->DtOfSgntr)
'   *
'   * @param string $dtOfSgntr
'   */
Public Property Let DateOfSignature(ByVal dtOfSgntr As String)
  
    varDateOfSignature = dtOfSgntr

End Property

'   * Getter for Amendment Indicator (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->AmdmntInd)
'   *
'   * @return string
'   */
Public Property Get AmendmentIndicator() As String

    AmendmentIndicator = varAmendmentIndicator

End Property

'   * Setter for Amendment Indicator (DrctDbtTxInf->DrctDbtTx->MndtRltdInf->AmdmntInd)
'   *
'   * @param string $amdmntInd
'   */
Public Property Let AmendmentIndicator(ByVal amdmntInd As String)
  
    If amdmntInd = True Or amdmntInd = "true" Then
        varAmendmentIndicator = "true"
    Else
        varAmendmentIndicator = "false"
    End If

End Property
   
'   * Getter for Debtor Agent BIC (DbtrAgt->FinInstnId->BIC)
'   *
'   * @return string
'   */
Public Property Get DebtorAgentBIC() As String

    DebtorAgentBIC = varDebtorAgentBIC

End Property

'   * Setter for Debtor Agent BIC (DbtrAgt->FinInstnId->BIC)
'   *
'   * @param string $bic
'   * @throws SEPAException
'   */
Public Property Let DebtorAgentBIC(ByVal bic As String)
  
    bic = Replace(bic, " ", "")
  
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[0-9a-z]{4}[a-z]{2}[0-9a-z]{2}([0-9a-z]{3})?$"
    
    If re.Test(bic) Then
        varDebtorAgentBIC = bic
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid debtor BIC.")
        
    End If

End Property

'   * Getter for Debtor Name (Dbtr->Nm)
'   *
'   * @return string
'   */
Public Property Get DebtorName() As String

    DebtorName = varDebtorName

End Property

'   * Setter for Debtor Name (Dbtr->Nm)
'   *
'   * @param string $dbtr
'   * @throws SEPAException
'   */
Public Property Let DebtorName(ByVal dbtr As String)

    '$dbtr = URLify::downcode($dbtr, "de");

    If Len(dbtr) = 0 Or Len(dbtr) > 70 Then
        Dim ex As New SEPAException
        ex.throw ("Invalid debtor name (max. 70 characters).")
    Else
        varDebtorName = dbtr
    End If

End Property

'   * Getter for Debtor IBAN (DbtrAcct->Id->IBAN)
'   *
'   * @return string
'   */
Public Property Get DebtorIBAN() As String

    DebtorIBAN = varDebtorIBAN

End Property
   
'   * Setter for Debtor IBAN (DbtrAcct->Id->IBAN)
'   *
'   * @param string $iban
'   * @throws SEPAException
'   */
Public Property Let DebtorIBAN(ByVal iban As String)
    
    iban = Replace(iban, " ", "")
    
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[a-zA-Z]{2}[0-9]{2}[a-zA-Z0-9]{4}[0-9]{7}([a-zA-Z0-9]?){0,16}$"
    
    If re.Test(iban) Then
        varDebtorIBAN = iban
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid debtor IBAN.")
        
    End If

End Property

'   * Getter for Remittance Information (RmtInf->Ustrd)
'   *
'   * @return string
'   */
Public Property Get RemittanceInformation() As String

    RemittanceInformation = varRemittanceInformation
    
End Property

'   * Getter for Remittance Information (RmtInf->Ustrd)
'   *
'   * @param string $ustrd
'   * @throws SEPAException
'   */
Public Property Let RemittanceInformation(ByVal ustrd As String)

    '$ustrd = URLify::downcode($ustrd, "de");
    
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([A-Za-z0-9]|[\+|\?|\/|\-|:|\(|\)|\.|,|'| ]){0,140}$"
    
    If re.Test(ustrd) Then
        varRemittanceInformation = ustrd
    Else
        Dim ex As New SEPAException
        ex.throw ("RmtInf contains invalid chars or is too long (max. 140).")
        
    End If

 End Property
 
'   * Returns a SimpleXMLElement for the SEPADirectDebitTransaction object.
'   *
'   * @return SimpleXMLElement object
'   */
Public Property Get XmlDirectDebitTransaction() As DOMDocument
      
    Dim objXML As DOMDocument
    
    Dim objXMLRootelement As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLelementSub As IXMLDOMElement
    Dim objXMLelementSubSub As IXMLDOMElement
    Dim objXMLelementSubSubSub As IXMLDOMElement
    Dim objXMLelementSubSubSubSub As IXMLDOMElement
    
    Set objXML = New DOMDocument
    
    Set objXMLRootelement = objXML.createElement("DrctDbtTxInf")
    objXML.appendChild objXMLRootelement
    
    Set objXMLelement = objXML.createElement("PmtId")
    Set objXMLelementSub = objXML.createElement("InstrId")
    objXMLelementSub.Text = InstructionIdentification
    '
    Set objXMLelementSubSub = objXML.createElement("EndToEndId")
    objXMLelementSubSub.Text = EndToEndIdentification
    '
    objXMLelement.appendChild objXMLelementSub
    objXMLelement.appendChild objXMLelementSubSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("InstdAmt")
    objXMLelement.Text = InstructedAmount
    objXMLelement.setAttribute "Ccy", varCURRENCY_CODE
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("ChrgBr")
    objXMLelement.Text = varCHARGE_BEARER
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("DrctDbtTx")
    Set objXMLelementSub = objXML.createElement("MndtRltdInf")
    Set objXMLelementSubSub = objXML.createElement("MndtId")
    objXMLelementSubSub.Text = MandateIdentification
    '
    Set objXMLelementSubSubSub = objXML.createElement("DtOfSgntr")
    objXMLelementSubSubSub.Text = DateOfSignature
    '
    Set objXMLelementSubSubSubSub = objXML.createElement("AmdmntInd")
    objXMLelementSubSubSubSub.Text = AmendmentIndicator
    '
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelementSub.appendChild objXMLelementSubSubSub
    objXMLelementSub.appendChild objXMLelementSubSubSubSub
    '
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("DbtrAgt")
    Set objXMLelementSub = objXML.createElement("FinInstnId")
    Set objXMLelementSubSub = objXML.createElement("BIC")
    objXMLelementSubSub.Text = DebtorAgentBIC
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("Dbtr")
    Set objXMLelementSub = objXML.createElement("Nm")
    objXMLelementSub.Text = DebtorName
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("DbtrAcct")
    Set objXMLelementSub = objXML.createElement("Id")
    Set objXMLelementSubSub = objXML.createElement("IBAN")
    objXMLelementSubSub.Text = DebtorIBAN
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("RmtInf")
    Set objXMLelementSub = objXML.createElement("Ustrd")
    objXMLelementSub.Text = RemittanceInformation
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    'objXML.save ("\\192.168.3.6\ablage_tmp$\testDT.xml")
    
    Set XmlDirectDebitTransaction = objXML

End Property
