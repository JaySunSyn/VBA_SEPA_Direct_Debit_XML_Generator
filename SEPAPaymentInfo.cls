VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SEPAPaymentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'**
'   * Specifies the means of payment that will be used to move the amount of money. (Tag: <PmtMtd>)
'   *
'   * Collection of an amount of money from the debtor's bank account by the creditor.
'   * The amount of money and dates of collections may vary.
'   *
'   * @var string
'   */
Private varPAYMENT_METHOD As String

'  /**
'   * Specifies a pre-agreed service or level of service between the parties,
'   * as published in an external service level code list. (Tag: <SvcLvl->Cd>)
'   *
'   * @var string
'   */
Private varSERVICELEVEL_CODE As String

'  /**
'   * Specifies which party/parties will bear the charges associated with the
'   * processing of the payment transaction. (Tag: <ChrgBr>)
'   *
'   * @var string
'   */
Private varCHARGE_BEARER As String

'  /**
'   * Name of the identification scheme, in a free text form. (Tag: <CdtrSchmeId->Id->PrvtId->Othr->SchmeNm->Prtry>)
'   *
'   * @var string
'   */
Private varPROPRIETARY_NAME As String

'  /**
'   * Unique identification, as assigned by a sending party, to unambiguously identify the
'   * payment information group within the message. (Tag: <PmtInfId>)
'   *
'   * @var string
'   */
Private varPaymentInformationIdentification As String

'  /**
'   * Identifies whether a single entry per individual transaction or a batch entry for the
'   * sum of the amounts of all transactions within the group of a message is requested (Tag: <BtchBookg>)
'   *
'   * @var string
'   */
Private varBatchBooking As String

'  /**
'   * Number of individual transactions contained in the payment information group (Tag: <NbOfTxs>)
'   *
'   * @var int
'   */
Private varNumberOfTransactions As Integer

'  /**
'   * Total of all individual amounts included in the group (Tag: <CtrlSum>)
'   *
'   * @var float
'   */
Private varControlSum As Currency

'  /**
'   * Specifies the local instrument, as published in an external local instrument code list.
'   * Only 'CORE', 'COR1', or 'B2B' are allowed. (Tag: <LclInstrm->Cd>)
'   *
'   * @var string
'   */
Private varLocalInstrumentCode As String

'  /**
'   * Identifies the direct debit sequence, such as first, recurrent, final or one-off. (Tag: <SeqTp>)
'   *
'   * @var string
'   */
Private varSequenceType As String

'  /**
'   * Date and time at which the creditor requests that the amount of money is to be
'   * collected from the debtor. (Tag: <ReqdColltnDt>)
'   *
'   * @var DateTime object
'   */
Private varRequestedCollectionDate As String

'  /**
'   * Party to which an amount of money is due. (Tag: <Cdtr->Nm>)
'   *
'   * @var string
'   */
Private varCreditorName As String

'  /**
'   * Unambiguous identification of the account of the creditor to which a credit entry will
'   * be posted as a result of the payment transaction. (Tag: <CdtrAcct->Id->IBAN>)
'   *
'   * @var string
'   */
Private varCreditorAccountIBAN As String

'  /**
'   * Financial institution servicing an account for the creditor. (Tag: <CdtrAgt->FinInstnId->BIC>)
'   *
'   * @var string
'   */
Private varCreditorAgentBIC As String

'  /**
'   * Credit party that signs the mandate. (Tag: <CdtrSchmeId->Id->PrvtId->Othr->Id>)
'   *
'   * @var string
'   */
Private varCreditorSchemeIdentification As String

'  /**
'   * A container for all transactions within this PaymentInfo.
'   *
'   * @var SEPADirectDebitTransaction[] array
'   */
Private varTransactions() As SEPADirectDebitTransaction

'Constructor
Private Sub Class_Initialize()
    
    'Constants initialization
    varPAYMENT_METHOD = "DD"
    varSERVICELEVEL_CODE = "SEPA"
    varCHARGE_BEARER = "SLEV"
    varPROPRIETARY_NAME = "SEPA"
    
    'Variable initialization
    varPaymentInformationIdentification = ""
    varBatchBooking = ""
    varNumberOfTransactions = 0
    varControlSum = 0#
    varLocalInstrumentCode = "CORE"
    varSequenceType = "FIRST"
    varRequestedCollectionDate = ""
    varCreditorName = ""
    varCreditorAccountIBAN = ""
    varCreditorAgentBIC = ""
    varCreditorSchemeIdentification = ""
    ReDim varTransactions(0 To 0)
    
    
End Sub

' PAYMENT_METHOD read access
Public Property Get PAYMENT_METHOD() As String
    PAYMENT_METHOD = varPAYMENT_METHOD
End Property

' SERVICELEVEL_CODE read access
Public Property Get SERVICELEVEL_CODE() As String
    SERVICELEVEL_CODE = varSERVICELEVEL_CODE
End Property

' CHARGE_BEARER read access
Public Property Get CHARGE_BEARER() As String
    CHARGE_BEARER = varCHARGE_BEARER
End Property

' PROPRIETARY_NAME read access
Public Property Get PROPRIETARY_NAME() As String
    PROPRIETARY_NAME = varPROPRIETARY_NAME
End Property

' PaymentInformationIdentification read access
Public Property Get PaymentInformationIdentification() As String
    PaymentInformationIdentification = varPaymentInformationIdentification
End Property

' PaymentInformationIdentification write access
Public Property Let PaymentInformationIdentification(ByVal payinf As String)
    
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([A-Za-z0-9]|[\+|\?|\/|\-|:|\(|\)|\.|,|'| ]){1,35}$"
    
    If re.Test(payinf) Then
        varPaymentInformationIdentification = payinf
    Else
        Dim ex As New SEPAException
        ex.throw ("PmtInfId empty, contains invalid characters or too long (max. 35).")
    
    End If
    
End Property

' BatchBooking read access
Public Property Get BatchBooking() As String
    BatchBooking = varBatchBooking
End Property

' BatchBooking write access
Public Property Let BatchBooking(ByVal bbook As String)
    
    If bbook = "true" Or bbook = True Then
        varBatchBooking = "true"
    Else
        varBatchBooking = "false"
    End If
    
End Property

' NumberOfTransactions read access
Public Property Get NumberOfTransactions() As Integer
    NumberOfTransactions = varNumberOfTransactions
End Property

' NumberOfTransactions write access
Public Property Let NumberOfTransactions(ByVal numTrans As Integer)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[0-9]{1,3}$"
    
    If re.Test(numTrans) Then
        varNumberOfTransactions = numTrans
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid NbOfTxs value (max. 15 digits).")
    End If
        
End Property

' ControlSum read access
Public Property Get ControlSum() As Currency
    ControlSum = varControlSum
End Property

' ControlSum write access
Public Property Let ControlSum(ByVal cosum As Currency)
    varControlSum = cosum
End Property

' LocalInstrumentCode read access
Public Property Get LocalInstrumentCode() As String
    LocalInstrumentCode = varLocalInstrumentCode
End Property

' LocalInstrumentCode write access
Public Property Let LocalInstrumentCode(ByVal locinst As String)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^(B2B|COR1|CORE)$"
    
    If re.Test(locinst) Then
        varLocalInstrumentCode = locinst
    Else
        Dim ex As New SEPAException
        ex.throw ("Only 'CORE', 'COR1', or 'B2B' are allowed.")
    End If
    
End Property

' SequenceType read access
Public Property Get SequenceType() As String
    SequenceType = varSequenceType
End Property


' SequenceType write access
Public Property Let SequenceType(ByVal seqty As String)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^(FNAL|FRST|OOFF|RCUR)$"
    
    If re.Test(seqty) Then
        varSequenceType = seqty
    Else
        Dim ex As New SEPAException
        ex.throw ("Only 'FNAL', 'FRST', 'OOFF', or 'RCUR' are allowed.")
    End If
    
End Property

' RequestedCollectionDate read access
Public Property Get RequestedCollectionDate() As String
    
    If varRequestedCollectionDate <> "" Then
        RequestedCollectionDate = varRequestedCollectionDate
    Else
        Dim reqDate As String
        reqDate = Format(Now(), "yyyy-mm-dd")
        RequestedCollectionDate = reqDate
    End If
          
End Property

' RequestedCollectionDate write access
Public Property Let RequestedCollectionDate(ByVal reqcol As String)
    varRequestedCollectionDate = reqcol
End Property

' CreditorName read access
Public Property Get CreditorName() As String
    CreditorName = varCreditorName
End Property

' CreditorName write access
Public Property Let CreditorName(ByVal cname As String)
    
    If Len(cname) = 0 Or Len(cname) > 70 Then
        Dim ex As New SEPAException
        ex.throw ("Invalid creditor name (max. 70 characters).")
    Else
        varCreditorName = cname
    End If
    
End Property

' CreditorAccountIBAN read access
Public Property Get CreditorAccountIBAN() As String
    CreditorAccountIBAN = varCreditorAccountIBAN
End Property

' CreditorAccountIBAN write access
Public Property Let CreditorAccountIBAN(ByVal creiban As String)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[a-zA-Z]{2}[0-9]{2}[a-zA-Z0-9]{4}[0-9]{7}([a-zA-Z0-9]?){0,16}$"
    
    If re.Test(creiban) Then
        varCreditorAccountIBAN = creiban
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid creditor IBAN.")
    End If
    
End Property

' CreditorAgentBIC read access
Public Property Get CreditorAgentBIC() As String
    CreditorAgentBIC = varCreditorAgentBIC
End Property

' CreditorAgentBIC write access
Public Property Let CreditorAgentBIC(ByVal crebic As String)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[0-9a-z]{4}[a-z]{2}[0-9a-z]{2}([0-9a-z]{3})?$"
    
    If re.Test(crebic) Then
        varCreditorAgentBIC = crebic
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid creditor BIC.")
    End If
    
End Property

' CreditorSchemeIdentification read access
Public Property Get CreditorSchemeIdentification() As String
    CreditorSchemeIdentification = varCreditorSchemeIdentification
End Property

' CreditorSchemeIdentification write access
Public Property Let CreditorSchemeIdentification(ByVal cresche As String)

    If cresche = "" Or cresche = Null Then
        Dim ex As New SEPAException
        ex.throw ("Invalid CreditorSchemeIdentification.")
    Else
        varCreditorSchemeIdentification = cresche
    End If
    
End Property


'/**
'   * Adds a transaction object to the list of transactions.
'   *
'   * @param SEPADirectDebitTransaction $transaction
'   */
Public Function addTransaction(ByVal transaction As SEPADirectDebitTransaction)
    
    Dim arrayLength As Integer
    
    If UBound(varTransactions) = -1 Then
        ReDim varTransactions(0 To 0)
    End If
    
    arrayLength = UBound(varTransactions)
    arrayLength = arrayLength + 1
    
    ReDim varTransactions(0 To arrayLength)
    
    Set varTransactions(arrayLength) = transaction
    
End Function


'  /**
'   * Returns a SimpleXMLElement for the SEPAPaymentInfo object.
'   *
'   * @return SimpleXMLElement object
'   */
Public Function getXmlPaymentInfo() As DOMDocument
    
    Dim objXML As DOMDocument
    Dim othr As IXMLDOMElement
    
    Dim domPaymentInfo As DOMDocument
    Dim domTransaction As DOMDocument
    
    Dim objXMLRootelement As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLelementSub As IXMLDOMElement
    Dim objXMLelementSubSub As IXMLDOMElement
    Dim objXMLelementSubSubSub As IXMLDOMElement
    
    Set objXML = New DOMDocument
    
    Set objXMLRootelement = objXML.createElement("PmtInf")
    objXML.appendChild objXMLRootelement
    
    Set objXMLelement = objXML.createElement("PmtInfId")
    objXMLelement.Text = PaymentInformationIdentification
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("PmtMtd")
    objXMLelement.Text = PAYMENT_METHOD
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("BtchBookg")
    objXMLelement.Text = BatchBooking
    objXMLRootelement.appendChild objXMLelement

    Set objXMLelement = objXML.createElement("NbOfTxs")
    objXMLelement.Text = NumberOfTransactions
    objXMLRootelement.appendChild objXMLelement

    Set objXMLelement = objXML.createElement("CtrlSum")
    objXMLelement.Text = ControlSum
    objXMLRootelement.appendChild objXMLelement

    Set objXMLelement = objXML.createElement("PmtTpInf")
    Set objXMLelementSub = objXML.createElement("SvcLvl")
    Set objXMLelementSubSub = objXML.createElement("Cd")
    objXMLelementSubSub.Text = SERVICELEVEL_CODE
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub '+
    '
    Set objXMLelementSub = objXML.createElement("LclInstrm")
    Set objXMLelementSubSub = objXML.createElement("Cd")
    objXMLelementSubSub.Text = LocalInstrumentCode
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub '+
    '
    Set objXMLelementSub = objXML.createElement("SeqTp")
    objXMLelementSub.Text = SequenceType
    objXMLelement.appendChild objXMLelementSub '+
    '
    objXMLRootelement.appendChild objXMLelement '+

    Set objXMLelement = objXML.createElement("ReqdColltnDt")
    objXMLelement.Text = RequestedCollectionDate
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("Cdtr")
    Set objXMLelementSub = objXML.createElement("Nm")
    objXMLelementSub.Text = CreditorName
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("CdtrAcct")
    Set objXMLelementSub = objXML.createElement("Id")
    Set objXMLelementSubSub = objXML.createElement("IBAN")
    objXMLelementSubSub.Text = CreditorAccountIBAN
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("CdtrAgt")
    Set objXMLelementSub = objXML.createElement("FinInstnId")
    Set objXMLelementSubSub = objXML.createElement("BIC")
    objXMLelementSubSub.Text = CreditorAgentBIC
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("ChrgBr")
    objXMLelement.Text = CHARGE_BEARER
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("CdtrSchmeId")
    Set objXMLelementSub = objXML.createElement("Id")
    Set objXMLelementSubSub = objXML.createElement("PrvtId")
    Set objXMLelementSubSubSub = objXML.createElement("Othr")
    '
    Set othr = objXMLelementSubSubSub
    '
    objXMLelementSubSub.appendChild objXMLelementSubSubSub
    objXMLelementSub.appendChild objXMLelementSubSub
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    If CreditorSchemeIdentification <> "" Then
        Set objXMLelement = objXML.createElement("Id")
        objXMLelement.Text = CreditorSchemeIdentification
        othr.appendChild objXMLelement
    End If
    
    Set objXMLelement = objXML.createElement("SchmeNm")
    Set objXMLelementSub = objXML.createElement("Prtry")
    objXMLelementSub.Text = PROPRIETARY_NAME
    objXMLelement.appendChild objXMLelementSub
    othr.appendChild objXMLelement
    
    'Add all transactions to the current PaymentInfo block and refresh NbOfTxs and CtrlSum
    Dim ctrlSum As Currency
    Dim i As Integer
    
    ctrlSum = 0#
    Me.NumberOfTransactions = UBound(varTransactions)
        
    For i = 1 To UBound(varTransactions)
    
        objXML.DocumentElement.appendChild varTransactions(i).XmlDirectDebitTransaction.DocumentElement.CloneNode(True)
        
        ctrlSum = ctrlSum + varTransactions(i).InstructedAmount
    
    Next i

    ' refresh
    ControlSum = ctrlSum
    objXMLRootelement.SelectSingleNode("NbOfTxs").Text = NumberOfTransactions
    objXMLRootelement.SelectSingleNode("CtrlSum").Text = ControlSum
    
    'objXML.save ("\\192.168.3.6\ablage_tmp$\test.xml")
    
    Set getXmlPaymentInfo = objXML

End Function
