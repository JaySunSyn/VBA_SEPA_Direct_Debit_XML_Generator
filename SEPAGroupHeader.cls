VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SEPAGroupHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'
'   * Point to point reference assigned by the instructing party and sent to the next party in the chain
'   * to unambiguously identify the message. (Tag: <MsgId>)
'   *
'   * @var string
'   */
Private varMessageIdentification As String

'
'   * Date and time at which a (group of) payment instruction(s) was created
'   * by the instructing party. (Tag: <CreDtTm>)
'   *
'   * @var DateTime object
'   */
Private varCreationDateTime As String

 
'   * Number of individual transactions contained in the message. (Tag: <NbOfTxs>)
'   *
'   * @var int
'   */
Private varNumberOfTransactions As Integer


'   * Total of all individual amounts included in the message. (Tag: <CtrlSum>)
'   *
'   * @var float
'   */
Private varControlSum As Currency


'   * Party that initiates the payment. This can either be the creditor or a party that initiates
'   * the direct debit on behalf of the creditor. (Tag: <InitgPty->Nm>)
'   *
'   * @var string
'   */
Private varInitiatingPartyName As String

'Constructor
Private Sub Class_Initialize()

    'Variable initialization
    varMessageIdentification = ""
    varCreationDateTime = ""
    varNumberOfTransactions = 0
    varControlSum = 0#
    varInitiatingPartyName = ""
    
End Sub

'   * Getter for Message Identification (MsgId)
'   *
'   * @return string
'   */
Public Property Get MessageIdentification() As String

    MessageIdentification = varMessageIdentification

End Property

'   * Setter for Message Identification (MsgId)
'   *
'   * @param $msgId
'   * @throws SEPAException
'   */
Public Property Let MessageIdentification(ByVal msgId As String)

    '$msgId = URLify::downcode($msgId, "de");
    
    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^([A-Za-z0-9]|[\+|\?|\/|\-|:|\(|\)|\.|,|'| ]){1,35}$"
    
    If re.Test(msgId) Then
        varMessageIdentification = msgId
    Else
        Dim ex As New SEPAException
        ex.throw ("MsgId empty, contains invalid characters or too long (max. 35).")
    End If

End Property

'   * Getter for Creation Date Time (CreDtTm)
'   *
'   * @return DateTime object
'   */
Public Property Get CreationDateTime() As String

    If varCreationDateTime <> "" Then
        CreationDateTime = varCreationDateTime
    Else
        Dim reqDate As String
        reqDate = Format(Now(), "yyyy-mm-ddTHH:MM:SS")
        CreationDateTime = reqDate
    End If

End Property

'   * Setter for Creation Date Time (CreDtTm)
'   *
'   * @param string $creDtTm
'   */
Public Property Let CreationDateTime(ByVal creDtTm As String)

    varCreationDateTime = creDtTm

End Property

'   * Getter for Number of Transactions (NbOfTxs)
'   *
'   * @return int
'   */
Public Property Get NumberOfTransactions() As Integer

    NumberOfTransactions = varNumberOfTransactions

End Property

'   * Setter for Number of Transactions (NbOfTxs)
'   *
'   * @param int $nbOfTxs
'   * @throws SEPAException
'   */
Public Property Let NumberOfTransactions(ByVal nbOfTxs As Integer)

    Dim re As New VBScript_RegExp_55.RegExp
    re.Pattern = "^[0-9]{1,15}$"
    
    If re.Test(nbOfTxs) Then
        varNumberOfTransactions = nbOfTxs
    Else
        Dim ex As New SEPAException
        ex.throw ("Invalid NbOfTxs value (max. 15 digits).")
    End If

End Property

'   * Getter for ControlSum (CtrlSum)
'   *
'   * @return float
'   */
Public Property Get ControlSum() As Currency

    ControlSum = varControlSum

End Property

'   * Setter for ControlSum (CtrlSum)
'   *
'   * @param float $ctrlSum
'   */
Public Property Let ControlSum(ByVal ctrlSum As Currency)

   varControlSum = ctrlSum

End Property

'   * Getter for Initiating Party Name (InitgPty->Nm)
'   *
'   * @return string
'   */
Public Property Get InitiatingPartyName() As String

    InitiatingPartyName = varInitiatingPartyName

End Property

'   * Setter for Initiating Party Name (InitgPty->Nm)
'   *
'   * @param string $initgPty
'   * @throws SEPAException
'   */
Public Property Let InitiatingPartyName(ByVal initgPty As String)

    '$initgPty = URLify::downcode($initgPty, "de");

    If Len(initgPty) = 0 Or Len(initgPty) > 70 Then
        Dim ex As New SEPAException
        ex.throw ("Invalid initiating party name (max. 70).")
    Else
        varInitiatingPartyName = initgPty
    End If

End Property

'   * Returns a SimpleXMLElement for the SEPAGroupHeader object.
'   *
'   * @return SimpleXMLElement object
'   */
Public Property Get XmlGroupHeader() As DOMDocument

    Dim objXML As DOMDocument
    Dim objXMLRootelement As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    Dim objXMLelementSub As IXMLDOMElement
    
    Set objXML = New DOMDocument
    
    Set objXMLRootelement = objXML.createElement("GrpHdr")
    objXML.appendChild objXMLRootelement
    
    Set objXMLelement = objXML.createElement("MsgId")
    objXMLelement.Text = MessageIdentification
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("CreDtTm")
    objXMLelement.Text = CreationDateTime
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("NbOfTxs")
    objXMLelement.Text = NumberOfTransactions
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("CtrlSum")
    objXMLelement.Text = ControlSum
    objXMLRootelement.appendChild objXMLelement
    
    Set objXMLelement = objXML.createElement("InitgPty")
    Set objXMLelementSub = objXML.createElement("Nm")
    objXMLelementSub.Text = InitiatingPartyName
    objXMLelement.appendChild objXMLelementSub
    objXMLRootelement.appendChild objXMLelement
    
    Set XmlGroupHeader = objXML

End Property
