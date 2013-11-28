VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SEPAMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'End Property
'   * The used ISO 20022 message format.
'   *
'   * @var string
'   */
Private varMessageDefintion As String

'  End Property
'   * GroupHeader object. (Tag: <GrpHdr>)
'   *
'   * @var SEPAGroupHeader object
'   */
Private varGroupHeader As SEPAGroupHeader

'  End Property
'   * A container for all PaymentInfo objects.
'   *
'   * @var SEPAPaymentInfo[] array
'   */
Private varPaymentInfos() As SEPAPaymentInfo

Private varMessageDoc As DOMDocument

'  End Property
'   * The constructor sets the used message format.
'   *
'   * @param string $messageDefinition
'   */

'Constructor
Private Sub Class_Initialize()

    'Variable initialization
    varMessageDefintion = ""
    'Set varGroupHeader = Nothing
    ReDim varPaymentInfos(0 To 0)
    

End Sub

'messageDefinition with constructor in vba not possible
Public Property Let MessageDefinition(ByVal msgDef As String)
    
    varMessageDefintion = msgDef

End Property

Public Property Get MessageDefinition() As String
    
    MessageDefinition = varMessageDefintion
    
End Property

Public Property Get groupHeader() As SEPAGroupHeader

    If varGroupHeader.NumberOfTransactions = 0 Then
        Dim gh As SEPAGroupHeader
        Set gh = New SEPAGroupHeader
        Set groupHeader = gh
        
    End If
    
    Set groupHeader = varGroupHeader


End Property

'   * Setter for the group header object.
'   *
'   * @param SEPAGroupHeader $groupHeader
'   */
Public Property Let groupHeader(ByVal gpHeader As SEPAGroupHeader)

    Set varGroupHeader = gpHeader

End Property

'   * Adds a PaymentInfo object to the collection.
'   *
'   * @param SEPAPaymentInfo $paymentInfo
'   */
Public Function addPaymentInfo(ByVal paymentInfo As SEPAPaymentInfo)
   
    Dim arrayLength As Integer
    
    If IsEmpty(varPaymentInfos) Then
        ReDim varPaymentInfos(0 To 0)
    End If
    
    If UBound(varPaymentInfos) = -1 Then
        ReDim varPaymentInfos(0 To 0)
    End If
    
    arrayLength = UBound(varPaymentInfos)
    arrayLength = arrayLength + 1
    
    ReDim varPaymentInfos(0 To arrayLength)
    
    
    Set varPaymentInfos(arrayLength) = paymentInfo

End Function

'   * Returns the message as SimpleXMLElement object.
'   *
'   * @return SimpleXMLElement object
'   */
Public Function getXML() As DOMDocument

    Dim message As DOMDocument
    Dim messageRootElement As IXMLDOMElement
    
    Set message = New DOMDocument
    
    Set messageRootElement = message.createElement("CstmrDrctDbtInitn")
    message.appendChild messageRootElement
    
    addSubtree message, groupHeader.XmlGroupHeader
    
    ' Add all payment blocks
    Dim nbOfTxs As Integer
    Dim ctrlSum As Currency
    Dim payInfo As Variant
    Dim i As Integer
    
    nbOfTxs = 0
    ctrlSum = 0#

    For i = 1 To UBound(varPaymentInfos)
        
        Dim payDoc As DOMDocument
        Dim payIn As SEPAPaymentInfo

        Set payIn = varPaymentInfos(i)

        'Set payDoc = New DOMDocument
        Set payDoc = payIn.getXmlPaymentInfo
        
        addSubtree message, payDoc
        
        nbOfTxs = nbOfTxs + 1
        ctrlSum = ctrlSum + payIn.ControlSum
        
    Next i
    
    Me.groupHeader.NumberOfTransactions = nbOfTxs
    Me.groupHeader.ControlSum = ctrlSum
    
    'Finally add the XML structure
    Dim doc As DOMDocument
    Dim docRoot As IXMLDOMElement
    Dim objXMLelement As IXMLDOMElement
    
    Set doc = New DOMDocument
    
'    Set docRoot = doc.createElement("xml")
'    docRoot.setAttribute "version", "1.0"
'    docRoot.setAttribute "encoding", "utf-8"
'    doc.appendChild docRoot
    
    
    Set objXMLelement = doc.createElement("Document")
    objXMLelement.setAttribute "xmlns", MessageDefinition
    objXMLelement.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
    doc.appendChild objXMLelement
    
    addSubtree doc, message
    
    Set getXML = doc

  End Function
  
'   * Returns a printable and formatted XML message.
'   *
'   * @return string
'   */
Public Function printXML()
'  {
'    @header('Content-type: text/xml');
'
'    $dom = new DOMDocument();
'    $dom->preserveWhiteSpace = false;
'    $dom->formatOutput = true;
'    $dom->loadXML($this->getXML()->asXML());
'
'    return $dom->saveXML();
'  }

    Dim doc As DOMDocument
    
    Set doc = New DOMDocument
    doc.LoadXML Me.getXML
    
    doc.save ("\\192.168.3.6\ablage_tmp$\test99.xml")

End Function

'   * Validate the generated XML over the .xsd specification.
'   *
'   * @param string $schemePath
'   * @return bool
'   */
Public Function validateXML(schemePath)
'  {
'    $dom = new DOMDocument();
'    $dom->loadXML($this->getXML()->asXML());
'
'    libxml_use_internal_errors(true);
'    if (!$dom->schemaValidate($schemePath))
'    {
'      echo "XML validation failed!\n";
'
'      $errors = libxml_get_errors();
'      for ($i = 0; $i < count($errors); $i++)
'        echo "Error " . $errors[$i]->code . ": " . $errors[$i]->message . "\n";
'
'      return false;
'    }
'
'    return true;
'  }

End Function

'   * Creates new DOM elements from two given SimpleXMLElement objects.
'   *
'   * @param SimpleXMLElement $xmlTarget
'   * @param SimpleXMLElement $xmlOrigin
'   */


Private Function addSubtree(ByVal xmlTarget As DOMDocument, ByVal xmlOrigin As DOMDocument)

    xmlTarget.DocumentElement.appendChild xmlOrigin.DocumentElement.CloneNode(True)
        
End Function
