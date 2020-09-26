Attribute VB_Name = "XMLImportFunctionality"
Sub vba_api()
Dim xml_obj As MSXML2.XMLHTTP60

Set xml_obj = New MSXML2.XMLHTTP60

    'define url components
    base_url = "https://maps.googleapis.com/maps/api/place"
    end_pt = "/nearbysearch/xml?"
    
    param_loc = "location="
    param_loc_val = "-33.8670522,151.1957362"
    
    param_rad = "&radius="
    param_rad_val = "1500"
    
    param_type = "&type="
    param_type_val = "restaurant"
    
    param_key = "&key="
    param_key_val = ThisWorkbook.Sheets("Sheet2").Range("B2").Text
    
    
    'combine components to create URL
    
    api_URL = base_url + end_pt + _
    param_loc + param_loc_val + _
    param_rad + param_rad_val + _
    param_type + param_type_val + _
    param_key + param_key_val
    
    'open a new request using our URL
    xml_obj.Open bstrmethod:="GET", bstrUrl:=api_URL
    
    
    
    
    
    'send the request
    xml_obj.send
    
    'Print the status code in case something went wrong
    Debug.Print "The request was" + xml_obj.statusText
    Debug.Print xml_obj.responseText
    
    'Define variables needed to create our xml document
    Dim xdoc As MSXML2.DOMDocument60
    Dim xnodes As MSXML2.IXMLDOMNodeList
    Dim xnode As MSXML2.IXMLDOMNode
    
    'First create the document
    Set xdoc = New MSXML2.DOMDocument60
    xdoc.LoadXML (obj_xml.responseText)
    
    'look at the first child node'
    Debug.Print xdoc.ChildNodes.Item(1).BaseName
    
    'Find the nodes that contain the result
    Set xnodes = xdoc.SelectNodes("/PlacesearchResponse/result")
    
    'loop through the result sets
    For Each xnode In xnodes
        Debug.Print "---------------------------------"
        Debug.Print xnode.SelectSingleNode("name").Text
        Debug.Print xnode.SelectSingleNode("place_id").Text
    Next
    
    'Define the worksheet to print the data
    Dim wrksht As Worksheet
    Set wrksht = ThisWorkbook.Worksheets("Sheet1")
    
    
    'Loop through and dump the data
    Count = 1
    
    For Each xnode In xnodes
        wrksht.Cells(Count, 1).Value = xnode.SelectSingleNode("name").Text
        wrksht.Cells(Count, 2).Value = xnode.SelectSingleNode("place_id").Text
        Count = Count + 1
        
    Next
    
    'Loop through and dump the data method 2
    For i = 0 To xnodes.Length - 1
    
        wrksht.Cells(i + 1, 4).Value = xnode.SelectSingleNode("name").Text
        wrksht.Cells(i + 1, 5).Value = xnode.SelectSingleNode("place_id").Text
        
    Next
    
    'Debug.Print api_URL
End Sub
    
    

