# Solidworks VBA

# Modelling 
## Initialize Model
#### To Create a New Document
```VB
' Create Solidworks document 
Dim swDoc As SldWorks.ModelDoc2
' Creating string type variable for storing default part location 
Dim defaultTemplate As String
' Setting value of this string type variable to "Default part template" 
defaultTemplate = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
' Setting to new part document (Teamplate Name, Paper Size, Width, Height) 
Set swDoc = swApp.NewDocument(defaultTemplate, 0, 0, 0)
```

#### To Use Existing & Active document
```VB
Set swDoc = swApp.ActiveDoc
' From Directory 
Set swDoc = swApp.OpenDoc("H:\part.SLDPRT", swDocumentTypes_e.swDocPART)
' For more option while opening a document  
//File, Type, Options(mode to open doc), Configuartion ("" to use previous one), Errors (swFileLoadError_e), Warnings (swFileLoadWarning_e)
Set swDoc = swApp.OpenDoc6("H:\part.SLDPRT", swDocumentTypes_e.swDocPART, swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0)
```

## Units
- Sw only take value in meters
- Use this when we need to define any length or angle value.
  
```VB
' Variables used as Conversion Factors 
  Dim LengthCF As Double
  Dim AngleCF As Double
  
  ' Use a Select Case, to get the length of active Unit and set the different factors 
  Select Case swDoc.GetUnits(0)       ' GetUnits function gives us, active unit
    
    Case swMETER    ' If length is in Meter
      LengthCF = 1 
      AngleCF = 1
    
    Case swMM       ' If length is in MM
      LengthCF = 1 / 1000
      AngleCF = 1 * 0.01745329
    
    Case swCM       ' If length is in CM
      LengthCF = 1 / 100
      AngleCF = 1 * 0.01745329
    
    Case swINCHES   ' If length is in INCHES
      LengthCF = 1 * 0.0254
      AngleCF = 1 * 0.01745329
    
    Case swFEET     ' If length is in FEET
      LengthCF = 1 * (0.0254 * 12)
      AngleCF = 1 * 0.01745329
    
    Case swFEETINCHES     ' If length is in FEET & INCHES
      LengthCF = 1 * 0.0254  ' For length we use sama as Inch
      AngleCF = 1 * 0.01745329
    
    Case swANGSTROM        ' If length is in ANGSTROM
      LengthCF = 1 / 10000000000#
      AngleCF = 1 * 0.01745329
    
    Case swNANOMETER       ' If length is in NANOMETER
      LengthCF = 1 / 1000000000
      AngleCF = 1 * 0.01745329
    
    Case swMICRON       ' If length is in MICRON
      LengthCF = 1 / 1000000
      AngleCF = 1 * 0.01745329
  End Select

```
