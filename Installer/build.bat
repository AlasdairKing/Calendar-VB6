@title Windows Installer using WiX
@if exist Calendar.wixobj del Calendar.wixobj
@candle Calendar.wxs -nologo -ext WixNetfxExtension -ext WixUtilExtension -ext wixTagExtension -ext WixUiExtension
@if exist Calendar.msi del Calendar.msi 
@light Calendar.wixobj -spdb -sice:ICE91 -nologo -ext WixNetfxExtension -ext WixUtilExtension -ext wixTagExtension -ext WixUiExtension 
@pause

