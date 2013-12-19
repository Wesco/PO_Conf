PO_Conf
=======

Purchase Order Confirmation Macro

### Email Setup

1. Add the GetPONumber macro to your current outlook session  
2. Go to "Manage Rules and Alerts" and create a rule that matches the following criteria

> sent only to (your_email@wesco.com)
> and from (your_email@wesco.com)
> and with 'WESCO International, Inc. PO' in the subject
> and with 'Confirmation! Your email for WESCO International, Inc.' or 'Confirmation! Your FAX for WESCO International, Inc.' in the body
> and on this computer only
> move it to the (Sent_POs_Folder) folder
> and run GetPONumber
> and display a Desktop Alert
> except if it has an attachment  

### Excel Setup

1. Go to "\\\\br3615gaps\\gaps\\Excel Add-Ins\\Prerequisites\\"
2. Install dotNetFx40_Full_x86_x64.exe  
3. Install vstor_redist.exe  
4. Go to "\\\\br3615gaps\\gaps\\Excel Add-Ins\\SaveWorkbook\\publish\\"  
5. Install SaveWorkbook.vsto

### NetTerm Setup

Go to Setup > QuickButton keys...  
![Quick Button Keys](/Images/QuickButtons_1.jpg)  

Add a label named 473  
Add 470^M3^M00^MAY36153615YYYYYXY^M as the transmitted key data  
![Quick Button Data](/Images/QuickButtons_2.jpg)

### Instructions

1. Send 1 or more POs
2. Open NetTerm
3. Click "473"
4. In Excel click "Save Report"
5. Run the macro