# relativity-azure-ocr
Relativity mass event handled for document OCR/ICR using Azure Cognitive Services.

# Install
## 1) Create Instance Settings
Create required Relativity Instance Settings entries:  
Name | Section | Value Type | Value (example) | Description
---- | ------- | ---------- | --------------- | -----------
AzureSubscriptionKey | Azure.OCR | Text | xxxxxxxxx | Azure OCR Subscription Key.
AzureEndpoint | Azure.OCR | Text | xxxxxxxxx | Azure Endpoint.
DestinationField | Azure.OCR | Text | Extracted Text OCRed | Document Field where to record the OCRed text.
LogField | Azure.OCR | Text | Azure OCR Log | Document Field to store the OCR log.

## 2) Compile DLL
Download the source code and compile the code using Microsoft Visual Studio 2019.  
For more details on how to setup your development environemnt, please follow official [Relativity documentation](https://platform.relativity.com/10.3/index.htm#Relativity_Platform/Setting_up_your_development_environment.htm).  
You can also use precompiled DLL from the repository.

## 3) Upload DLL
Upload `RelativityAzureOcr.dll` to Relativity Resource Files.  
You may need to install also additional libraries that are required. These libraries were required for Relativity Server 2022:
* System.Text.Json.dll
* System.Memory.dll
* Microsoft.Bcl.AsyncInterfaces.dll
* System.Text.Encodings.Web.dll
* System.Runtime.CompilerServices.Unsafe.dll
* System.Buffers.dll
* System.Numerics.Vectors.dll

## 4) Add to Workspace
For desired workspaces add mass event handler to Document Object:
* Browse to Document Object (Workspace->Workspace Admin->Object Type->Document)
* In Mass Operations section click New and add the handler:
  * Name: Azure.OCR
  * Pop-up Directs To: Mass Operation Handler
  * Select Mass Operation Handler: RelativityAzureOcr.dll

# Log
Mass operation generates OCR log to fiels specified by the Relativity Instance Settings.  
Log entry is added after each OCR. There can be multiple log entries for one Document.  
Log entry has following fields:
* OCR engine
* User email address
* Timestamp
* Character count of the OCR text

OCR log can be viewed from the Relativity front-end via attached Relativity Script.

# Notes
Relativity Azure OCR mass operation was developed and tested in Relativity Server 2022.  
Relativity Azure OCR mass operation works correctly only with UTF-8 text.
