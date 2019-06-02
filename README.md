# Reply-And-Keep-Attachments

**Reply and Keep Attachments** is a simple VSTO Outlook Add-In to help keep attachments when replying to an e-mail in Outlook.

![alt text](https://github.com/levoment/Reply-And-Keep-Attachments/blob/master/Images/Reply_And_Keep_Attachments_Screenshot.PNG "Reply and Keep Attachments screen capture")

The Add-In tab appears when an e-mail is opened for reading in a new windowv(double click on e-mail in the list).

# Features:

  - Keep attachments when replying
  - Keep attachments when replying to all

# How to test:

  - **Clone** the repository
  - Open the solution in **Visual Studio**
  - Delete: **Reply_And_Keep_Attachments_TemporaryKey.pfx** if found in the **Solution Explorer**
  - Right-click the project in the **Solution Explorer**
  - Click **Properties**
  - Go to the **Signing** tab
  - Click **Create Test Certificate...**
  - Enter a new password, choose the **Signature algorithm**, and click **OK**
  - **Save All** (Ctrl + Shift + S)
  
  The project will be ready to debug if no errors ocurred during the creation of the test signing key.
  
  # Troubleshooting:
  
  ## Add-In crashed Outlook while debugging and now it doesn't appear despite attempts to re-enable it on Outlook
  
   1. Open Outlook again
   2. Close Outlook
   3. Open the Registry Editor
   4. Locate: **HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems** and delete the binary value corresponding to the Add-In
   5. Locate: **HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\CrashingAddinList** and delete the binary value corresponding to the Add-In
   6. Locate: **HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Addins\Reply And Keep Attachments** and delete the key
   7. Locate: **HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\Reply And Keep Attachments** and delete the key
  
  # Deploying:
  - [Using a Windows Installer](https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-windows-installer?view=vs-2019)
  - [Using ClickOnce](https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution-by-using-clickonce?view=vs-2019)
  
