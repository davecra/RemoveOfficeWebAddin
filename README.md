![Logo: Outlook icon with a cursor clicking on a red-x to the side](https://davecra.files.wordpress.com/2019/09/remove-1.png)
# The "Remove Office WebAddin from Outlook" Add-in
This is a Visual Studio Tools for Office (VSTO) Add-in that removes a Web Add-in's Customization from the Outlook for Windows Ribbon.

This is useful for when you have created a Web Add-in but it is not as feature rich as your COM add-in. So on the Windows clients, you still want your COM (or VSTO) solution to run there and not the web version of the add-in.


**NOTE:** If you are following the latest updates the the OfficeJS API’s, you will see that PowerPoint, Excel and Word have this new feature that allows you to specify an equivalent COM add-in in place of the Web Add-in. Here is the link:  https://docs.microsoft.com/en-us/office/dev/add-ins/develop/make-office-add-in-compatible-with-existing-com-add-in

## Configuration
Configuring and setting up this add-in is documented in the following Blog post:
[Removing Web Add-in Ribbon Customization in Outlook for Windows](https://theofficecontext.com/2019/09/26/removing-web-add-in-ribbon-customization-in-outlook-on-windows/)
