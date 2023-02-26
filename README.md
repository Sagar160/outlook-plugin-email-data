# Outlook-Addin-python
Template to get start started writing a TaskPane Outlook Add-in using python for the backend

## Test webapp before deployment
You can run flask locally and then `temporarily` change the manifest_python.xml file to your local development url.
1. 
4. Go to OWA and add your maniest.xml. There are multiple methods dependings on your permissions and rights to your azure ad - 
<br><br>If you are an admin: https://docs.servicenow.com/bundle/quebec-employee-service-management/page/product/workplace-reservations-outlook-addin/task/upload-the-manifest-file-office365.html <br><br>
  If you are a normal user: <br>
    Navigate to the “Add-Ins” menu after logging into Outlook on the web <br>
    Go to `Add-in management` under “...” ('More actions' menu which is visible from a selected email) <br>
    Click “Get Add-ins” from the "..." menu <br>
    From the “Custom Add-ins” pop-up click the “My Add-ins” tab <br>
    On the “My Add-ins” tab scroll down to bottom of the page and click the “+ Add a Custom Add-in link. <br>
    In the older version of Outlook web you may have to go to Add-in management is available under “Settings” > “Manage Add-Ins” >  “Custom Add-ins” > “My Add-ins” > “+ Add a Custom Add-in”
    select `file` and chose your `manifest_python.xml` file, then choose `install`
  
 **Note** - You will have to change your manifest file to the hosting url later. So you will have to remove and readd the new manifest file later on. 

## Deploy the sample
Follow the same instructions given from the microsoft website: https://docs.microsoft.com/en-us/azure/app-service/quickstart-python?tabs=powershell&pivots=python-framework-flask#deploy-the-sample

Once the sample is deployed, test your webapp's urls to make sure they work. 
After, you can modify the `manifest_python.xml` to route all `https://localhost:3000` urls to your own hosted url.
