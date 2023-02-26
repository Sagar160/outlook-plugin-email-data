# Outlook-Addin-python
Execute the action instead of task pane. For env requirement: flask, openssl

## Test webapp before deployment
You can run flask locally and then `temporarily` change the manifest_python.xml file to your local development url.

For xml/plugin installation. follow these steps:
  If you are a normal user: <br>
    Navigate to the “Add-Ins” menu after logging into Outlook on the web <br>
    Go to `Add-in management` under “...” ('More actions' menu which is visible from a selected email) <br>
    Click “Get Add-ins” from the "..." menu <br>
    From the “Custom Add-ins” pop-up click the “My Add-ins” tab <br>
    On the “My Add-ins” tab scroll down to bottom of the page and click the “+ Add a Custom Add-in link. <br>
    In the older version of Outlook web you may have to go to Add-in management is available under “Settings” > “Manage Add-Ins” >  “Custom Add-ins” > “My Add-ins” > “+ Add a Custom Add-in”
    select `file` and chose your `manifest_python.xml` file, then choose `install` 

## Deploy flask app
flask run --host=localhost --port 3000 --cert=adhoc
