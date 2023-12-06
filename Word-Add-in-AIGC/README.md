# Microsoft Excel Mail Merge Sample Office Add-in

[![Node.js build](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml/badge.svg)](https://github.com/microsoftgraph/msgraph-training-office-addin/actions/workflows/node.js.yml) ![License.](https://img.shields.io/badge/license-MIT-green.svg)

This sample demonstrates how to use the Microsoft Graph JavaScript SDK to send emails in Excel from Office Add-ins.

## How the sample add-in works
### Features
- Create Sample Data, including valid email address (required) and other information.
- Verify Template and Data, the To Line must contain the column name of the email address.
- Send Email, which will pop up a dialog to get the consent of Microsoft Graph. After sign-in, the email will be send out.

### Play the sample add-in demo
Click the button below and play the sample add-in demo:<br><br>
[<img src="https://github.com/MingjiaLiu1995/Word-Scenario-based-Add-in-Samples/assets/107099441/6155cd00-5e16-405e-82f4-28ed2e4ce54d" width="120"/>](https://office.live.com/start/Excel.aspx?culture=en-US&omextemplateclient=Excel&omexsessionid=c0a9c7a1-b954-45df-9295-8c1e21201f34&omexcampaignid=none&templateid=WA200006296&templatetitle=Mail%20Merge%20Add-in%20for%20Excel&omexsrctype=1)
<br>

#### Note：
- You need to have a Microsoft 365 account to try the sample. You can [sign up for the Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program) to get a free Microsoft 365 subscription.<br>
- [Optional] You need to register a web application with the Azure Active Directory admin center to send out emails. You can skip this step if you only want to try other features.<br>
    
    > 1. Open a browser and navigate to the [Microsoft Entra admin center](https://aad.portal.azure.com). Login using a **personal account** (aka: Microsoft Account) or **Work or School Account**.
    > 1. Select **Identity** in the left-hand navigation, then select **App registrations** under **Applications**.
    > 1. Select **New registration**. On the **App registrations** page, set the values as follows.
        - Set **Name** to `Office Add-in Graph Tutorial`.
        - Set **Supported account types** to **Accounts in any organizational directory and personal Microsoft accounts**.
        - Under **Redirect URI**, set the first drop-down to `Single-page application (SPA)` and set the value to `https://localhost:3000/consent.html`.
    > 1. Select **Register**. On the **Office Add-in Graph Tutorial** page, copy the value of the **Application (client) ID** and save it, you will need it in the next step.
    > **Note**: This step needs to be **performed only once** by add-in developer, aiming to integrate your app with the Microsoft identity platform and establishing the information that it uses to get tokens. After successful registration and add-in published, **customer can use it directly**, do not need to register again. 
    > 1. Edit the `taskpane.js` file and make the following changes.
        - Replace `YOUR_APP_ID_HERE` with the **Application Id** you got from the App Registration Portal.

### Expected result
When you click the button, you will open Excel online in a new browser tab, and the sample add-in will launch automatically.

![image](https://github.com/MingjiaLiu1995/Word-Scenario-based-Add-in-Samples/assets/107099441/8ce88850-66d0-4880-824c-443595e55172)

## Build, run and debug the sample code 
To run this sample in your local environment, copy the following command and paste in your console to run:

    iwr aka.ms/wordaddin/aigc -o launch_addin.bat; saps launch_addin.bat

When you run the command, it will install all the required dependencies, and execute all the required steps to automatically launch the sample in your Visual Studio Code. Start debugging the project by hitting the `F5` key.
If you don't have VSC installed, the command will automatically open the sample folder.
 
### Sideload the sample add-in on Excel Online
The previous steps show you how to run our sample on Desktop. As for the Excel Online, please follow the following steps to sideload the manifest.xml file on web.
1.  **Keep the webpack server on** to host your sample add-in.
1.  Open [Office on the web](https://office.live.com/).
1.  Choose **Excel**, and then open a new document.
1.  On the **Home** tab, in the **Add-ins** section, choose **Add-ins** and click **More Add-ins** on the lower-right corner to open Add-in Store Page.
1.  On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.
   ![image](https://github.com/MingjiaLiu1995/Word-Scenario-based-Add-in-Samples/assets/107099441/bae7caf0-e3d1-4f7f-940a-ff6a537dc13e)
1.  Browse to the localhost add-in manifest file(manifest-localhost.xml), and then select **Upload**.
 ![image](https://github.com/MingjiaLiu1995/Word-Scenario-based-Add-in-Samples/assets/107099441/f91ce0ac-8a05-40b4-8139-dac68f80ed15)
1.  Verify that the add-in loaded successfully. 

## Additional resources
You may explore additional resources at the following links:
- More samples: [Office Add-ins code samples](https://github.com/OfficeDev/Office-Add-in-samples)
- Office add-ins documentation: [Office Add-ins documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)

## Feedback
Did you experience any problems with the sample? [Create an issue]( https://github.com/OfficeDev/Word-Scenario-based-Add-in-Samples/issues/new) and we'll help you out.

## Copyright
Copyright (c) 2021 Microsoft Corporation. All rights reserved.
This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
<br>**Note**: The taskpane.html file contains an image URL that tracks diagnostic data for this sample add-in. Please remove the image tag if you reuse this sample in your own code project.
<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/word-add-in-aigc">
