# WordJS-O365-Angular-POTemplate2

Sample shows how you can get contact information from Office 365 and insert the contact 
information into a Word document by using WordJS.

## Run the app

Open *app/scripts/config.js* and replace *{your_tenant}* with the subdomain of .onmicrosoft you specified for your Office 365 tenant and replace *{client_ID}* with the client ID of your registered Azure application (found on the **Configure** tab of your application's entry in the Azure Management Portal).

Next, install the necessary dependencies and run the project via the command line. Begin by opening a command prompt and navigating to the root folder. Once there, follow the steps below.

1. Install project dependencies by running ```npm install```.
2. Now that all the project dependencies are installed, start the development server by running ```node server.js``` in the root folder.
3. Navigate to ```http://127.0.0.1:8080/``` in your web browser.
