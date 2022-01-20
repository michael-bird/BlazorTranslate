# ThreeShape.SilverLake.Experiments.SIL85.BlazorReact

This project explores various ways of implementing the NG components into a Blazor application.

 - Embedded in iframe
 - Embedded within Blazor application
 - Recreated as Blazor components

> **Note:** This project requires node.js be installed. You must also have your local Silverlake environment up and running

One of the scripts also requires the following package:
````
npm install --global del-cli
````


## Getting started

 - Start the node.js services in your local environment. For the iframe to work you will need to wait for the frontend to be packed which can take up to 5 minutes.
 - Run the Application - It should use webpack to bundle the JS and SCSS files.
 - This application communicates with the services over port: 3000 - If you
   have changed this you will need to update Services/LabstarService.cs.