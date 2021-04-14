# Dynamics 365 Integration

## Overview

The Onsight Connect and Workspace integration consists of two Dynamics 365 solutions:
* Onsight Connect
    - Contains ribbon/command bar buttons for Connect and Workspace, with hooks
      into the custom javascript handlers for each button.
* Onsight Connect resources
    - JavaScript customizations
    - Command Button icons
    - Environment variable "OnsightAPIKey" for the end customer's Onsight API key.

The integration is split into two solutions due to limitations in the Ribbon Workbench editor, the tool used to add custom buttons to the D365 ribbon/command bar.

Both solutions have been left as "unmanaged", meaning that they can both continue to be modified as needed.

## Pre-requisites

The Onsight extension stores your Onsight API key in an environment variable. In order to read this variable's value at runtime, each user must have the *Environment Variable Definition* permission enabled. The easiest way to enable this is by enabling it on one or more user roles within Dynamics 365's *Advanced Settings > Security > Security Roles* page, under the **Custom Entities** tab.

## Installation

1) From source, zip up both the *OnsightConnect_1_0_0* and *OnsightConnectResources_1_0_0* folders into their own respective zip files.
2) From your Dynamics 365 account, navigate to the Power Apps site.
3) Select *Solutions* from the left-hand navigation panel.
4) From the top toolbar, select *Import*.
5) Browse to the *OnsightConnectResources_1_0_0.zip* file you archived in step 1 and upload it.
6) Upload the *OnsightConnect_1_0_0.zip* file as well.
7) Ensure that both solutions are now listed under *Solutions*.
8) Click on the *Onsight Connect resources* solution.
9) Click on the *OnsightAPIKey* list item. A panel should open up on the right-hand side.
10) In the *Current Value* text box, enter your Onsight API Key and click the Save button.

## Using the Integration

### Launching Onsight Connect
* From the Work Order main form:
![](images/Connect-FromWorkOrder.png)
    - This calls the first Bookable Resource Booking (i.e., field worker) assigned to the Work Order
        - Design Note: this is based on the "EmailMappings" JSON structure (in new_OnsightConnectApi.js) which assumes a WorkOrder will map to a single systemuser (it could map to several systemusers, as a WorkOrder is a complex entity).
        Future enhancements might require the EmailMapping structure to yield multiple systemusers for a single WorkOrder, such as the Support Contact/Expert, and any
        additional Bookable Resource Bookings.
        - When creating the Onsight Connect call, the Work Order ID and the Work Order's primary Customer Asset ID are sent as metadata and can be later queried using Workspace.
* From other locations (see **IMPORTANT NOTE** below)
    - From the Bookings sub grid:
        ![](images/Connect-FromBookingsSubGrid.png)
    - From the Bookable Resource booking dialog:
        ![](images/Connect-FromBookableResourceBookingDialog.png)
    - From the Bookable Resource main form:
        ![](images/Connect-FromBookableResource.png)

#### IMPORTANT NOTE:
If launching Onsight Connect from *anywhere other than* the Work Order's main form, there will be no Work Order in context. In other words, Onsight Connect can be launched, but no Work Order-related metadata will be available and associated with the subsequent call.

### Launching Workspace
* From the Work Order main form:
![](images/Workspace-FromWorkOrder.png)
The Work Order ID will be used as the Workspace search term.
* From the Customer Asset main form:
![](images/Workspace-FromCustomerAsset.png)
The Customer Asset ID will be used as the Workspace search term.
