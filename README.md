# Dynamics 365 Integration

## Overview

The Onsight Connector provides integration within Dynamics 365 to Onsight Connect and Onsight Workspace. It is provided as a set of installable solutions via the Microsoft AppSource catalog.

## Pre-requisites

The Onsight Connector stores your Onsight API key in an environment variable. In order to read this variable's value at runtime, each user must have the *Environment Variable Definition* permission enabled. The easiest way to enable this is by enabling it on one or more user roles within Dynamics 365's *Advanced Settings > Security > Security Roles* page, under the **Custom Entities** tab.

## Installation

To install the Connector into your D365 tenant, go to the Microsoft AppSource site, https://appsource.microsoft.com/, and search for "Onsight Field Service Connector".

From the *Onsight Dynamics 365 Field Service Connector* page, click the "Get It Now" button. After entering some contact information, click "Continue".

Select the D365 environment into which the Connector should be installed.

## Post Installation

After installing, you must enter your Onsight API Key into the designated environment variable.

As a D365 administrator, go to the Power Apps Solutions page https://make.powerapps.com/environments/{tenantId}/solutions and open the Default Solution.

![](images/Connect-EnvVariables.png)

Under the Default Solution's Objects list, locate the Environment Variable named OnsightAPIKey. Open this Environment Variable and set its Current Value to your Onsight API Key. Click Save.


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
