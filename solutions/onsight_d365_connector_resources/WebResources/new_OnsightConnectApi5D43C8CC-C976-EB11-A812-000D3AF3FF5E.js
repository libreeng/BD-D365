'use strict';

// depending on where this script is utilized (e.g. within an HTML web resource),
// we may have to obtain the Xrm object from the parent
var xrm = typeof(Xrm) !== "undefined" ? Xrm : parent.Xrm;
var globalContext = xrm ? xrm.Utility.getGlobalContext() : {};
var clientContext = globalContext ? globalContext.client : {};

const workspaceUrl = "https://workspace.librestream.com/onsight/ui/#!/search/default";

// URI for the launch request REST endpoint
const apiEndpointUrl = "https://onsight.librestream.com/OamRestAPI/api/launchrequest";

// OPM intermediate launch page
const opmLaunchPage = "https://onsight.librestream.com/OamAdministrator/AccountServices/Default.aspx";
const externalUri = "https://onsight.librestream.com/OamAdministrator/"

// Platform types supported for launching Onsight Connect
const PlatformTypes = {
    PC: "PC",
    Android: "Android",
    iOS: "iOS",
    Unknown: "Unknown"
};

// Regex for Xrm.WebApi calls returning an array of entities
const WebApiExpandExpr = /(\w+)\(\$select\=(\w+)\)$/;

const SystemUser2Email = [
    {
        entityLogicalName: "systemuser",
        select: "internalemailaddress"
    }
];

const BookableResource2Email = [
    {
        entityLogicalName: "bookableresource",
        select: "_userid_value"
    },
    ...SystemUser2Email
];

const BookableResourceBooking2Email = [
    {
        entityLogicalName: "bookableresourcebooking",
        select: "_resource_value"
    },
    ...BookableResource2Email
];

const WorkOrder2FieldTechEmail = [
    {
        entityLogicalName: "msdyn_workorder",
        expand: "msdyn_msdyn_workorder_bookableresourcebooking_WorkOrder($select=bookableresourcebookingid)"
    },
    ...BookableResourceBooking2Email
];

const WorkOrder2RemoteExpertEmail = [
    {
        entityLogicalName: "msdyn_workorder",
        select: "_msdyn_supportcontact_value"
    },
    ...BookableResource2Email
];


/**
 * Removes all curly braces from the given string.
 * @param {*} text The string from which curly braces are to be removed.
 */
function removeBraces(text) {
    return text.replace(/[\{\}]/g, '');
}

/**
 * Reads the customer's Onsight API key from the pre-defined environment variable "OnsightAPIKey".
 * Customers must set this environment variable before using the Onsight integration.
 */
async function getApiKeyAsync() {
    // Librestream API key issued to the domain to which the Onsight user is part of
    const result = await Xrm.WebApi.retrieveRecord("environmentvariablevalue", "738f3d87-5a7c-eb11-a812-000d3af3a657", "?$select=value");
    return result.value;
}

/**
 * Uses a mapping struct (such as an element of EmailMappings) to extract a value
 * from the given recordResult.
 * @param {*} recordResult 
 * @param {*} recordMapping 
 */
function mapResultToValue(recordResult, recordMapping) {
    // If the mapping defines an "expand" property, it means we expect an array result
    if (recordMapping.expand) {
        // Result must be an array
        let matches = recordMapping.expand.match(WebApiExpandExpr);
        if (matches && matches.length === 3) {
            let arrName = matches[1];
            let elemName = matches[2];

            // Grab the first element in the array (if available) and map the result value
            // from the expand expression's inner "$select".
            let arr = recordResult[arrName];
            if (Array.isArray(arr) && arr.length > 0) {
                return arr[0][elemName];
            }
        } 
    }
    else if (recordMapping.select) {
        // Result must be a single value
        return recordResult[recordMapping.select];
    }

    // If we get here, we have no explicit way of mapping to a result value, so just return the resultRecord
    return recordResult;
}

/**
 * Retrieve an entity value based on the given initial entity name and ID,
 * along with a mappings object which directs how the record retrieval is
 * mapped to a result value.
 * 
 * Use this method to chain together record retrievals when the result value
 * is referenced by another entity (the starting point).
 * @param {[*]} mappings 
 * @param {string} entityId 
 * @returns {Promise<any>}
 */
async function retrieveRecordAsync(mappings, entityId) {
    let mapping = mappings[0];
    let options = "";

    if (mapping.select) {
        options += (options.length === 0) ? "?" : "&";
        options += "$select=" + mapping.select;
    }
    if (mapping.expand) {
        options += (options.length === 0) ? "?" : "&";
        options += "$expand=" + mapping.expand;
    }

    const entityLogicalName = mapping.entityLogicalName;

    let result = await Xrm.WebApi.retrieveRecord(entityLogicalName, entityId, options);
    let resultValue = mapResultToValue(result, mapping);

    if (mappings.length === 1) {
        return resultValue;
    }

    return await retrieveRecordAsync(mappings.slice(1), resultValue);
}

/**
 * Get the email address associated with the given entity. For any entity other than systemuser,
 * we will drill down into "child" entities until we find the corresponding systemuser and their email.
 * 
 * The entityType may be one of:
 *      msdyn_workorder, bookableresourcebooking, bookableresource, or systemuser
 * @param {string} entityType 
 * @param {string} entityId 
 * @param {string} callTargetType "fieldtech" | "expert" | null | undefined
 * @returns {Promise<string>} The email address.
 */
function getEmailAddressAsync(entityType, entityId, callTargetType) {
    let mappings = [];
    switch (entityType) {
        case "msdyn_workorder":
            mappings = callTargetType === "fieldtech" ? WorkOrder2FieldTechEmail : WorkOrder2RemoteExpertEmail;
            break;
        case "bookableresourcebooking":
            mappings = BookableResourceBooking2Email;
            break;
        case "bookableresource":
            mappings = BookableResource2Email;
            break;
        case "systemuser":
            mappings = SystemUser2Email;
            break;
        case "contact":
            mappings = Contact2Email;
            break;
    }

    return retrieveRecordAsync(mappings, entityId);
}

/**
 * Get the client OS platform type
 * @return the OS platform
 */
function getPlatformType() {
    var osFamily = platform && platform.os ? platform.os.family : "";
    switch(osFamily) {
        case "Windows":
            return PlatformTypes.PC;
        case "Android":
            return PlatformTypes.Android;
        case "iOS":
            return PlatformTypes.iOS;
        default:
            return PlatformTypes.Unknown;
    }
}

/**
 * Retrieve the Customer Asset associated with the given Work Order.
 * @param {*} workOrderId ID of the Work Order for which the primary Customer Asset is requested.
 */
async function getWorkOrderAssetAsync(workOrderId) {
    const result = await Xrm.WebApi.online.retrieveRecord("msdyn_workorder", workOrderId, "?$select=_msdyn_customerasset_value");
    return result["_msdyn_customerasset_value"];
}

/**
 * Build the launch request body parameter object.
 * @param {string} callerEmail Email address of user initiating the call
 * @param {string} calleeEmail Email address of user contact to call
 * @param {object} metadataItems List of Key-value pair metadata items
 * @return launch request body parameter object
 */
function buildLaunchRequestData(callerEmail, calleeEmail, metadataItems) {
    var platformType = getPlatformType();
    return {
        email: callerEmail,
        platform: platformType,
        calleeEmail: calleeEmail,
        metadataItems: metadataItems
    };
}

/**
 * Navigate to the given URL using the Dynamics Xrm.Navigation API
 * @param {string} url URL to navigate to
 */
function navigateToUrl(url) {
    if (xrm) {
        xrm.Navigation.openUrl(url);
    }
}

/**
 * Open the given URL in a new browser tab
 * @param {string} url URL to open in a new browser tab
 */
function openNewTab(url) {
    var clientType = clientContext.getClient();
    if (clientType === "Mobile") {
        // the Dynamics mobile application does not support opening
        // new tabs in a browser using window.open
        navigateToUrl(url);
    }
    else {
        window.open(url, '_blank');
    }
}

/**
 * Get the entity which triggered the button click event. This could be the primary
 * entity (if the main form command button was clicked) or a sub-entity if the button
 * is located within a subgrid (such as the Bookable Resources subgrid section within
 * the Work Order main form).
 * @param {*} primaryControl 
 */
function getTriggeringEntity(primaryControl) {
    // If there's a grid associated w/the given control, we have a list
    // of entities and must pick the first selected one.
    const grid = primaryControl.getGrid && primaryControl.getGrid();
    if (grid) {
        const selectedEntities = grid.getSelectedRows().getAll();
        return (selectedEntities.length === 0) ? null : selectedEntities[0];
    }

    // Otherwise we're dealing with a top-level (MainForm) entity
    return primaryControl.data.entity;
}

/**
 * Launch Onsight Connect, optionally including Onsight user and callee information
 * in the launch request.
 * 
 * This is the main hook for all Onsight Connect command buttons.
 * @param {any} primaryControl
 * @param {bool} callTargetType the entity type of the target callee. If null or undefined,
       the default target type will be used, based on the current context. For Work Orders,
       this can be set to either "fieldtech" (to call the assigned field tech/bookable resource booking)
       or to "expert" (to call the remote expert/support contact).
 * @param {bool} includeContactInfo true to include callee info in request
 * @return none
 */
async function launchOnsightConnect(primaryControl, callTargetType) {
    // Get the entity in context (could be a selected entity within a subgrid, for example)
    const contextEntity = getTriggeringEntity(primaryControl);
    if (!contextEntity) {
        // Skipping launch; no entity in context
        return;
    }

    var platformType = getPlatformType();
    var clientType = clientContext.getClient();
    const callerEmail = await getEmailAddressAsync("systemuser", xrm.Page.context.getUserId());
    let emailTargetKey = contextEntity._entityType;

    let metadata = {};
    if (primaryControl.entityReference) {
        // Get the page's top-level entity, which may differ from the contextEntity above
        const topLevelEntity = primaryControl.entityReference;
        if (topLevelEntity.entityType === "msdyn_workorder") {
            metadata["WorkOrder"] = removeBraces(topLevelEntity.id);
            metadata["Asset"] = await getWorkOrderAssetAsync(topLevelEntity.id);
        }
    }

    let calleeEmail = null;
    if (emailTargetKey && callTargetType) {
        // Map a generic contact (eg, "remote expert" or "fieldtech") to an actual person's email address
        calleeEmail = await getEmailAddressAsync(emailTargetKey, contextEntity._entityId.guid, callTargetType);
    }

    const launchRequestData = buildLaunchRequestData(callerEmail, calleeEmail, metadata);
    if (clientType === "Mobile" && platformType == PlatformTypes.Android) {
        // can't launch directly by setting the window location or navigating
        // using an Android intent uri (intent://) or custom uri (onsightconnect://).
        // Go to the intermediate launch page
        launchOCPage(true, true, launchRequestData);
    }
    else {
        launchOCAjax(launchRequestData);
    }
}

/**
 * Launch Onsight connect by opening the intermediate OPM launch page
 * @param {bool} includeUserInfo true to include current user info in request
 * @param {bool} includeContactInfo true to include callee info in request
 * @param {object} launchRequestData Onsight Connect launch request parameters
 * @return none
 */
async function launchOCPage(includeUserInfo, includeContactInfo, launchRequestData) {
    const apiKey = await getApiKeyAsync();

    var url = opmLaunchPage + "?mlaunch";
        url += "&ak=" + encodeURIComponent(apiKey);
        if (includeUserInfo) url += "&u=" + encodeURIComponent(launchRequestData.username);
        if (includeUserInfo) url += "&p=" + encodeURIComponent(launchRequestData.password);
        if (includeContactInfo) url += "&ce=" + encodeURIComponent(launchRequestData.calleeEmail);
        url += "&m=" + encodeURIComponent(JSON.stringify(launchRequestData.metadataItems));
    openNewTab(url);
}

/**
 * Launch Onsight Connect by calling the OPM REST API and using the
 * resulting URL to launch tha application directly
 * @param {object} launchRequestData Onsight Connect launch request parameters
 * @return none
 */
function launchOCAjax(launchRequestData) {   
    sendOCApiLaunchRequest(launchRequestData).then((launchUri) => {
        if (getPlatformType() == PlatformTypes.iOS || 
            getPlatformType() == PlatformTypes.Android) {
            navigateToUrl(launchUri)
        }
        else {
            // use the protocol check helper library to launch the
            // app with the custom URI scheme in a browser-specific way
            protocolCheck(launchUri,
                () => {
                    //xrm.Navigation.openAlertDialog( { text: "Failed to launch Onsight" });
                    //console.log("OC failed to launch, is it installed?");
                },
                () => {
                    //console.log("OC should have launched");
            });
        }
    },
    (reason) => {
        //console.log("launchOCAjax: failed to get launch URI, error: " + reason);
    });
}

/**
 * Use XHR to send launch request to OPM REST API
 * @param {object} launchRequestBodyData Launch request body parameter
 * @return {Promise} A Promise the will resolve the OC launch URL when fulfilled 
 */
async function sendOCApiLaunchRequest(launchRequestBodyData) {
    const apiKey = await getApiKeyAsync();
    return new Promise((resolve, reject) => {
        var launchRequest = new XMLHttpRequest();
        launchRequest.addEventListener("readystatechange", function() {
            if (this.readyState == XMLHttpRequest.DONE) {
                if (this.status == 200) {
                    var launchUri = this.responseText.replace(/"/g, '');
                    resolve(launchUri);
                }
                else if (this.status == 400) {
                    var errorResult = JSON.parse(this.responseText);
                    reject(errorResult);
                }
            }
        });
        launchRequest.open("POST", apiEndpointUrl);
        launchRequest.setRequestHeader("Authorization", "ls Bearer: " + apiKey);
        launchRequest.setRequestHeader("Content-Type", "application/json");
        var launchRequestBody = JSON.stringify(launchRequestBodyData);
        launchRequest.send(launchRequestBody);
    });
}

/**
 * Open Workspace in a new tab. Search for assets by given full text term.
 * @return none
 */
async function searchWorkspace(primaryControl) {
    const entity = getTriggeringEntity(primaryControl);
    if (!entity) {
        // Skipping launch; no entity in context
        return;
    }

    // Use the primary entity's ID as the Workspace search term.
    // IOW, use a WorkOrder ID or Asset ID, depending on the current main form.
    let searchTerm = removeBraces(entity._entityId.guid);

    // From the Customer Asset page, include any and all related Work Order Ids in the Workspace search
    if (entity._entityType === "msdyn_customerasset") {
        let relatedWOsResponse = await Xrm.WebApi.online.retrieveRecord("msdyn_customerasset", entity._entityId.guid, "?$expand=msdyn_msdyn_customerasset_msdyn_workorder_CustomerAsset($select=msdyn_workorderid)");
        var msdyn_customerassetid = relatedWOsResponse["msdyn_customerassetid"];
        for (let i = 0; i < relatedWOsResponse.msdyn_msdyn_customerasset_msdyn_workorder_CustomerAsset.length; i++) {
            var relatedWorkOrderId = relatedWOsResponse.msdyn_msdyn_customerasset_msdyn_workorder_CustomerAsset[i]["msdyn_workorderid"];
            searchTerm += " " + relatedWorkOrderId;
        }
    }

    var wsUrl = workspaceUrl + "?f=" + encodeURIComponent("\"" + searchTerm + "\"");
    openNewTab(wsUrl);
}

/**
 * Utility method to test different navigation methods from Dynamics 365
 * browser and mobile client applications
 * @param {string} navType Navigation method (WindowLocation, NewTab, XrmNavigation)
 * @param {string} locationType Destination type (IntentOrUniversal, CustomUri, or External)
 * @return none
 */
function navigationTest(navType, locationType) {
    if (getPlatformType() == PlatformTypes.iOS) {
        setTimeout(() => navigationTestInternal(navType, locationType), 2500);
    }
    else {
        navigationTestInternal(navType, locationType);
    }
}

/**
 * Utility method to test different navigation methods from Dynamics 365
 * browser and mobile client applications
 * @param {string} navType Navigation method (WindowLocation, NewTab, XrmNavigation)
 * @param {string} locationType Destination type (IntentOrUniversal, CustomUri, or External)
 */
function navigationTestInternal(navType, locationType) {
    var uri = opmLaunchPage;
    switch (locationType) {
        case "IntentOrUniversal":
        uri = getPlatformType() == PlatformTypes.Android ? intentUri : universalLinkBase;
        break;
        case "CustomUri":
        uri = customUri;
        break;
        case "External":
        uri = externalUri;
        break;
        default:
        uri = opmLaunchPage;
    }

    switch (navType) {
        case "WindowLocation":
            window.location = uri;
        break;
        case "NewTab":
            window.open(uri, "_blank");
        break;
        case "XrmNavigation":
            navigateToUrl(uri);
        break;
    }
}
