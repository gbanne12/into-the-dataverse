const webApiUrl = "/api/data/v9.2/";

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    (async function () {

        const environmentUrl = request.url;
        const entityName = request.entity;
        const recordsToAdd = request.quantity;

        if (request.action === 'getForms') {
            console.log('we will try to get forms');
            // Forms request
            try {
                const allFormsResponse = await fetch(`${environmentUrl}${webApiUrl}systemforms?$filter=objecttypecode eq '${entityName}' and type eq 2&$select=name,formid`);
                const allFormsJson = await allFormsResponse.json();
                const allformsArray = allFormsJson.value;

                const forms = getForms(allformsArray);

                for (const form of forms) {
                    const formResponse = await fetch(`${environmentUrl}${webApiUrl}systemforms(${form.formid})`);
                    const formJson = await formResponse.json();
                    const formXml = formJson["formxml"];
                    form.formXml = formXml;
                }

                sendResponse({ response: forms });
            } catch (error) {
                sendResponse({ response: `Unable to get forms: ${error.message}` });
            }
        }

        if (request.action === 'addRecords') {

            const isRequiredFieldsOnly = request.requiredOnly;
            const isFormFields = !request.requiredOnly && (request.form != "" || undefined);

            const requestBody = {};

            if (isRequiredFieldsOnly) {

                let requiredFields = [];
                // Metadata request
                try {
                    const metadataResponse = await fetch(`${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes`);

                    const metadataResponseJson = await metadataResponse.json();
                    if (metadataResponseJson.error != null || undefined) {
                        throw metadataResponseJson.error;
                    }

                    const metadataArray = metadataResponseJson.value;

                    //const validForCreateFields = getValidForCreateFields(metadataArray);
                    requiredFields = getRequiredFields(metadataArray);

                    for (const field of requiredFields) {
                        if (field.type === 'String') {
                            requestBody[field.value] = field.value;
                        }
                    }

                } catch (error) {
                    sendResponse({ response: `Unable to get required fields metadata: ${error.message}` });
                }

            } else if (isFormFields) {

                // same as above
                try {
                    const metadataResponse = await fetch(`${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes`);
                    const metadataResponseJson = await metadataResponse.json();
                    if (metadataResponseJson.error != null || undefined) {
                        throw metadataResponseJson.error;
                    }
                    const metadataArray = metadataResponseJson.value;

                    const allFields = request.form;
                    const formFields = allFields.split(',');
                    console.log("All fields to attempt to enter...." + formFields)


                    formFields.forEach(formField => {

                        const matchingField = metadataArray.find(field =>
                            field.LogicalName === formField &&
                            field.IsValidForCreate === true);

                        let attributeType;
                        try {
                            attributeType = matchingField.AttributeType;
                        } catch (error) {
                            attributeType = undefined;
                            console.log(`No attribute type, will not attempt to populate ${matchingField}`);
                        }


                        console.log("processing field..." + matchingField);
                        switch (attributeType) {
                            case undefined:
                                break;
                            case 'String':
                                if (matchingField.Format == 'Email') {
                                    requestBody[formField] = formField + Date.now() + "@gmail.com";
                                } else {
                                    requestBody[formField] = formField.slice(0, matchingField.MaxLength);
                                }
                                break;
                            case 'DateTime':
                                const dateTime = new Date().toISOString();
                                const dateOnly = dateTime.slice(0, 10)

                                if (matchingField.Format === 'DateOnly') {
                                    requestBody[formField] = dateOnly;
                                } else {
                                    requestBody[formField] = dateTime;
                                }
                                break;
                            case 'Boolean':
                                requestBody[formField] = Math.random() < 0.5;
                                break;
                            case 'Integer':
                                requestBody[formField] = Math.floor(Math.random() * (100) + 1);
                                break;
                            case 'Double':
                                requestBody[formField] = Math.random() * (100 - 1) + 1;
                                break;
                            case 'Picklist':
                                (async () => {
                                    const optionsSetEndpoint = `${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes(LogicalName='${matchingField.LogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=GlobalOptionSet($select=Options)`;
                                    const response = await fetch(optionsSetEndpoint);
                                    const json = await response.json();
                                    const label = json.GlobalOptionSet.Options[0].Value;
                                    requestBody[formField] = label;
                                })();
                                break;
                            case 'Lookup':
                                console.log(matchingField.Targets);
                                break;
                            case 'Money':
                                requestBody[formField] = generateRandomMoney(0, 500);
                                break;
                            case 'Memo':
                                requestBody[formField] = generateRandomMemo(140);
                            default:
                                console.log(`${matchingField.LogicalName} not any of the expected types.  Is a ${matchingField.AttributeType}; `)
                        }

                    });

                } catch (error) {
                    sendResponse({ response: `Unable to get form fields metadata: ${error.message}` });
                }

            } else {
                sendResponse({ response: 'Error: No form to populate' });
            }


            // POST request
            try {
                const entityDefinitions = await fetch(`${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')`);
                const entityDefinitionJson = await entityDefinitions.json();
                const tableDataName = entityDefinitionJson.LogicalCollectionName;

                console.log("********");
                console.log(requestBody);
                console.log("********");
                let postResponse;
                for (let count = 0; count < recordsToAdd; count++) {
                    postResponse = await postData(environmentUrl + webApiUrl + tableDataName, requestBody);
                }

                if (postResponse.status >= 200 && postResponse.status < 300) {
                    sendResponse({ response: `${postResponse.status} : Success, Record(s) were added for you` });
                } else {
                    sendResponse({ response: `Post request failed: ${postResponse.status}` });
                }

            } catch (error) {
                console.log(error);
                sendResponse({ response: `Error with the POST request: ${error.message}` });
            }
        }
    })();

    return true;
});

function getRequiredFields(metadataArray) {
    const requiredFields = [];
    for (let item = 0; item < metadataArray.length; item++) {

        // logic based on documentation 
        // https://github.com/MicrosoftDocs/powerapps-docs/blob/main/powerapps-docs/developer/data-platform/entity-attribute-metadata.md
        if (metadataArray[item].RequiredLevel.Value === "ApplicationRequired" &&
            metadataArray[item].IsRequiredForForm === true &&
            metadataArray[item].IsValidForCreate === true) {

            const fieldInfo = { value: metadataArray[item].LogicalName, type: metadataArray[item].AttributeType }
            requiredFields.push(fieldInfo);
        }
    }
    return requiredFields;
}



function getValidForCreateFields(metadataArray) {
    const validForCreateFields = [];
    for (let item = 0; item < metadataArray.length; item++) {

        if (metadataArray[item].IsValidForCreate === true) {
            const fieldInfo = { value: metadataArray[item].LogicalName, type: metadataArray[item].AttributeType };
            validForCreateFields.push(fieldInfo);
        }
    }
    return validForCreateFields;
}

function getForms(metadataArray) {
    const forms = [];
    for (let item = 0; item < metadataArray.length; item++) {
        const form = { name: metadataArray[item].name, formid: metadataArray[item].formid };
        forms.push(form);
    }
    return forms;
}


async function postData(url = "", data = {}) {
    const requestInit = {
        method: "POST",
        headers: {
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Content-Type": "application/json; charset=utf-8",
            "Accept": "application/json",
            "Prefer": "odata.include-annotations=*"
        },
        body: JSON.stringify(data)
    }
    const response = await fetch(url, requestInit);
    if (response.ok) {
        var uri = response.headers.get("OData-EntityId");
        var regExp = /\(([^)]+)\)/;
        var matches = regExp.exec(uri);
        var newId = matches[1];
        console.log(newId);
    } else {
        return response.json().then((json) => { throw json.error; });
    }
    return response;
}

function generateRandomMoney(min, max) {
    const decimalPrecision = 2; // Number of decimal places expected by the money field
    const randomValue = Math.random() * (max - min) + min;
    const roundedValue = randomValue.toFixed(decimalPrecision);
    return parseFloat(roundedValue);
}

function generateRandomMemo(length) {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    let result = '';
    const charactersLength = characters.length;

    for (let i = 0; i < length; i++) {
        result += characters.charAt(Math.floor(Math.random() * charactersLength));
    }

    return result;
}




