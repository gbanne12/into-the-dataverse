import './data-robot.js';

const webApiUrl = "/api/data/v9.2/";
let environmentUrl;
let entityName;

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    (async function () {

        environmentUrl = request.url;
        entityName = request.entity;
        const recordsToAdd = request.quantity;

        if (request.action === 'getForms') {
            try {
                const entityForms = await fetchForms(environmentUrl, entityName);

                for (const form of entityForms) {
                    form.formXml = await fetchFormXml(environmentUrl, form.formid)
                }
                sendResponse({ response: entityForms });

            } catch (error) {
                sendResponse({ response: `Unable to get forms: ${error.message}` });
            }
        }

        if (request.action === 'addRecords') {
            const isRequiredFields = request.requiredOnly;
            const isFormFields = !request.requiredOnly && (request.form != "" || undefined);
            const requestBody = {};

            if (isRequiredFields) {
                try {
                    const metadataArray = await fetchMetadata(environmentUrl, entityName);
                    const requiredFields = getRequiredFields(metadataArray);

                    for (const field of requiredFields) {
                        if (field.type === 'String') {
                            requestBody[field.value] = field.value;
                        }
                    }

                    const collectionName = await fetchLogicalCollectionName(environmentUrl, entityName);
                    for (let count = 0; count < recordsToAdd; count++) {
                        const response = await postData(environmentUrl + webApiUrl + collectionName, requestBody);
                        sendResponse({ response: `Post request sent: ${response} ` });
                    }

                } catch (error) {
                    sendResponse({ response: `Unable to add required fields: ${error.message}` });
                }

            } else if (isFormFields) {
                try {
                    const entityMetadata = await fetchMetadata(environmentUrl, entityName);
                    const userSelectedForm = request.form;
                    const formFields = userSelectedForm.split(',');

                    const fieldsForPostrequest = entityMetadata.filter(record => {
                        return formFields.includes(record.LogicalName) && record.IsValidForCreate === true;
                    });

                    for (let count = 0; count < fieldsForPostrequest.length; count++) {
                        const value = await getInputValueForField(fieldsForPostrequest[count]);
                        try {
                            const fieldName = fieldsForPostrequest[count].LogicalName;
                            const attributeType = fieldsForPostrequest[count].AttributeType;

                            const isCustomerField = attributeType === 'Customer';
                            const isLookupField = attributeType === 'Lookup';

                            if (value !== undefined) {
                                if (isCustomerField) {
                                    const referencedEntity = 'contact';
                                    const property = await buildODataValue(fieldName, referencedEntity, value);
                                    requestBody[property.key] = property.value;

                                } else if (isLookupField) {
                                    const referencedEntity = fieldsForPostrequest[count].Targets[0]
                                    const property = await buildODataValue(fieldName, referencedEntity, value);
                                    requestBody[property.key] = property.value;

                                } else {
                                    requestBody[fieldName] = value;
                                }
                            }

                        } catch (error) {
                            console.log(`no logicalname found for field ${JSON.stringify(fieldsForPostrequest[count])}, not including in request`);
                        }
                    }

                    const collectionName = await fetchLogicalCollectionName(environmentUrl, entityName);
                    for (let count = 0; count < recordsToAdd; count++) {
                        const response = await postData(environmentUrl + webApiUrl + collectionName, requestBody);
                        sendResponse({ response: `Post request sent: ${response} ` });
                    }

                } catch (error) {
                    sendResponse({ response: `Error with form field attempt : ${error.message}`, json: requestBody });
                }

            } else {
                sendResponse({ response: 'Error: No form was found. Pick one from the list.' });
            }
        }
    })();

    return true;
});

async function getInputValueForField(matchingField) {
    let value;
    let attributeType;
    try {
        attributeType = matchingField.AttributeType;
    } catch (error) {
        attributeType = undefined;
        console.log(`No attribute type, will not attempt to populate ${matchingField}`);
    }


    if (attributeType === undefined) {
        // do nothing
    } else if (attributeType === 'String') {
        const emailAddress = matchingField.LogicalName + Date.now() + "@gmail.com"
        const fieldName = matchingField.LogicalName.slice(0, matchingField.MaxLength);
        value = matchingField.Format == 'Email' ? emailAddress : fieldName;

    } else if (attributeType === 'DateTime') {
        const dateTime = new Date().toISOString();
        const dateOnly = dateTime.slice(0, 10);
        value = matchingField.Format === 'DateOnly' ? dateOnly : dateTime;

    } else if (attributeType === 'Boolean') {
        value = Math.random() < 0.5;

    } else if (attributeType === 'Integer') {
        value = Math.floor(Math.random() * (100) + 1);

    } else if (attributeType === 'Double') {
        value = Math.random() * (100 - 1) + 1;

    } else if (attributeType === 'Customer') {
        const id = await fetchLookupIdentifier(environmentUrl, 'contact');
        value = id;

    } else if (attributeType === 'Picklist') {
        const values = await fetchPicklistValues(matchingField);
        const randomIndex = 1;
        value = values[randomIndex].Value;

    } else if (attributeType === 'Lookup') {
        const lookupEntity = matchingField.Targets[0];
        const id = await fetchLookupIdentifier(environmentUrl, lookupEntity);
        return id === undefined ? undefined : id;

    } else if (attributeType === 'Money') {
        value = generateRandomMoney(0, 500);

    } else if (attributeType === 'Memo') {
        value = generateRandomMemo(140);

    } else {
        console.log(`${matchingField.LogicalName} not any of the expected types. Is a ${matchingField.AttributeType};`);
    }
    return value;
}


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

async function fetchForms(environmentUrl, entityName) {
    const allFormsResponse = await fetch(`${environmentUrl}${webApiUrl}systemforms?$filter=objecttypecode eq '${entityName}' and type eq 2&$select=name,formid`);
    const allFormsJson = await allFormsResponse.json();
    const allFormsArray = allFormsJson.value;

    const forms = [];
    for (let item = 0; item < allFormsArray.length; item++) {
        const form = { name: allFormsArray[item].name, formid: allFormsArray[item].formid };
        forms.push(form);
    }
    return forms;
}

async function fetchFormXml(environmentUrl, formId) {
    const formResponse = await fetch(`${environmentUrl}${webApiUrl}systemforms(${formId})`);
    const json = await formResponse.json();
    const formXml = json["formxml"];
    return formXml;
}

async function fetchMetadata(environmentUrl, entityName) {
    const metadataResponse = await fetch(`${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes`);

    const metadataResponseJson = await metadataResponse.json();
    if (metadataResponseJson.error != null || undefined) {
        throw metadataResponseJson.error;
    }
    const metadataArray = metadataResponseJson.value;
    return metadataArray;
}

async function fetchLookupIdentifier(environmentUrl, entityName) {
    const logicalCollectionName = await fetchLogicalCollectionName(environmentUrl, entityName);
    const response = await fetch(`${environmentUrl}${webApiUrl}${logicalCollectionName}`);
    const json = await response.json();
    const randomIndex = 1;
    const chosenRecord = json.value[randomIndex];
    if (chosenRecord === undefined) {
        return undefined;
    }
    const id = chosenRecord[entityName + 'id'];
    return id;
}

async function fetchNavigationPropertyName(environmentUrl, referencingAttribute, referencedEntity) {
    const relationshipMetadataUrl = `${environmentUrl}${webApiUrl}RelationshipDefinitions/Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata`;
    const relationshipDefinitions = await fetch(`${relationshipMetadataUrl}?$filter=ReferencingAttribute eq '${referencingAttribute}' and ReferencedEntity eq '${referencedEntity}'`);

    const json = await relationshipDefinitions.json();
    return json.value[0].ReferencingEntityNavigationPropertyName;
}

async function fetchOptionSet(environmentUrl, entityName, logicalName) {
    const optionsSetEndpoint = `${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes(LogicalName='${logicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=GlobalOptionSet($select=Options)`;
    const response = await fetch(optionsSetEndpoint);
    const json = await response.json();
    const optionSet = json.GlobalOptionSet.Options;
    return optionSet;
}

async function fetchLogicalCollectionName(environmentUrl, entityName) {
    const entityDefinitions = await fetch(`${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')`);
    const entityDefinitionJson = await entityDefinitions.json();
    const logicalCollectionName = entityDefinitionJson.LogicalCollectionName;
    return logicalCollectionName;
}

async function fetchPicklistValues(field) {
    const optionsSetEndpoint = `${environmentUrl}${webApiUrl}EntityDefinitions(LogicalName='${entityName}')/Attributes(LogicalName='${field.LogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=GlobalOptionSet($select=Options)`;
    const response = await fetch(optionsSetEndpoint);
    const json = await response.json();
    const values = json.GlobalOptionSet.Options;
    return values;
}

async function buildODataValue(fieldName, referencedEntity, value) {
    const navigationPropertyName = await fetchNavigationPropertyName(environmentUrl, fieldName, referencedEntity);
    const collectionName = await fetchLogicalCollectionName(environmentUrl, referencedEntity);
    return { key: `${navigationPropertyName}@odata.bind`, value: `/${collectionName}(${value})` };
}

/* Returns the new record Id if successfull, or error message if not  */
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
        const uri = response.headers.get("OData-EntityId");
        const regExp = /\(([^)]+)\)/;
        const matches = regExp.exec(uri);
        const newId = matches[1];
        console.log(newId);
        return newId;
    } else {
        const json = await response.json();
        return json.error.message;
    }
}


function generateRandomMoney(min, max) {
    const decimalPrecision = 2;
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