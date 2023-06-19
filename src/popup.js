let currentDomain;
let userInput;

async function getCurrentUrl() {
    let queryOptions = { active: true, lastFocusedWindow: true };
    // `tab` will either be a `tabs.Tab` instance or `undefined`.
    let [tab] = await chrome.tabs.query(queryOptions);
    return tab.url;
}

/***
 * Returns the top level domain if a dynamics environment is opened in current tab
 * Otherwise it returns am empty string if on unmatching domain
 */
function parseEnvironmentFromUrl(url) {
    let environment = "";
    console.log(url);
    const expectedSuffix = '.dynamics.com';
    const domainStartIndex = url.indexOf(expectedSuffix);
    if (domainStartIndex !== -1) {
        environment = url.substring(0, domainStartIndex + expectedSuffix.length);
        console.log(environment);
    }
    return environment;
}

/***
 * Show error message and disable the button form
 */
function showDomainError(document) {
    const environmentSection = document.getElementById('environmentSection');
    const lineBreak = document.createElement('br');
    environmentSection.append(lineBreak);

    const wrongTabMessage = document.createElement('span');
    wrongTabMessage.innerText = 'Dynamics environment needs to be open in the current tab';
    wrongTabMessage.setAttribute('class', 'warning');
    environmentSection.append(wrongTabMessage);

    document.getElementById('fetch-button').disabled = true;
}

function onEnvironmentUpdate(event) {
    environmentInput = document.getElementById('environment');
    userInput = environmentInput.value;
    if (userInput.includes('dynamics.com')) {
        try {
            const existingTabMessage = document.getElementsByClassName('warning');
            existingTabMessage[0].remove();
            document.getElementById('fetch-button').disabled = false;
            currentDomain = userInput;
        } catch (error) {
            console.log(error);
        }
    } else {
        try {
            const existingTabMessage = document.getElementsByClassName('warning');
            if (existingTabMessage.length < 1) {
                showDomainError(document);
            }
        } catch (error) {
            console.log(error);
        }
    }
}

function removeFormLists() {
    const existingLists = document.getElementsByTagName('select');
    if (existingLists != null) {
        for (list of existingLists) {
            list.remove();
        }
    }
}

function removeFormTextBox() {
    const textarea = document.getElementById('field-text-box');
    if (textarea != null) {
        textarea.remove();
    }
}

function removeResultMessages() {
    const existingResultMessage = document.getElementById('result-message');
    if (existingResultMessage != null) {
            message.remove();
    }
}

function appendSpinner(id, sectionId) {
    const spinner = document.createElement('div');
    spinner.setAttribute('id', id);
    spinner.setAttribute('class', 'loader');
    spinner.setAttribute('alt', 'loading-spinner');
    const formSection = document.getElementById(sectionId);
    formSection.append(spinner);
}

function removeSpinner(id) {
    const spinner = document.getElementById(id);
    spinner.remove();
}

async function populateFormListWithXML(forms, formList) {
    let dataFieldNames;
    for (let count = 0; count < forms.length; count++) {
        // create select list
        const element = document.createElement('option');
        element.text = forms[count].name;

        const parser = new DOMParser();
        const xmlDoc = parser.parseFromString(forms[count].formXml, 'text/xml');
        const tagsElement = xmlDoc.querySelector('tabs');
        const controlElements = tagsElement.querySelectorAll('control[datafieldname]');
        dataFieldNames = Array.from(controlElements).map(function (control) {
            return control.getAttribute('datafieldname');
        });
        // assign the item the full field list as value attribute 
        element.value = dataFieldNames;

        // This is why it isnt showin after initial fetch
        // need to update the the disPLAYED VALUE of the text box as the first item in the array 
        /*        if (count == 0) {
                    const fieldsTextBox = document.getElementById('textarea');
                    for (const name of dataFieldNames) {
                        fieldsTextBox.value = fieldsTextBox.value + name + "\n"
                    } */

        formList.add(element);
    }
    return formList;
}


document.addEventListener('DOMContentLoaded', async function () {
    const currentUrl = await getCurrentUrl();
    let currentDomain = parseEnvironmentFromUrl(currentUrl);

    const environmentInput = document.getElementById('environment');
    let userInput;

    environmentInput.addEventListener('input', onEnvironmentUpdate);

    if (currentDomain != "" || undefined && environment.value === "") {
        environmentInput.value = currentDomain;
    } else {
        showDomainError(document);
    }


    // form radio buttons listeners
    const requiredFieldsButton = document.getElementById('required-fields-input');
    requiredFieldsButton.addEventListener('click', async () => {
        removeFormLists();
        removeResultMessages();
        removeFormTextBox();
    });

    const formsButton = document.getElementById('form-fields-input');
    formsButton.addEventListener('click', async () => {
        removeFormLists();
        removeResultMessages();
        const spinnerId = 'spinner'
        appendSpinner(spinnerId, 'fields-section');

        const entityInput = document.getElementById('entity-input');
        const entityValue = entityInput.value;
        const formResult = await chrome.runtime.sendMessage({
            url: currentDomain,
            entity: entityValue,
            action: 'getForms'
        });

        removeSpinner(spinnerId);

        const fieldsTextBox = document.createElement('textarea');
        fieldsTextBox.setAttribute('id', 'field-text-box');

        const formList = document.createElement('select');
        formList.setAttribute('id', 'forms-list');
        formList.addEventListener('change', (event) => {
            fieldsTextBox.value = "";
            const fieldString = event.target.value;
            const fieldsArray = fieldString.split(',');
            for (const field of fieldsArray) {
                fieldsTextBox.value = fieldsTextBox.value + field + "\n"
            }
        })

        const xml = formResult.response;
        await populateFormListWithXML(xml, formList);

        const formSection = document.getElementById('fields-section');
        formSection.append(formList);
        formSection.append(fieldsTextBox);

        if (fieldsTextBox.value == "" || undefined) {
            const listValue = formList.value;
            fieldsTextBox.value = listValue.replaceAll(',', '\n')
        }
    })

    // post request listener
    const fetchButton = document.getElementById('fetch-button');
    fetchButton.addEventListener('click', async () => {

        const environmentInput = document.getElementById('environment');
        const userInput = environmentInput.value;


        console.log('current domain is: ' + currentDomain);
        if (currentDomain == "" || undefined && userInput == "" || undefined) {
            console.log("Incorrect domain, not sending request")
            //showDomainError();
        }

        if (currentDomain == "" || undefined && userInput != "" || undefined) {
            currentDomain = userInput;
        }

        try {
            const existingResultMessage = document.getElementById('result-message');
            existingResultMessage.remove();
        } catch {
            console.log('no existing mesage');
        }

        const requestSpinner = document.createElement('div');
        requestSpinner.setAttribute('id', 'spinner');
        requestSpinner.setAttribute('class', 'loader');

        const loadingSection = document.getElementById('loading');
        loadingSection.appendChild(requestSpinner);


        const entityInput = document.getElementById('entity-input');
        const entityValue = entityInput.value;

        const quantityInput = document.getElementById('quantity-input');
        const quantityValue = quantityInput.value;

        const requiredFieldsInput = document.getElementById('required-fields-input');
        const requiredOnly = requiredFieldsInput.checked;

        const formFieldsInput = document.getElementById('form-fields-input');
        const formFieldsList = document.getElementById('forms-list');
        let formDataFields;
        if (formFieldsInput.checked) {
            formDataFields = formFieldsList.value;
        }


        console.log('button is clicked, sending details to service worker: ' + currentDomain + entityValue + quantityValue);
        const result = await chrome.runtime.sendMessage({
            url: currentDomain,
            entity: entityValue,
            quantity: quantityValue,
            requiredOnly: requiredOnly,
            form: formDataFields,
            action: 'addRecords'
        });

        requestSpinner.remove();

        const resultMessage = document.createElement('span');
        resultMessage.setAttribute('id', 'result-message');

        resultMessage.innerHTML = result.response;
        const bottomSection = document.getElementById('bottom-section');
        bottomSection.append(resultMessage);
    });

});


