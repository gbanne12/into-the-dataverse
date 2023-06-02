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

    document.getElementById('fetchButton').disabled = true;
}


document.addEventListener('DOMContentLoaded', async function () {
    const currentUrl = await getCurrentUrl();
    const currentDomain = parseEnvironmentFromUrl(currentUrl)

    if (currentDomain != "" || undefined) {
        const environmentInput = document.getElementById('environment');
        environmentInput.value = currentDomain;
    } else {
        showDomainError(document);
    }


    // form radio buttons listeners
    const requiredFieldsButton = document.getElementById('required-fields-input');
    requiredFieldsButton.addEventListener('click', async () => {

        try {
            const allLists = document.getElementsByTagName('select');
            for (list of allLists) {
                list.remove();
            }
        } catch {
            console.log('no existing form list');
        }

        try {
            const textarea = document.getElementById('field-text-box');
            textarea.remove();
        } catch {
            console.log('no existing text area');
        }

        try {
            const existingResultMessage = document.getElementById('result-message');
            existingResultMessage.remove();
        } catch {
            console.log('no existing mesage');
        }
    });

    const formsButton = document.getElementById('form-fields-input');
    formsButton.addEventListener('click', async () => {

        const allLists = document.getElementsByTagName('select');
        for (list of allLists) {
            list.remove();
        }

        try {
            const existingResultMessage = document.getElementById('result-message');
            existingResultMessage.remove();
        } catch {
            console.log('no existing mesage');
        }

        const entityInput = document.getElementById('entity-input');
        const entityValue = entityInput.value;

        console.log('button is clicked attempting to get forms');

        const formSpinner = document.createElement('div');
        formSpinner.setAttribute('id', 'spinner');
        formSpinner.setAttribute('class', 'loader');
        const formSection = document.getElementById('fields-section');
        formSection.append(formSpinner);

        const result = await chrome.runtime.sendMessage({
            url: currentDomain,
            entity: entityValue,
            action: 'getForms'
        });

        formSpinner.remove();

        const forms = result.response;
        const formList = document.createElement('select');
        formList.setAttribute('id', 'forms-list');

        const fieldsTextBox = document.createElement('textarea');
        fieldsTextBox.setAttribute('id', 'field-text-box');

        formList.addEventListener('change', (event) => {
            fieldsTextBox.value = "";
            const fieldString = event.target.value;
            const fieldsArray = fieldString.split(',');
            for (const field of fieldsArray) {
                fieldsTextBox.value = fieldsTextBox.value + field + "\n"
            }
        })

        let dataFieldNames;
        for (let count = 0; count < forms.length; count++) {
            const element = this.createElement('option');
            element.text = forms[count].name;

            const parser = new DOMParser();
            const xmlDoc = parser.parseFromString(forms[count].formXml, 'text/xml');
            const tagsElement = xmlDoc.querySelector('tabs');
            const controlElements = tagsElement.querySelectorAll('control[datafieldname]');
            dataFieldNames = Array.from(controlElements).map(function (control) {
                return control.getAttribute('datafieldname');
            });
            element.value = dataFieldNames;

            if (count == 0) {
                for (const name of dataFieldNames) {
                    fieldsTextBox.value = fieldsTextBox.value + name + "\n"
                }
            }

            formList.add(element);
        }

        
        formSection.append(formList);

        formSection.append(fieldsTextBox);
        console.log('dataFieldNames:', dataFieldNames);

    })

    // post request listener
    const fetchButton = document.getElementById('fetch-button');
    fetchButton.addEventListener('click', async () => {

        if (currentDomain == "" || undefined) {
            console.log("Incorrect domain, not sending request")
            return;
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


