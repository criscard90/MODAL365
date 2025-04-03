class Modal365 {
    constructor(config) {
        this.modal = document.createElement('div');
        this.modal.style.width = '50%';
        this.modal.style.position = 'fixed';
        this.modal.style.top = '50%';
        this.modal.style.left = '50%';
        this.modal.style.transform = 'translate(-50%, -50%)';
        this.modal.style.backgroundColor = 'white';
        this.modal.style.padding = '30px';
        this.modal.style.boxShadow = '0 0 10px rgba(0, 0, 0, 0.5)';
        this.modal.style.zIndex = '1000';

        let title = document.createElement('h2');
        title.textContent = config.title;
        title.style.marginBottom = '20px';
        this.modal.appendChild(title);

        this.inputs = {};
        config.inputs.forEach(input => {
            let label = document.createElement('label');
            label.textContent = input.label;
            label.style.display = 'block';
            label.style.marginBottom = '5px';

            let inputElement;
            if (input.type === 'select') {
                // Create a dropdown (optionset)
                inputElement = document.createElement('select');
                input.options.forEach(option => {
                    let optionElement = document.createElement('option');
                    optionElement.value = option.value;
                    optionElement.textContent = option.label;
                    inputElement.appendChild(optionElement);
                });
            } else {
                // Default to a text input
                inputElement = document.createElement('input');
                inputElement.type = input.type || 'text'; // Default to 'text' if no type is provided
            }

            inputElement.style.display = 'block';
            inputElement.style.marginBottom = '20px';

            this.inputs[input.label] = inputElement;

            this.modal.appendChild(label);
            this.modal.appendChild(inputElement);
        });

        let executeButton = document.createElement('button');
        executeButton.textContent = config.executeButtonText;
        executeButton.style.display = 'inline';
        executeButton.style.marginTop = '20px';
        this.modal.executeButton = executeButton;
        this.modal.appendChild(executeButton);

        let closeButton = document.createElement('button');
        closeButton.textContent = 'Chiudi';
        closeButton.style.display = 'inline';
        closeButton.style.marginTop = '10px';
        closeButton.style.marginLeft = '10px';
        closeButton.onclick = () => document.body.removeChild(this.modal);
        this.modal.appendChild(closeButton);

        let progressBarContainer = document.createElement('div');
        progressBarContainer.style.width = '100%';
        progressBarContainer.style.backgroundColor = '#e0e0e0';
        progressBarContainer.style.marginTop = '20px';
        progressBarContainer.style.marginBottom = '10px';

        let progressBar = document.createElement('div');
        progressBar.style.width = '0%';
        progressBar.style.height = '20px';
        progressBar.style.backgroundColor = '#76c7c0';
        progressBarContainer.appendChild(progressBar);

        this.modal.appendChild(progressBarContainer);

        let progressInfo = document.createElement('div');
        progressInfo.style.marginTop = '10px';
        progressInfo.innerText = 'Pronto';
        this.modal.appendChild(progressInfo);

        let logContainer = document.createElement('div');
        logContainer.style.height = '150px';
        logContainer.style.marginTop = '10px';
        logContainer.style.maxHeight = '150px';
        logContainer.style.overflowY = 'auto';
        logContainer.style.border = '1px solid #ccc';
        logContainer.style.padding = '10px';
        this.modal.appendChild(logContainer);

        document.body.appendChild(this.modal);

        this.progressSteps = config.progressSteps;
        this.logContainer = logContainer;

        this.modal.executeButton.onclick = () => {
            let inputData = {};
            for (let key in this.inputs) {
                inputData[key] = this.inputs[key].value;
            }
            this.mainThread(inputData);
        };
    }

    async mainThread(inputData) {
        let totalPercentage = 0;
        this.logContainer.innerHTML = "";

        for (let step of this.progressSteps) {
            await step.operation(inputData);
            totalPercentage += step.percentage;
            this.modal.querySelector('div > div').style.width = `${totalPercentage}%`;
            this.modal.querySelector('div > div').style.backgroundColor = "#00c98c";
            this.modal.querySelector('div > div + div').textContent = step.label;
        }
    }
}

// Define the Utility class as a static property of Modal365
Modal365.Utility = class {
    static async loadScript(url) {
        let response = await fetch(url);
        let script = await response.text();
        return script;
    }

    static async await(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    static createLogEntry(message, logContainer) {
        let logEntry = document.createElement('div');
        logEntry.textContent = message;
        logEntry.style.marginTop = '5px';
        logEntry.style.color = '#333';
        logContainer.appendChild(logEntry);
        logContainer.scrollTop = logContainer.scrollHeight;
    }
};

// Define the CRUD class as a static property of Modal365
Modal365.CRUD = class {
    static async executeRetrieve(entity, options) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + options;
        const req = new XMLHttpRequest();
        req.open("GET", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Prefer", 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"');

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 200 || req.status === 201) {
                    resolve(JSON.parse(req.responseText).value);
                } else {
                    reject(new Error('Error in API request'));
                }
            };

            req.onerror = () => reject(new Error('Error in API request'));

            try {
                req.send();
            } catch (error) {
                reject(error);
            }
        });
    }

    static async executeCreate(entity, data) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + "s";
        const req = new XMLHttpRequest();
        req.open("POST", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 200 || req.status === 201) {
                    resolve(JSON.parse(req.responseText));
                } else {
                    reject(new Error('Error in API request'));
                }
            };

            req.onerror = () => reject(new Error('Error in API request'));

            try {
                req.send(JSON.stringify(data));
            } catch (error) {
                reject(error);
            }
        });
    }

    static async executeDelete(entity, id) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + "s(" + id + ")";
        const req = new XMLHttpRequest();
        req.open("DELETE", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 204) {
                    resolve(true);
                } else {
                    reject(new Error('Error in API request'));
                }
            };

            req.onerror = () => reject(new Error('Error in API request'));

            try {
                req.send();
            } catch (error) {
                reject(error);
            }
        });
    }

    static async executeUpdate(entity, id, data) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + "s(" + id + ")";
        const req = new XMLHttpRequest();
        req.open("PATCH", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 200 || req.status === 204) {
                    resolve(true);
                } else {
                    reject(new Error('Error in API request'));
                }
            };

            req.onerror = () => reject(new Error('Error in API request'));

            try {
                req.send(JSON.stringify(data));
            } catch (error) {
                reject(error);
            }
        });
    }
};

export default Modal365;
