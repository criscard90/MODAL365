class Modal365 {
    constructor() {
        this.modal = null;
    }

    async loadScript(url) {
        let response = await fetch(url);
        let script = await response.text();
        return script;
    }

    async initialize() {
        let scriptUrlSHEETJS = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
        var scriptSHEETJS = await this.loadScript(scriptUrlSHEETJS);
        const scriptElement = document.createElement('script');
        scriptElement.textContent = scriptSHEETJS;
        document.body.appendChild(scriptElement);
    }

    createModal(config) {
        let modal = document.createElement('div');
        modal.style.width = '50%';
        modal.style.position = 'fixed';
        modal.style.top = '50%';
        modal.style.left = '50%';
        modal.style.transform = 'translate(-50%, -50%)';
        modal.style.backgroundColor = 'white';
        modal.style.padding = '30px';
        modal.style.boxShadow = '0 0 10px rgba(0, 0, 0, 0.5)';
        modal.style.zIndex = '1000';

        let title = document.createElement('h2');
        title.textContent = config.title;
        title.style.marginBottom = '20px';
        modal.appendChild(title);

        modal.inputs = {};

        config.inputs.forEach(input => {
            let label = document.createElement('label');
            label.textContent = input.label;
            label.style.display = 'block';
            label.style.marginBottom = '5px';

            let inputElement = document.createElement('input');
            inputElement.type = 'text';
            inputElement.style.display = 'block';
            inputElement.style.marginBottom = '20px';

            modal.inputs[input.label] = inputElement;

            modal.appendChild(label);
            modal.appendChild(inputElement);
        });

        let executeButton = document.createElement('button');
        executeButton.textContent = config.executeButtonText;
        executeButton.style.display = 'inline';
        executeButton.style.marginTop = '20px';
        modal.executeButton = executeButton;
        modal.appendChild(executeButton);

        let closeButton = document.createElement('button');
        closeButton.textContent = 'Chiudi';
        closeButton.style.display = 'inline';
        closeButton.style.marginTop = '10px';
        closeButton.style.marginLeft = '10px';
        closeButton.onclick = () => document.body.removeChild(modal);
        modal.appendChild(closeButton);

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

        modal.appendChild(progressBarContainer);

        let progressInfo = document.createElement('div');
        progressInfo.style.marginTop = '10px';
        progressInfo.innerText = 'Pronto';
        modal.appendChild(progressInfo);

        let logContainer = document.createElement('div');
        logContainer.style.height = '150px';
        logContainer.style.marginTop = '10px';
        logContainer.style.maxHeight = '150px';
        logContainer.style.overflowY = 'auto';
        logContainer.style.border = '1px solid #ccc';
        logContainer.style.padding = '10px';
        modal.appendChild(logContainer);

        document.body.appendChild(modal);

        modal.progressSteps = config.progressSteps;
        modal.logContainer = logContainer;

        modal.executeButton.onclick = () => {
            let inputData = {};
            for (let key in modal.inputs) {
                inputData[key] = modal.inputs[key].value;
            }
            this.mainThread(inputData); // Chiama direttamente mainThread
        };

        this.modal = modal; // Salva il riferimento al modal creato
    }

    await(ms)   {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    createLogEntry(message) {
        let logEntry = document.createElement('div');
        logEntry.textContent = message;
        logEntry.style.marginTop = '5px';
        logEntry.style.color = '#333';
        this.modal.logContainer.appendChild(logEntry);
        this.modal.logContainer.scrollTop = this.modal.logContainer.scrollHeight;
    }

    async mainThread(inputData) {
        let totalPercentage = 0;
        this.modal.logContainer.innerHTML = "";

        for (let step of this.modal.progressSteps) {
            var result = await step.operation(inputData);
            step.results.push(result);
            totalPercentage += step.percentage;
            this.modal.querySelector('div > div').style.width = `${totalPercentage}%`;
            this.modal.querySelector('div > div').style.backgroundColor = "#00c98c";
            this.modal.querySelector('div > div + div').textContent = step.label;
        }
    }

    async backupRecord(entityName, recordId) {
        try {
            // Recupera l'ID del record e il nome dell'entità dalla pagina corrente
            
            if (!entityName || !recordId) {
                alert("Impossibile ottenere ID o entità dalla pagina corrente.");
                return;
            }
            
            // Ottieni il client URL (token di base)
            const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();
            
            // 1. Ottieni i dettagli del record principale
            const mainRecordUrl = `${clientUrl}/api/data/v9.2/${entityName}s(${recordId})`;
            const mainRecord = await fetch(mainRecordUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json",
                    "OData-MaxVersion": "4.0",
                    "OData-Version": "4.0"
                }
            }).then(res => res.json());
            
            // 2. Ottieni le relazioni 1:N dell'entità
            const metadataUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityName}')?$expand=OneToManyRelationships`;
            const metadata = await fetch(metadataUrl, {
                method: "GET",
                headers: {
                    "Accept": "application/json",
                    "OData-MaxVersion": "4.0",
                    "OData-Version": "4.0"
                }
            }).then(res => res.json());
            
            const relationships = metadata.OneToManyRelationships;
            if (!relationships || relationships.length === 0) {
                alert("Nessuna relazione 1:N trovata per questa entità.");
                return;
            }
            
            // 3. Prepara un workbook Excel
            const workbook = XLSX.utils.book_new();
            
            // Aggiungi i dati del record principale come primo foglio
            const mainSheet = XLSX.utils.json_to_sheet([mainRecord]);
            XLSX.utils.book_append_sheet(workbook, mainSheet, "Main Record");
            
            // 4. Itera su tutte le relazioni 1:N e recupera i dati
            for (const relationship of relationships) {
                const navigationProperty = relationship.ReferencedEntityNavigationPropertyName;
                const relatedEntity = relationship.ReferencingEntity;
                
                console.log(navigationProperty);
                
                if (navigationProperty) {
                    const relatedRecordsUrl = `${clientUrl}/api/data/v9.2/${entityName}s(${recordId})/${navigationProperty}`;
                    const relatedRecords = await fetch(relatedRecordsUrl, {
                        method: "GET",
                        headers: {
                            "Accept": "application/json",
                            "OData-MaxVersion": "4.0",
                            "OData-Version": "4.0"
                        }
                    }).then(res => res.json());
                    
                    // Se ci sono record correlati, aggiungili come un foglio
                    if (relatedRecords && relatedRecords.value && relatedRecords.value.length > 0) {
                        const relatedSheet = XLSX.utils.json_to_sheet(relatedRecords.value);
                        XLSX.utils.book_append_sheet(workbook, relatedSheet, relatedEntity.substring(0, 30));
                    }
                }
            }
            
            // 5. Scarica il file Excel
            XLSX.writeFile(workbook, `${entityName}_${recordId}_backup.xlsx`);
        } catch (error) {
            console.error("Errore durante il backup:", error);
            alert("Si è verificato un errore. Controlla la console per maggiori dettagli.");
        }
        
    }

    async executeRetrieve(entity, options) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + options;

        const req = new XMLHttpRequest();
        req.open("GET", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Prefer", "odata.include-annotations=\"OData.Community.Display.V1.FormattedValue\"");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 200 || req.status === 201) {
                    resolve(JSON.parse(req.responseText).value);
                } else {
                    console.error('Errore nella richiesta API:', req.status, req.statusText);
                    reject(new Error('Errore nella richiesta API'));
                }
            };

            req.onerror = () => {
                console.error('Errore nella richiesta API:', req.status, req.statusText);
                reject(new Error('Errore nella richiesta API'));
            };

            try {
                req.send();
            } catch (error) {
                console.error('Errore nella richiesta API:', error);
                reject(error);
            }
        });
    }

    async executeDelete(entity, id) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + "('" + id + "')";

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
                    console.error('Errore nella richiesta API:', req.status, req.statusText);
                    reject(new Error('Errore nella richiesta API'));
                }
            };

            req.onerror = () => {
                console.error('Errore nella richiesta API:', req.status, req.statusText);
                reject(new Error('Errore nella richiesta API'));
            };

            try {
                req.send();
            } catch (error) {
                console.error('Errore nella richiesta API:', error);
                reject(error);
            }
        });
    }

    async executeCreate(entity, data) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity;

        const req = new XMLHttpRequest();
        req.open("POST", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json;odata.metadata=minimal");
        req.setRequestHeader("Content-Type", "application/json;odata.metadata=minimal");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 201) {
                    resolve(JSON.parse(req.responseText));
                } else {
                    console.error('Errore nella richiesta API:', req.status, req.statusText);
                    reject(new Error('Errore nella richiesta API'));
                }
            };

            req.onerror = () => {
                console.error('Errore nella richiesta API:', req.status, req.statusText);
                reject(new Error('Errore nella richiesta API'));
            };

            try {
                req.send(JSON.stringify(data));
            } catch (error) {
                console.error('Errore nella richiesta API:', error);
                reject(error);
            }
        });
    }

    async executeUpdate(entity, id, data) {
        const url = Xrm.Utility.getGlobalContext().getClientUrl() + "/api/data/v9.2/" + entity + "('" + id + "')";

        const req = new XMLHttpRequest();
        req.open("PATCH", url, true);
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.setRequestHeader("Accept", "application/json;odata.metadata=minimal");
        req.setRequestHeader("Content-Type", "application/json;odata.metadata=minimal");

        return new Promise((resolve, reject) => {
            req.onload = () => {
                if (req.status === 204) {
                    resolve(true);
                } else {
                    console.error('Errore nella richiesta API:', req.status, req.statusText);
                    reject(new Error('Errore nella richiesta API'));
                }
            };

            req.onerror = () => {
                console.error('Errore nella richiesta API:', req.status, req.statusText);
                reject(new Error('Errore nella richiesta API'));
            };

            try {
                req.send(JSON.stringify(data));
            } catch (error) {
                console.error('Errore nella richiesta API:', error);
                reject(error);
            }
        });
    }
}
