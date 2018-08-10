/**
 * config: {
 *     name: "visualName",
 *     displayName: "Visual Display Name",
 *     description: "Visual Description",
 *     guid: "visualNameAAFF...",
 *     version: "1.0.0",
 *     author: {
 *         name: "Author Name",
 *         email: "author@example.com"
 *     }
 * }
*/

let JSZip = require("jszip");
let resources = require("./resources.json");

function randomHEX32() {
    let s = "";
    for (let k = 0; k < 8; k++) {
        s += Math.floor(Math.random() * 2147483648).toString(16).toUpperCase().slice(-4);
    }
    return s;
}

class PowerBIVisualGenerator {
    constructor(template, containerScriptURL) {
        this.template = template;
        this.containerScriptURL = containerScriptURL;
    }

    getProperties() {
        return [
            { displayName: "Visual Name", name: "visualName", type: "string", default: "MyVisual" },
            { displayName: "Description", name: "description", type: "string", default: "" },
            { displayName: "Author Name", name: "authorName", type: "string", default: "Anonymous" },
            { displayName: "Author Email", name: "authorEmail", type: "string", default: "anonymous@example.com" }
        ];
    }

    getFileExtension() {
        return "pbiviz";
    }

    // Return a Promise<base64>
    generate(properties) {
        let template = this.template;
        let config = {
            name: properties.visualName,
            description: properties.description,
            guid: properties.visualName + randomHEX32(),
            author: {
                name: properties.authorName,
                email: properties.authorEmail
            }
        };

        let visual_json = {
            name: config.name,
            displayName: config.displayName,
            guid: config.guid,
            visualClassName: "CharticulatorPowerBIVisual",
            version: config.version,
            description: config.description,
            supportUrl: "",
            gitHubUrl: ""
        };

        for (let slot of template.dataSlots) {
            // Make sure we have a display name for each slot
            slot.powerBIName = slot.name.replace(/[^0-9a-zA-Z\_]/g, "");
            if (!slot.displayName) {
                slot.displayName = slot.powerBIName;
            }
            console.log("Data Slot: " + slot.displayName, slot.kind);
        }

        let dataViewMappingsConditions = {
        };
        let objectProperties = {};

        for (let slot of template.dataSlots) {
            dataViewMappingsConditions[slot.powerBIName] = { max: 1 };
        }

        for (let id in template.properties) {
            if (!template.properties.hasOwnProperty(id)) continue;
            for (let p of template.properties[id]) {
                let type = { text: true };
                switch (p.type) {
                    case "number":
                        type = { numeric: true };
                        break;
                }
                objectProperties[p.name] = {
                    displayName: p.displayName || p.name,
                    type: type
                };
            }
        }

        let capabilities = {
            dataRoles: [
                {
                    displayName: "Granularity (Level of Detail)",
                    name: "category",
                    kind: "Grouping"
                },
                ...template.dataSlots.map(slot => {
                    return {
                        displayName: slot.displayName,
                        name: slot.powerBIName,
                        kind: "Measure"
                    };
                })
            ],
            dataViewMappings: [
                {
                    conditions: [
                        dataViewMappingsConditions
                    ],
                    categorical: {
                        categories: {
                            for: {
                                in: "category"
                            }
                        },
                        values: {
                            select: template.dataSlots.map(slot => {
                                return {
                                    bind: { to: slot.powerBIName }
                                };
                            })
                        }
                    }
                }
            ],
            objects: {
                chartOptions: {
                    displayName: "Charticulator",
                    properties: objectProperties
                }
            }
        };

        let jsConfig = {
            visualGuid: config.guid,
            pluginName: config.guid,
            visualDisplayName: config.displayName,
            visualVersion: config.version,
            apiVersion: "1.6.0",
            templateData: template
        };
        let visual = resources.visual.replace(/\'\<\%\= *([0-9a-zA-Z\_]+) *\%\>\'/g, (x, a) => JSON.stringify(jsConfig[a]));

        let pbiviz_json = {
            visual: visual_json,
            apiVersion: "1.6.0",
            author: config.author,
            assets: {
                icon: "assets/icon.png"
            },
            externalJS: [],
            style: "style/visual.less",
            capabilities: capabilities,
            stringResources: [],
            content: {
                js: [
                    resources.libraries,
                    resources.container,
                    visual
                ].join("\n"),
                css: "",
                iconBase64: resources.icon
            }
        }

        let package_json = {
            version: config.version,
            author: config.author,
            resources: [
                {
                    resourceId: "rId0",
                    sourceType: 5,
                    file: "resources/" + config.guid + ".pbiviz.json"
                }
            ],
            visual: visual_json,
            metadata: {
                pbivizjson: {
                    resourceId: "rId0"
                }
            }
        }

        let zip = new JSZip();
        zip.file("package.json", JSON.stringify(package_json));
        let fresources = zip.folder("resources");
        fresources.file(config.guid + ".pbiviz.json", JSON.stringify(pbiviz_json));

        return zip.generateAsync({
            type: "base64",
            compression: "DEFLATE"
        });
    }
}

class PowerBIVisualBuilder {
    constructor(containerScriptURL) {
        this.containerScriptURL = containerScriptURL;
    }

    activate(context) {
        this.context = context;
        context.getApplication().registerExportTemplateTarget("PowerBI Custom Visual", (template) => new PowerBIVisualGenerator(template, this.containerScriptURL));
    }

    deactivate() {
        this.context.getApplication().mainStore.unregisterExportTemplateTarget("PowerBI Custom Visual");
    }
}

module.exports = {
    PowerBIVisualBuilder: PowerBIVisualBuilder,
    PowerBIVisualGenerator: PowerBIVisualGenerator
};