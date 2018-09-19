// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

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

/// <reference path="../api/app.bundle.d.ts" />

import * as JSZip from "jszip";

import * as Charticulator from "Charticulator";
import { ExportTemplateTarget } from "Charticulator/app/template";
import { Specification } from "Charticulator/core";

import {
  SchemaCapabilities,
  DataRole
} from "../api/v2.1.0/schema.capabilities";

interface Resources {
  icon: string;
  libraries: string;
  visual: string;
}

interface PowerBIColumn extends Specification.Template.Column {
  powerBIName: string;
}

interface PowerBIProperty extends Specification.Template.Property {
  powerBIName: string;
}

const resources: Resources = require("./resources.json");

function randomHEX32() {
  let s = "";
  for (let k = 0; k < 8; k++) {
    s += Math.floor(Math.random() * 2147483648)
      .toString(16)
      .toUpperCase()
      .slice(-4);
  }
  return s;
}

function getText(url: string) {
  return fetch(url, {
    credentials: "include"
  }).then(resp => resp.text());
}

class PowerBIVisualGenerator implements ExportTemplateTarget {
  constructor(
    public template: Specification.Template.ChartTemplate,
    public containerScriptURL: string
  ) {}

  public getProperties() {
    return [
      {
        displayName: "Visual Name",
        name: "visualName",
        type: "string",
        default: "MyVisual"
      },
      {
        displayName: "Description",
        name: "description",
        type: "string",
        default: ""
      },
      {
        displayName: "Author Name",
        name: "authorName",
        type: "string",
        default: "Anonymous"
      },
      {
        displayName: "Author Email",
        name: "authorEmail",
        type: "string",
        default: "anonymous@example.com"
      }
    ];
  }

  public getFileExtension() {
    return "pbiviz";
  }

  // Return a Promise<base64>
  public async generate(properties: { [name: string]: any }): Promise<string> {
    const template = this.template;
    const config = {
      name: properties.visualName,
      displayName: properties.visualName,
      description: properties.description,

      // Add a indicator that this came from Charticulator
      // also replace any non alphanumeric characters with _, spaces can screw up PBI
      guid: 'CHARTICULATOR_VISUAL_' + properties.visualName.replace(/[^A-Za-z0-9]+/, "_") + randomHEX32(),
      version: "1.0.0",
      author: {
        name: properties.authorName,
        email: properties.authorEmail
      }
    };

    const visual_json = {
      name: config.name,
      displayName: config.displayName,
      guid: config.guid,
      visualClassName: "CharticulatorPowerBIVisual",
      version: config.version,
      description: config.description,
      supportUrl: "",
      gitHubUrl: ""
    };

    const dataViewMappingsConditions: { [name: string]: any } = {};
    const objectProperties: { [name: string]: any } = {};

    // TODO: for now, we assume there's only one table
    const columns = template.tables[0].columns as PowerBIColumn[];

    for (const column of columns) {
      // Refine column names
      column.powerBIName = column.name.replace(/[^0-9a-zA-Z\_]/g, "_");
      dataViewMappingsConditions[column.powerBIName] = { max: 1 };
    }

    // Populate properties
    for (const property of template.properties as PowerBIProperty[]) {
      property.powerBIName = (property.objectID + property.displayName).replace(
        /[^0-9a-zA-Z\_]/g,
        "_"
      );
      let type: any = { text: true };
      switch (property.type) {
        case "number":
          {
            type = { numeric: true };
          }
          break;
      }
      objectProperties[property.powerBIName] = {
        displayName: property.displayName,
        type
      };
    }

    const capabilities: SchemaCapabilities = {
      dataRoles: [
        {
          displayName: "Granularity (Level of Detail)",
          name: "category",
          kind: "Grouping"
        },
        ...columns.map(column => {
          return {
            displayName: column.displayName,
            name: column.powerBIName,
            kind: "Measure"
          } as DataRole;
        })
      ],
      dataViewMappings: [
        {
          conditions: [dataViewMappingsConditions],
          categorical: {
            categories: {
              for: {
                in: "category"
              },
              dataReductionAlgorithm: {
                top: {
                  count: 30000 // That's the maximum
                }
              }
            },
            values: {
              select: columns.map(column => {
                return {
                  bind: { to: column.powerBIName }
                };
              })
            }
          }
        }
      ],

      /* Tell PBI to allow for sorting */
      sorting: {
        default: {}
      },
      objects: {
        chartOptions: {
          displayName: "Charticulator",
          properties: objectProperties
        }
      },
      // Declare that the visual supports highlight.
      // Power BI will give us a set of highlight values instead of filtering the data.
      supportsHighlight: true
    };

    const jsConfig: { [name: string]: any } = {
      visualGuid: config.guid,
      pluginName: config.guid,
      visualDisplayName: config.displayName,
      visualVersion: config.version,
      apiVersion: "2.1.0",
      templateData: template
    };
    const visual = resources.visual.replace(
      /[\'\"]\<\%\= *([0-9a-zA-Z\_]+) *\%\>[\'\"]/g,
      (x, a) => JSON.stringify(jsConfig[a])
    );

    const containerScript = await getText(this.containerScriptURL);

    const pbiviz_json = {
      visual: visual_json,
      apiVersion: "2.1.0",
      author: config.author,
      assets: {
        icon: "assets/icon.png"
      },
      externalJS: [] as string[],
      style: "style/visual.less",
      capabilities,
      stringResources: {},
      content: {
        js: [resources.libraries, containerScript, visual].join("\n"),
        css: "",
        iconBase64: resources.icon
      }
    };

    const package_json = {
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
    };

    const zip = new JSZip();
    zip.file("package.json", JSON.stringify(package_json));
    const fresources = zip.folder("resources");
    fresources.file(config.guid + ".pbiviz.json", JSON.stringify(pbiviz_json));

    return await zip.generateAsync({
      type: "base64",
      compression: "DEFLATE"
    });
  }
}

class PowerBIVisualBuilder {
  public context: Charticulator.ApplicationExtensionContext;

  constructor(public containerScriptURL: string) {}

  public activate(context: Charticulator.ApplicationExtensionContext) {
    this.context = context;
    context
      .getApplication()
      .registerExportTemplateTarget(
        "PowerBI Custom Visual",
        (template: Specification.Template.ChartTemplate) =>
          new PowerBIVisualGenerator(template, this.containerScriptURL)
      );
  }

  public deactivate() {
    this.context
      .getApplication()
      .unregisterExportTemplateTarget("PowerBI Custom Visual");
  }
}

module.exports = {
  PowerBIVisualBuilder,
  PowerBIVisualGenerator
};
