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
  DataRole,
  Objects
} from "../api/v2.1.0/schema.capabilities";
import { AttributeMap } from "Charticulator/core/specification";
import {
  LinksProperties,
  LinksObject
} from "Charticulator/core/prototypes/links";
import { boolean } from "Charticulator/core/expression";

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

const symbolTypes = Charticulator.Core.Prototypes.Marks.symbolTypesList;

class PowerBIVisualGenerator implements ExportTemplateTarget {
  constructor(
    public template: Specification.Template.ChartTemplate,
    public containerScriptURL: string
  ) {}

  public getProperties() {
    return [
      {
        displayName: "Enable Tooltip",
        name: "enableTooltip",
        type: "boolean",
        default: true
      },
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
      },
      {
        displayName: "Icon",
        name: "visualIcon",
        type: "file",
        default: {
          src: resources.icon,
          name: "default.png"
        }
      }
    ];
  }

  public getFileName(properties: { [name: string]: any }) {
    if (properties.visualName) {
      // Make sure filename doesn't contain illegal chars
      const filename = properties.visualName
        .replace(/[/\\?%*:|"<>]/g, "-")
        .trim();
      if (filename != "") {
        return filename + ".pbiviz";
      } else {
        // In case the user put nothing in filename
        return "charticulator.pbiviz";
      }
    } else {
      return "charticulator.pbiviz";
    }
  }

  // gets data table for links if links are anchored
  private hasAnchoredLinksAndTable(
    template: Charticulator.Core.Specification.Template.ChartTemplate
  ) {
    const link: LinksObject = template.specification.elements.find(element =>
      Boolean(
        element.classID === "links.table" &&
          element.properties.anchor1 &&
          element.properties.anchor2
      )
    ) as LinksObject;

    if (link) {
      const properties = link.properties as LinksProperties;
      const linkTableName = properties.linkTable && properties.linkTable.table;
      const linkTable = template.tables.find(
        table => table.name === linkTableName
      );

      return linkTable;
    }
  }

  private mapPropertyName(name: string) {
    switch (name) {
      case "fontSize": {
        return "Font size";
      }
      default:
        let words = name.replace(/([a-z])([A-Z])/g, "$1 $2").split(" ");
        words = words.map(w => w.toLowerCase());
        words[0] = words[0][0].toUpperCase() + words[0].slice(1);

        return words.join(" ");
    }
  }

  private mapPropertyToType(property: PowerBIProperty) {
    let type: any = { text: true };
    switch (property.type) {
      case "number": {
        type = { numeric: true };
        switch (property.target.attribute) {
          case "fontSize": {
            type = { formatting: { fontSize: true } };
            break;
          }
          case "fontFamily": {
            type = { formatting: { fontFamily: true } };
            break;
          }
        }
        break;
      }
      case "font-family": {
        type = { formatting: { fontFamily: true } };
        break;
      }
      case "color": {
        type = { fill: { solid: { color: true } } };
        break;
      }
      case "boolean": {
        type = { bool: true };
        break;
      }
      case "enum": {
        switch (property.target.attribute) {
          case "symbol": {
            type = {
              enumeration: symbolTypes.map(sym => {
                return {
                  displayName: sym,
                  value: sym
                };
              })
            };
            break;
          }
        }
        switch (property.target.property) {
          case "alignY":
          case "alignX": {
            type = {
              enumeration: ["end", "middle", "start"].map(sym => {
                return {
                  displayName: sym[0].toLocaleUpperCase() + sym.slice(1),
                  value: sym
                };
              })
            };
            break;
          }
          default: {
            type = { text: true };
          }
        }
        break;
      }
      default: {
        type = { text: true };
      }
    }
    return type;
  }

  private propertyFilter(property: PowerBIProperty) {
    if (
      property.target.attribute === "visible" ||
      property.target.property === "visible"
    ) {
      return false;
    }
    return true;
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
      guid:
        "CHARTICULATOR_VISUAL_" +
        properties.visualName.replace(/[^A-Za-z0-9]+/g, "_") +
        randomHEX32(),
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
    const powerBIObjects: Objects = {};
    for (const property of template.properties.filter(
      this.propertyFilter
    ) as PowerBIProperty[]) {
      const powerBIObjectName = property.objectID;

      const object = Charticulator.Core.Prototypes.findObjectById(
        template.specification,
        powerBIObjectName
      );

      // const object = common.findObjectById(template.specification, powerBIObjectName) as Specification.Object;
      if (!object || !(object as any).exposed) {
        continue;
      }
      if (!powerBIObjects[powerBIObjectName]) {
        powerBIObjects[powerBIObjectName] = {
          displayName: object.properties.name,
          properties: {}
        };
      }
      property.powerBIName = (
        property.objectID +
        "_" +
        property.displayName
      ).replace(/[^0-9a-zA-Z\_]/g, "_");
      const type = this.mapPropertyToType(property);
      powerBIObjects[powerBIObjectName].properties[property.powerBIName] = {
        displayName: this.mapPropertyName(
          property.displayName.split("/").pop()
        ),
        type
      };
    }

    const capabilities: SchemaCapabilities = {
      dataRoles: [
        ...columns.map(column => {
          return {
            displayName: column.displayName,
            name: column.powerBIName,
            kind: "GroupingOrMeasure"
          } as DataRole;
        }),
        properties.enableTooltip && {
          displayName: "Tooltips",
          name: "powerBITooltips",
          kind: "GroupingOrMeasure"
        }
      ],
      dataViewMappings: [
        {
          conditions: [dataViewMappingsConditions],
          categorical: {
            categories: {
              select: [
                ...columns.map(column => {
                  return {
                    bind: { to: column.powerBIName }
                  };
                }),
                {
                  bind: { to: "powerBITooltips" }
                }
              ],
              dataReductionAlgorithm: {
                top: {
                  count: 30000 // That's the maximum
                }
              }
            },
            values: {
              select: [
                ...columns.map(column => {
                  return {
                    bind: { to: column.powerBIName }
                  };
                }),
                {
                  bind: { to: "powerBITooltips" }
                }
              ]
            }
          }
        }
      ],

      /* Tell PBI to allow for sorting */
      sorting: {
        default: {}
      },
      objects: powerBIObjects,
      // Declare that the visual supports highlight.
      // Power BI will give us a set of highlight values instead of filtering the data.
      supportsHighlight: true,
      tooltips: {
        supportedTypes: {
          default: true,
          canvas: true
        }
      }
    };

    const linksTable = this.hasAnchoredLinksAndTable(template);
    if (linksTable) {
      // only source_id and target_id columns are used for creating the links.
      const links = linksTable.columns as PowerBIColumn[];

      for (const column of links) {
        // Refine column names
        column.powerBIName = column.name.replace(/[^0-9a-zA-Z\_]/g, "_");
        dataViewMappingsConditions[column.powerBIName] = { max: 1 };
      }

      links.forEach(link => {
        const linksRole = {
          displayName: link.name,
          name: link.powerBIName,
          kind: "GroupingOrMeasure"
        } as DataRole;

        capabilities.dataRoles.push(linksRole);

        capabilities.dataViewMappings[0].categorical.categories.select.push({
          for: {
            in: link.powerBIName
          }
        });
      });
    }

    if (properties.enableTooltip) {
      template.tables.push({
        columns: [{
          displayName: "Tooltips",
          name: "powerBITooltips",
          type: "string" as any,
          metadata: {
            kind: "categorical" as any
          }
        }],
        name: "powerBITooltips"
      })  
    }

    const apiVersion = "2.1.0";

    const jsConfig: { [name: string]: any } = {
      visualGuid: config.guid,
      pluginName: config.guid,
      visualDisplayName: config.displayName,
      visualVersion: config.version,
      apiVersion,
      templateData: template,
      enableTooltip: properties.enableTooltip
    };
    const visual = resources.visual.replace(
      /[\'\"]\<\%\= *([0-9a-zA-Z\_]+) *\%\>[\'\"]/g,
      (x, a) => JSON.stringify(jsConfig[a])
    );

    const containerScript = await getText(this.containerScriptURL);

    const pbiviz_json = {
      visual: visual_json,
      apiVersion,
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
        iconBase64: properties.visualIcon || resources.icon
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
        "Power BI Custom Visual",
        (template: Specification.Template.ChartTemplate) =>
          new PowerBIVisualGenerator(template, this.containerScriptURL)
      );
  }

  public deactivate() {
    this.context
      .getApplication()
      .unregisterExportTemplateTarget("Power BI Custom Visual");
  }
}

module.exports = {
  PowerBIVisualBuilder,
  PowerBIVisualGenerator
};
