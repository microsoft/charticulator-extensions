// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// The main file for the visual

namespace powerbi.extensibility.visual {
  let isInitialized = false;
  const initializeCallbacks = [];

  function runAfterInitialized(f) {
    if (isInitialized) {
      f();
    } else {
      initializeCallbacks.push(f);
    }
  }

  CharticulatorContainer.initialize({
    MapService: null
  }).then(() => {
    isInitialized = true;
    initializeCallbacks.forEach(f => f());
  });

  class CharticulatorPowerBIVisual {
    public selectionManager: ISelectionManager;
    public template: CharticulatorContainer.Specification.Template.ChartTemplate;
    public containerElement: HTMLDivElement;
    public host: IVisualHost;
    public properties: { [name: string]: any };
    public chartTemplate: CharticulatorContainer.ChartTemplate;
    public chartInstance: CharticulatorContainer.ChartTemplateInstance;

    constructor(options: VisualConstructorOptions) {
      try {
        this.selectionManager = options.host.createSelectionManager();
        this.template = "<%= templateData %>" as any;
        this.containerElement = document.createElement("div");
        this.host = options.host;
        this.properties = {};
        options.element.appendChild(this.containerElement);
        this.chartTemplate = new CharticulatorContainer.ChartTemplate(
          this.template
        );
        this.chartInstance = null;
        for (const item of this.template.properties) {
          this.properties[item.objectID + "/" + item.displayName] =
            item.default;
        }
      } catch (e) {
        console.log(e);
      }
    }

    public resize(width: number, height: number) {
      this.containerElement.style.width = width + "px";
      this.containerElement.style.height = height + "px";
    }

    /** Get a Charticulator dataset from the options */
    public getDataset(options: VisualUpdateOptions) {}

    public getProperties(options) {}

    public update(options) {
      runAfterInitialized(() => this.updateRun(options));
    }

    public updateRun(options) {}

    public enumerateObjectInstances(options) {
      const objectName = options.objectName;
      const objectEnumeration = [];
      switch (objectName) {
        case "chartOptions":
          objectEnumeration.push({
            objectName,
            properties: this.properties,
            selector: null
          });
      }
      return objectEnumeration;
    }
  }

  (powerbi as any).extensibility.visual["<%= visualGuid %>"] = {
    CharticulatorPowerBIVisual
  };
}
namespace powerbi.visuals.plugins {
  (powerbi as any).visuals.plugins["<%= pluginName %>"] = {
    name: "<%= pluginName %>",
    displayName: "<%= visualDisplayName %>",
    class: "CharticulatorPowerBIVisual",
    version: "<%= visualVersion %>",
    apiVersion: "<%= apiVersion %>",
    create: options => {
      return new powerbi.extensibility.visual[
        "<%= visualGuid %>"
      ].CharticulatorPowerBIVisual(options);
    },
    custom: true
  };
}
