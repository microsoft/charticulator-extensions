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

  interface PowerBIColumn
    extends CharticulatorContainer.Specification.Template.Column {
    powerBIName: string;
  }

  interface PowerBIProperty
    extends CharticulatorContainer.Specification.Template.Property {
    powerBIName: string;
  }

  class CharticulatorPowerBIVisual {
    public selectionManager: ISelectionManager;
    public template: CharticulatorContainer.Specification.Template.ChartTemplate;
    public containerElement: HTMLDivElement;
    public host: IVisualHost;
    public properties: { [name: string]: any };
    public chartTemplate: CharticulatorContainer.ChartTemplate;
    public chartContainer: CharticulatorContainer.ChartContainer;
    protected currentDatasetJSON: string;

    constructor(options: VisualConstructorOptions) {
      try {
        this.selectionManager = options.host.createSelectionManager();
        this.template = "<%= templateData %>" as any;
        this.containerElement = document.createElement("div");
        this.containerElement.style.cursor = "default";
        this.containerElement.style.pointerEvents = "none";
        this.host = options.host;
        options.element.appendChild(this.containerElement);
        this.chartTemplate = new CharticulatorContainer.ChartTemplate(
          this.template
        );
        this.properties = this.getProperties(null);
        // Seems this won't be called when using highlight mode
        // this.selectionManager.registerOnSelectCallback(
        //   (ids: ISelectionId[]) => {
        //     const rowIndices: number[] = [];
        //     for (const id of ids) {
        //       if (this.selectionID2RowIndex.has(id)) {
        //         const row = this.selectionID2RowIndex.get(id);
        //         rowIndices.push(row);
        //       }
        //     }
        //     console.log("OnSelectCallback", rowIndices, ids.length);
        //     // if (rowIndices.length > 0) {
        //     //   this.chartContainer.setSelection("default", rowIndices);
        //     // } else {
        //     //   this.chartContainer.clearSelection();
        //     // }
        //   }
        // );
      } catch (e) {
        console.log(e);
      }
    }

    public resize(width: number, height: number) {
      this.containerElement.style.width = width + "px";
      this.containerElement.style.height = height + "px";
    }

    /** Get a Charticulator dataset from the options */
    public getDataset(
      options: VisualUpdateOptions
    ): {
      dataset: CharticulatorContainer.Dataset.Dataset;
      rowInfo: Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number }
      >;
    } {
      if (
        !options.dataViews ||
        !options.dataViews[0] ||
        !options.dataViews[0].categorical ||
        !options.dataViews[0].categorical.categories ||
        !options.dataViews[0].categorical.categories[0]
      ) {
        return null;
      }
      const dv = options.dataViews[0];
      const categorical = dv.categorical;
      const category = categorical.categories[0];
      const valueColumns = categorical.values;
      if (!valueColumns) {
        return null;
      }

      const columnToValues: { [name: string]: DataViewValueColumn } = {};

      const columns = this.template.tables[0].columns as PowerBIColumn[];

      for (const column of columns) {
        let found = false;
        for (const v of valueColumns) {
          if (v.source.roles[column.powerBIName]) {
            columnToValues[column.powerBIName] = v;
            found = true;
          }
        }
        if (!found) {
          return null;
        }
      }
      const rowInfo = new Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number }
      >();
      const dataset: CharticulatorContainer.Dataset.Dataset = {
        name: "Dataset",
        tables: [
          {
            name: "default",
            columns: columns.map(column => {
              return {
                name: column.powerBIName,
                type: column.type,
                metadata: column.metadata
              };
            }),
            rows: category.values
              .map((_, i) => {
                const obj: CharticulatorContainer.Dataset.Row = {
                  _id: "ID" + i.toString()
                };
                let allHighlighted = true;
                for (const column of columns) {
                  const valueColumn = columnToValues[column.powerBIName];
                  const value = valueColumn.values[i];
                  if (value == null) {
                    return null;
                  }
                  obj[column.powerBIName] = value.valueOf();
                  const highlight =
                    valueColumn.highlights && // Highlights exists
                    valueColumn.highlights[i] != null && // Not null
                    valueColumn.highlights[i].valueOf() == value.valueOf(); // Same value
                  if (!highlight) {
                    allHighlighted = false;
                  }
                  if (
                    column.metadata &&
                    column.metadata.kind == "categorical"
                  ) {
                    obj[column.powerBIName] = obj[
                      column.powerBIName
                    ].toString();
                  }
                }
                rowInfo.set(obj, {
                  highlight: allHighlighted,
                  index: i
                });
                return obj;
              })
              .filter(x => x != null)
          }
        ]
      };
      return { dataset, rowInfo };
    }

    public getProperties(options: VisualUpdateOptions) {
      const defaultProperties: { [name: string]: any } = {};
      for (const p of this.template.properties as PowerBIProperty[]) {
        defaultProperties[p.powerBIName] = p.default;
      }

      if (
        !options ||
        !options.dataViews ||
        !options.dataViews[0] ||
        !options.dataViews[0].metadata
      ) {
        return defaultProperties;
      }
      const objects = options.dataViews[0].metadata.objects;
      if (!objects) {
        return defaultProperties;
      }

      for (const p of this.template.properties as PowerBIProperty[]) {
        const object = objects.chartOptions;
        if (object[p.powerBIName] != undefined) {
          defaultProperties[p.powerBIName] = object[p.powerBIName];
        }
      }

      return defaultProperties;
    }

    public unmountContainer() {
      if (this.chartContainer != null) {
        this.chartContainer.unmount();
        this.chartContainer = null;
      }
    }

    public update(options: VisualUpdateOptions) {
      this.properties = this.getProperties(options);
      runAfterInitialized(() => this.updateRun(options));
    }

    public updateRun(options: VisualUpdateOptions) {
      try {
        this.resize(options.viewport.width, options.viewport.height);

        const getDatasetResult = this.getDataset(options);
        if (getDatasetResult == null) {
          // If dataset is null, show a warning message
          this.containerElement.innerHTML = `
                <h2>Dataset incomplete. Please specify all data fields.</h2>
            `;
          this.currentDatasetJSON = null;
          this.unmountContainer();
        } else {
          const { dataset } = getDatasetResult;
          // Check if dataset is the same
          const datasetJSON = JSON.stringify(dataset);
          if (datasetJSON != this.currentDatasetJSON) {
            this.unmountContainer();
          }
          this.currentDatasetJSON = datasetJSON;

          // Recreate chartContainer if not exist
          if (!this.chartContainer) {
            this.containerElement.innerHTML = "";
            this.chartTemplate.reset();

            const columns = this.template.tables[0].columns as PowerBIColumn[];
            this.chartTemplate.assignTable(
              this.template.tables[0].name,
              "default"
            );
            for (const column of columns) {
              this.chartTemplate.assignColumn(
                this.template.tables[0].name,
                column.name,
                column.powerBIName
              );
            }
            const chart = this.chartTemplate.instantiate(dataset);
            // Apply chart properties
            for (const property of this.template
              .properties as PowerBIProperty[]) {
              if (this.properties[property.powerBIName] === undefined) {
                continue;
              }
              if (property.target.property) {
                CharticulatorContainer.ChartTemplate.SetChartProperty(
                  chart,
                  property.objectID,
                  property.target.property,
                  this.properties[property.powerBIName]
                );
              } else {
                CharticulatorContainer.ChartTemplate.SetChartAttributeMapping(
                  chart,
                  property.objectID,
                  property.target.attribute,
                  {
                    type: "value",
                    value: this.properties[property.powerBIName]
                  } as CharticulatorContainer.Specification.ValueMapping
                );
              }
            }

            // Make selection ids:
            const selectionIDs = [];
            dataset.tables[0].rows.forEach((row, i) => {
              const selectionID = this.host
                .createSelectionIdBuilder()
                .withCategory(
                  options.dataViews[0].categorical.categories[0],
                  getDatasetResult.rowInfo.get(row).index
                )
                .createSelectionId();
              selectionIDs.push(selectionID);
            });
            this.chartContainer = new CharticulatorContainer.ChartContainer(
              chart,
              dataset
            );
            this.chartContainer.addSelectionListener((table, rowIndices) => {
              if (
                table != null &&
                rowIndices != null &&
                rowIndices.length > 0
              ) {
                const ids = rowIndices
                  .map(i => selectionIDs[i])
                  .filter(x => x != null);
                this.selectionManager.select(ids);
                this.chartContainer.setSelection(table, rowIndices);
              } else {
                this.selectionManager.clear();
              }
            });
            this.chartContainer.mount(this.containerElement);
          }

          if (this.chartContainer) {
            // Feed in properties
            for (const property of this.template
              .properties as PowerBIProperty[]) {
              if (this.properties[property.powerBIName] === undefined) {
                continue;
              }
              if (property.target.property) {
                this.chartContainer.setProperty(
                  property.objectID,
                  property.target.property,
                  this.properties[property.powerBIName]
                );
              }
              if (property.target.attribute) {
                this.chartContainer.setAttributeMapping(
                  property.objectID,
                  property.target.attribute,
                  {
                    type: "value",
                    value: this.properties[property.powerBIName]
                  } as CharticulatorContainer.Specification.ValueMapping
                );
              }
            }
            if (getDatasetResult && getDatasetResult.rowInfo) {
              const rows = dataset.tables[0].rows;
              const indices = [];
              for (let i = 0; i < rows.length; i++) {
                if (getDatasetResult.rowInfo.get(rows[i]).highlight) {
                  indices.push(i);
                }
              }
              if (indices.length > 0) {
                this.chartContainer.setSelection("default", indices);
              } else {
                this.chartContainer.clearSelection();
              }
            }
            this.chartContainer.resize(
              options.viewport.width,
              options.viewport.height
            );
          }
        }
      } catch (e) {
        console.log(e);
      }
    }

    public enumerateObjectInstances(options) {
      const objectName = options.objectName;
      const objectEnumeration = [];
      const properties: { [name: string]: any } = {};
      for (const p of this.template.properties as PowerBIProperty[]) {
        if (this.properties[p.powerBIName] !== undefined) {
          properties[p.powerBIName] = this.properties[p.powerBIName];
        } else {
          if (this.chartContainer) {
            if (p.target.property) {
              properties[p.powerBIName] = this.chartContainer.getProperty(
                p.objectID,
                p.target.property
              );
            }
            if (p.target.attribute) {
              const mapping = this.chartContainer.getAttributeMapping(
                p.objectID,
                p.target.attribute
              );
              if (mapping.type == "value") {
                properties[p.powerBIName] = (mapping as any).value;
              }
            }
          } else {
            if (p.target.property) {
              properties[
                p.powerBIName
              ] = CharticulatorContainer.ChartTemplate.GetChartProperty(
                this.template.specification,
                p.objectID,
                p.target.property
              );
            }
            if (p.target.attribute) {
              const mapping = CharticulatorContainer.ChartTemplate.GetChartAttributeMapping(
                this.template.specification,
                p.objectID,
                p.target.attribute
              );
              if (mapping.type == "value") {
                properties[p.powerBIName] = (mapping as any).value;
              }
            }
          }
        }
      }
      switch (objectName) {
        case "chartOptions":
          objectEnumeration.push({
            objectName,
            properties,
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