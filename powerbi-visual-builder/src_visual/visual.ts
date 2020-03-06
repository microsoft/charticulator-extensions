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

  function arrayEquals(a1: number[], a2: number[]) {
    const s1 = new Set(a1);
    const s2 = new Set(a2);
    return a1.every(i => s2.has(i)) && a2.every(i => s1.has(i));
  }

  class CharticulatorPowerBIVisual {
    protected host: IVisualHost;
    protected selectionManager: ISelectionManager;

    protected template: CharticulatorContainer.Specification.Template.ChartTemplate;
    protected enableTooltip: boolean;

    protected container: HTMLElement;
    protected divChart: HTMLDivElement;
    protected currentDatasetJSON: string;

    protected properties: { [name: string]: any };
    protected chartTemplate: CharticulatorContainer.ChartTemplate;
    protected chartContainer: CharticulatorContainer.ChartContainer;

    protected currentX: number;
    protected currentY: number;
    protected handleMouseMove: () => void;

    constructor(options: VisualConstructorOptions) {
      try {
        this.selectionManager = options.host.createSelectionManager();
        this.template = "<%= templateData %>" as any;
        this.enableTooltip = "<%= enableTooltip %>" as any;
        this.container = options.element;
        this.divChart = document.createElement("div");
        this.divChart.style.cursor = "default";
        this.divChart.style.pointerEvents = "none";
        this.host = options.host;
        options.element.appendChild(this.divChart);
        this.chartTemplate = new CharticulatorContainer.ChartTemplate(
          this.template
        );
        this.properties = this.getProperties(null);

        if (this.enableTooltip) {
          // Handles mouse move events for displaying tooltips.
          window.addEventListener("mousemove", e => {
            const bbox = this.container.getBoundingClientRect();
            this.currentX = e.pageX - bbox.left;
            this.currentY = e.pageY - bbox.top;
            if (this.handleMouseMove) {
              this.handleMouseMove();
            }
          });
        }
      } catch (e) {
        console.log(e);
      }
    }

    public resize(width: number, height: number) {
      this.divChart.style.width = width + "px";
      this.divChart.style.height = height + "px";
    }

    /** Get a Charticulator dataset from the options */
    protected getDataset(
      options: VisualUpdateOptions
    ): {
      dataset: CharticulatorContainer.Dataset.Dataset;
      rowInfo: Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number; granularity: string }
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
      const categories = options.dataViews[0].categorical.categories;
      const valueColumns = categorical.values;

      // Match columns
      const columnToValues: {
        [name: string]: {
          values: CharticulatorContainer.Specification.DataValue[];
          highlights: boolean[];
        };
      } = {};
      const columns = this.template.tables[0].columns as PowerBIColumn[];
      for (const column of columns) {
        let found = false;
        if (valueColumns != null) {
          for (const v of valueColumns) {
            if (v.source.roles[column.powerBIName]) {
              columnToValues[column.powerBIName] = {
                values: CharticulatorContainer.Dataset.convertColumnType(
                  v.values.map(x => (x == null ? null : x.valueOf())),
                  column.type
                ),
                highlights: v.values.map((value, i) => {
                  return v.highlights
                    ? v.highlights[i] != null && value != null
                      ? v.highlights[i].valueOf() == value.valueOf()
                      : false
                    : false;
                })
              };
              found = true;
            }
          }
        }
        for (const v of categories) {
          if (
            v.source.roles[column.powerBIName] &&
            !columnToValues[column.powerBIName]
          ) {
            columnToValues[column.powerBIName] = {
              values: CharticulatorContainer.Dataset.convertColumnType(
                v.values.map(x => (x == null ? null : x.valueOf())),
                column.type
              ),
              highlights: v.values.map((value, i) => {
                return false;
              })
            };
            found = true;
          }
        }
        if (!found) {
          return null;
        }
      }

      const linkColumns = options.dataViews[0].categorical.categories;
      const links =
        this.template.tables[1] &&
        (this.template.tables[1].columns as PowerBIColumn[]);

      if (links && linkColumns) {
        for (const column of links) {
          for (const v of linkColumns) {
            if (v.source.roles[column.powerBIName]) {
              columnToValues[column.powerBIName || column.name] = {
                values: CharticulatorContainer.Dataset.convertColumnType(
                  v.values.map(x => (x == null ? null : x.valueOf())),
                  column.type
                ),
                highlights: v.values.map((value, i) => {
                  return null;
                })
              };
            }
          }
        }
      }

      const rowInfo = new Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number; granularity: string }
      >();

      const uniqueRows = new Set<string>();

      const rows = categories[0].values
        .map((categoryValue, i) => {
          const obj: CharticulatorContainer.Dataset.Row = {
            _id: /*"ID" +*/ i.toString()
          };
          let allHighlighted = true;
          let rowHash = "";
          for (const column of columns) {
            const valueColumn = columnToValues[column.powerBIName];
            const value = valueColumn.values[i];

            if (value == null) {
              return null;
            }
            obj[column.powerBIName] = value;
            rowHash += value.toString();
            if (!valueColumn.highlights[i]) {
              allHighlighted = false;
            }
          }

          const catDate = categoryValue as Date;
          let granularity = categoryValue.valueOf().toString();

          // Try to do some extra formatting for dates
          if (catDate && typeof catDate.toDateString === "function") {
            granularity = catDate.toDateString();
          }

          if (!uniqueRows.has(rowHash)) {
            uniqueRows.add(rowHash);
            rowInfo.set(obj, {
              highlight: allHighlighted,
              index: i,
              granularity
            });
            return obj;
          } else {
            return null;
          }
        })
        .filter(x => x != null);

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
            rows
          },
          links && linkColumns && {
            name: "links",
            columns:
              linkColumns.length >= 2
                ? links.map(column => {
                    return {
                      name: column.powerBIName,
                      type: column.type,
                      metadata: column.metadata
                    };
                  })
                : null,
            rows: links && linkColumns && categories[0].values
              .map((source, index) => {
                const obj: CharticulatorContainer.Dataset.Row = {
                  _id: index.toString()
                };
                for (const column of links) {
                  const valueColumn =
                    columnToValues[column.powerBIName || column.name];
                  if (valueColumn) {
                    const value = valueColumn.values[index];
                    obj[column.powerBIName || column.name] = value;
                  }
                }
                if (obj.source_id && obj.target_id) {
                  return obj;
                } else {
                  return null;
                }
              })
              .filter(row => row)
          }
        ].filter(table => table && table.columns)
      };
      return { dataset, rowInfo };
    }

    protected getProperties(options: VisualUpdateOptions) {
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

    protected unmountContainer() {
      if (this.chartContainer != null) {
        this.chartContainer.unmount();
        this.chartContainer = null;
      }
    }

    public update(options: VisualUpdateOptions) {
      if (options.type & (VisualUpdateType.Data | VisualUpdateType.Style)) {
        // If data or properties changed, re-generate the visual
        this.properties = this.getProperties(options);
        runAfterInitialized(() => this.updateRun(options));
      } else if (
        options.type &
        (VisualUpdateType.Resize | VisualUpdateType.ResizeEnd)
      ) {
        // If only size changed, just run resize
        if (this.chartContainer) {
          this.resize(options.viewport.width, options.viewport.height);
          this.chartContainer.resize(
            options.viewport.width,
            options.viewport.height
          );
        }
      }
    }

    protected updateRun(options: VisualUpdateOptions) {
      try {
        this.resize(options.viewport.width, options.viewport.height);

        const getDatasetResult = this.getDataset(options);
        if (getDatasetResult == null) {
          // If dataset is null, show a warning message
          this.divChart.innerHTML = `
                <h2>Dataset incomplete. Please specify all data fields.</h2>
            `;
          this.currentDatasetJSON = null;
          this.handleMouseMove = null;
          this.unmountContainer();
        } else {
          const { dataset } = getDatasetResult;
          // Check if dataset is the same
          const datasetJSON = JSON.stringify(dataset);
          if (datasetJSON != this.currentDatasetJSON) {
            this.handleMouseMove = null;
            this.unmountContainer();
          }
          this.currentDatasetJSON = datasetJSON;

          // Recreate chartContainer if not exist
          if (!this.chartContainer) {
            this.divChart.innerHTML = "";
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

            // links table
            const links = this.template.tables[1] && this.template.tables[1].columns as PowerBIColumn[];
            if (links) {
              this.chartTemplate.assignTable(
                this.template.tables[1].name,
                "links"
              );
              for (const column of links) {
                this.chartTemplate.assignColumn(
                  this.template.tables[1].name,
                  column.name,
                  column.powerBIName
                );
              }
            }
            
            const instance = this.chartTemplate.instantiate(dataset);
            const { chart } = instance;

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
            const selectionIDs: visuals.ISelectionId[] = [];
            const selectionID2RowIndex = new WeakMap<ISelectionId, number>();
            dataset.tables[0].rows.forEach((row, i) => {
              const selectionID = this.host
                .createSelectionIdBuilder()
                .withCategory(
                  options.dataViews[0].categorical.categories[0],
                  getDatasetResult.rowInfo.get(row).index
                )
                .createSelectionId();
              selectionIDs.push(selectionID);
              selectionID2RowIndex.set(selectionID, i);
            });
            this.chartContainer = new CharticulatorContainer.ChartContainer(
              instance,
              dataset
            );
            this.chartContainer.addSelectionListener((table, rowIndices) => {
              if (
                table != null &&
                rowIndices != null &&
                rowIndices.length > 0
              ) {
                // Power BI's toggle behavior
                let alreadySelected = false;
                if (this.selectionManager.hasSelection()) {
                  const ids = this.selectionManager
                    .getSelectionIds()
                    .map(id => selectionID2RowIndex.get(id));
                  if (arrayEquals(ids, rowIndices)) {
                    alreadySelected = true;
                  }
                }
                if (alreadySelected) {
                  this.selectionManager.clear();
                  this.chartContainer.clearSelection();
                } else {
                  const ids = rowIndices
                    .map(i => selectionIDs[i])
                    .filter(x => x != null);
                  this.selectionManager.select(ids);
                  this.chartContainer.setSelection(table, rowIndices);
                }
              } else {
                this.selectionManager.clear();
              }
            });
            if (this.enableTooltip && this.host.tooltipService.enabled()) {
              const service = this.host.tooltipService;

              this.chartContainer.addMouseEnterListener((table, rowIndices) => {
                const ids = rowIndices
                  .map(i => selectionIDs[i])
                  .filter(x => x != null);
                const info = {
                  coordinates: [this.currentX, this.currentY],
                  isTouchEvent: false,
                  dataItems: Array.prototype.concat(
                    ...rowIndices.map(idx => {
                      const row = dataset.tables[0].rows[idx];
                      return Object.keys(row)
                        .filter(x => x != "_id")
                        .map(key => {
                          let value = row[key];
                          const column = dataset.tables[0].columns.filter(
                            n => n.name === key
                          )[0];

                          // Attempt to format numbers/dates nicely
                          /** Data type in memory (number, string, Date, boolean, etc) */
                          if (value !== undefined && value !== null) {
                            if (
                              column.type ===
                              CharticulatorContainer.Specification.DataType
                                .Number
                            ) {
                              value = parseFloat(value + "").toFixed(2);
                            } else if (
                              column.type ===
                              CharticulatorContainer.Specification.DataType.Date
                            ) {
                              const numVal = value as number;
                              if (typeof numVal.toFixed === "function") {
                                const parsed = new Date(numVal);
                                value = parsed.toDateString();
                              }
                            }
                          }

                          const header = getDatasetResult.rowInfo.get(row)
                            .granularity;
                          return {
                            displayName: key,
                            header,
                            value
                          };
                        });
                    })
                  ),
                  identities: ids
                };
                service.show(info);
                this.handleMouseMove = () => {
                  service.move({
                    ...info,
                    coordinates: [this.currentX, this.currentY]
                  });
                };
              });
              this.chartContainer.addMouseLeaveListener((table, rowIndices) => {
                service.hide({ isTouchEvent: false, immediately: false });
                this.handleMouseMove = null;
              });
            }
            this.chartContainer.mount(this.divChart);
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
