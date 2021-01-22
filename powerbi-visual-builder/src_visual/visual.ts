// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// The main file for the visual

namespace powerbi.extensibility.visual {
  import AttributeMap = CharticulatorContainer.Specification.AttributeMap;
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

  const rawColumnPostFix = CharticulatorContainer.Dataset.rawColumnPostFix;
  const applyDateFormat = CharticulatorContainer.Utils.applyDateFormat;
  const powerBITooltipsTablename = "powerBITooltips";
  const propertyAndObjectNamePrefix = "ID_";
  const rawColumnFilter = (
    allColumns: CharticulatorContainer.Specification.Template.Column[]
  ) => column =>
    !column.metadata.isRaw &&
    (column.type === "date" || column.type === "boolean") &&
    !allColumns.filter(c => c.name === column.metadata.rawColumnName).length;

  const rawColumnMapper = column => {
    const rawName = `${refineColumnName(column.name)}${rawColumnPostFix}`;
    return {
      ...column,
      name: rawName,
      displayName: rawName,
      powerBIName: rawName,
      metadata: {
        ...column.metadata,
        rawColumnName: null,
        isRaw: true
      }
    };
  };

  const parseSafe = (val: string, defaultVal: any) => {
    try {
      return JSON.parse(val);
    } catch (ex) {
      return defaultVal;
    }
  };

  const refineColumnName = (name: string) =>
    name.replace(/[^0-9a-zA-Z\_]/g, "_");

  class CharticulatorPowerBIVisual {
    protected host: IVisualHost;
    protected selectionManager: ISelectionManager;

    protected template: CharticulatorContainer.Specification.Template.ChartTemplate;
    protected enableDrillDown: boolean;
    protected drillDownColumns: string[];

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
        this.enableDrillDown = "<%= enableDrillDown %>" as any;
        this.drillDownColumns = "<%= drillDownColumns %>" as any;
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

        // Handles mouse move events for displaying tooltips.
        window.addEventListener("mousemove", e => {
          const bbox = this.container.getBoundingClientRect();
          this.currentX = e.pageX - bbox.left;
          this.currentY = e.pageY - bbox.top;
          if (this.handleMouseMove) {
            this.handleMouseMove();
          }
        });
      } catch (e) {
        console.log(e);
      }
    }

    public resize(width: number, height: number) {
      this.divChart.style.width = width + "px";
      this.divChart.style.height = height + "px";
    }

    private mapColumns(
      powerBIColumn: DataViewValueColumn | DataViewCategoryColumn,
      type: CharticulatorContainer.Specification.DataType,
      rawFormat?: string
    ): Array<{
      values: CharticulatorContainer.Specification.DataValue[];
      highlights: boolean[];
    }> {
      const columns = [
        {
          values: CharticulatorContainer.Dataset.convertColumnType(
            powerBIColumn.values.map(x => (x == null ? null : x.valueOf())),
            type
          ),
          highlights: powerBIColumn.values.map((value, i) => {
            return (powerBIColumn as DataViewValueColumn).highlights &&
              (powerBIColumn as DataViewValueColumn).highlights[i] != null &&
              value != null
              ? (powerBIColumn as DataViewValueColumn).highlights[
                  i
                ].valueOf() <= value.valueOf()
              : false;
          })
        }
      ];
      if (type === "date" || type === "boolean") {
        if (rawFormat) {
          columns.push({
            values: CharticulatorContainer.Dataset.convertColumnType(
              powerBIColumn.values.map(x => (x == null ? null : x.valueOf())),
              type
            ).map(x =>
              rawFormat ? applyDateFormat(x as Date, rawFormat) : x.toString()
            ),
            highlights: powerBIColumn.values.map((value, i) => {
              return (powerBIColumn as DataViewValueColumn).highlights &&
                (powerBIColumn as DataViewValueColumn).highlights[i] != null &&
                value != null
                ? (powerBIColumn as DataViewValueColumn).highlights[
                    i
                  ].valueOf() <= value.valueOf()
                : false;
            })
          });
        } else {
          columns.push({
            values: powerBIColumn.values.map(x =>
              x == null
                ? null
                : rawFormat
                ? applyDateFormat(x as Date, rawFormat)
                : x.toString()
            ),
            highlights: powerBIColumn.values.map((value, i) => {
              return (powerBIColumn as DataViewValueColumn).highlights &&
                (powerBIColumn as DataViewValueColumn).highlights[i] != null &&
                value != null
                ? (powerBIColumn as DataViewValueColumn).highlights[
                    i
                  ].valueOf() <= value.valueOf()
                : false;
            })
          });
        }
      }

      return columns;
    }

    private deepClone<T>(obj: T): T {
      return JSON.parse(JSON.stringify(obj));
    }

    private getDefaultTable(
      template: CharticulatorContainer.Specification.Template.ChartTemplate
    ): CharticulatorContainer.Specification.Template.Table {
      const plotSegment = template.specification.elements.find(element =>
        (CharticulatorContainer as any).Prototypes.isType(
          element.classID,
          "plot-segment"
        )
      ) as CharticulatorContainer.Specification.PlotSegment;
      const tableName = plotSegment.table;
      return template.tables.find(table => table.name === tableName);
    }

    private getLinksTable(
      template: CharticulatorContainer.Specification.Template.ChartTemplate
    ): CharticulatorContainer.Specification.Template.Table {
      const link: CharticulatorContainer.Specification.Links = template.specification.elements.find(
        element =>
          Boolean(
            element.classID === "links.table" &&
              element.properties.anchor1 &&
              element.properties.anchor2
          )
      );

      if (link) {
        const properties = link.properties as AttributeMap;
        const linkTableName =
          properties.linkTable && (properties.linkTable as AttributeMap).table;
        const linkTable = template.tables.find(
          table => table.name === linkTableName
        );

        return linkTable;
      }
    }

    private getTooltipsTable(
      template: CharticulatorContainer.Specification.Template.ChartTemplate
    ): CharticulatorContainer.Specification.Template.Table {
      return template.tables.find(
        table => table.name === powerBITooltipsTablename
      );
    }

    public getUserColumnName(dv: DataView, columnName: string) {
      const column = dv.metadata.columns.find(
        column => column.roles[columnName]
      );
      return column && column.displayName;
    }

    /** Get a Charticulator dataset from the options */
    protected getDataset(
      dataView: DataView
    ): {
      dataset: CharticulatorContainer.Dataset.Dataset;
      rowInfo: Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number; granularity: string }
      >;
    } {
      if (
        !dataView ||
        !dataView.categorical ||
        !dataView.categorical.categories ||
        !dataView.categorical.categories[0]
      ) {
        return null;
      }
      const categorical = dataView.categorical;
      const categories = categorical.categories;
      const valueColumns = categorical.values;

      // Match columns
      const columnToValues: {
        [name: string]: {
          values: CharticulatorContainer.Specification.DataValue[];
          highlights: boolean[];
        };
      } = {};
      const defaultTable = this.getDefaultTable(this.template);
      let columns = defaultTable.columns.filter(
        col => !col.metadata.isRaw
      ) as PowerBIColumn[];
      for (const chartColumn of columns) {
        let found = false;
        if (valueColumns != null) {
          for (const powerBIColumn of valueColumns) {
            if (powerBIColumn.source.roles[chartColumn.powerBIName]) {
              const [converted, raw] = this.mapColumns(
                powerBIColumn,
                chartColumn.type,
                chartColumn.metadata.format
              );
              columnToValues[
                chartColumn.powerBIName || refineColumnName(chartColumn.name)
              ] = converted;
              if (raw && !chartColumn.metadata.isRaw) {
                columnToValues[
                  `${chartColumn.powerBIName ||
                    refineColumnName(chartColumn.name)}${rawColumnPostFix}`
                ] = raw;
              }
              found = true;
            }
          }
        }
        for (const powerBIColumn of categories.reverse()) {
          // reverse - to take the latest column (if drill down was enabled)
          if (
            powerBIColumn.source.roles[
              chartColumn.powerBIName || chartColumn.name
            ] &&
            !columnToValues[chartColumn.powerBIName || chartColumn.name]
          ) {
            const [converted, raw] = this.mapColumns(
              powerBIColumn,
              chartColumn.type,
              chartColumn.metadata && chartColumn.metadata.format
            );
            columnToValues[
              chartColumn.powerBIName || refineColumnName(chartColumn.name)
            ] = converted;
            if (raw && !chartColumn.metadata.isRaw) {
              columnToValues[
                `${chartColumn.powerBIName ||
                  refineColumnName(chartColumn.name)}${rawColumnPostFix}`
              ] = raw;
            }
            found = true;
          }
        }
        if (!found) {
          return null;
        }
      }
      const defaultTableRawColumns = defaultTable.columns
        // add powerBIName to raw columns
        .map((column: PowerBIColumn) => {
          if (!column.powerBIName) {
            column.powerBIName = refineColumnName(column.name);
          }
          return column;
        })
        .filter(col => !col.metadata.isRaw)
        .filter(rawColumnFilter(defaultTable.columns))
        .map(rawColumnMapper);
      columns = defaultTable.columns = defaultTable.columns.concat(
        defaultTableRawColumns
      ) as PowerBIColumn[];

      const linksTable = this.getLinksTable(this.template);
      const powerBILinkColumns = dataView.categorical.categories;
      let chartLinks =
        linksTable &&
        (linksTable.columns.filter(
          col => !col.metadata.isRaw
        ) as PowerBIColumn[]);

      if (chartLinks && powerBILinkColumns) {
        for (const chartColumn of chartLinks) {
          for (const powerBIColumn of powerBILinkColumns) {
            if (powerBIColumn.source.roles[chartColumn.powerBIName]) {
              const [converted, raw] = this.mapColumns(
                powerBIColumn,
                chartColumn.type,
                chartColumn.metadata && chartColumn.metadata.format
              );
              columnToValues[
                chartColumn.powerBIName || chartColumn.name
              ] = converted;
              if (raw && !chartColumn.metadata.isRaw) {
                columnToValues[
                  `${chartColumn.powerBIName ||
                    refineColumnName(chartColumn.name)}${rawColumnPostFix}`
                ] = raw;
              }
            }
          }
        }

        const linksTableRawColumns = linksTable.columns
          .filter(rawColumnFilter(linksTable.columns))
          .map(rawColumnMapper);
        chartLinks = linksTable.columns = linksTable.columns.concat(
          linksTableRawColumns
        ) as PowerBIColumn[];
      }

      const tooltipsTable = this.getTooltipsTable(this.template);
      const tooltipsTableColumns = [
        ...categorical.categories.filter(
          cat => cat.source.roles.powerBITooltips
        ),
        ...(categorical.values
          ? categorical.values.filter(cat => cat.source.roles.powerBITooltips)
          : [])
      ];

      if (tooltipsTable && tooltipsTableColumns) {
        const type =
          (tooltipsTable.columns.length && tooltipsTable.columns[0].type) ||
          CharticulatorContainer.Specification.DataType.String;
        const metadata = (tooltipsTable.columns.length &&
          tooltipsTable.columns[0].metadata) || {
          kind: "categorical"
        };
        tooltipsTable.columns = [];
        tooltipsTableColumns.forEach(powerBIColumn => {
          if (!columnToValues[powerBIColumn.source.displayName]) {
            const [converted, raw] = this.mapColumns(powerBIColumn, type);
            columnToValues[powerBIColumn.source.displayName] = converted;
            if (raw) {
              columnToValues[
                `${refineColumnName(
                  powerBIColumn.source.displayName
                )}${rawColumnPostFix}`
              ] = raw;
            }
          }
          tooltipsTable.columns.push({
            displayName: powerBIColumn.source.displayName,
            name: powerBIColumn.source.displayName,
            metadata,
            type
          });
        });
      }
      const tooltips =
        tooltipsTable && (tooltipsTable.columns as PowerBIColumn[]);

      const rowInfo = new Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number; granularity: string }
      >();

      const uniqueRows = new Set<string>();

      const rowIdentity = categories.filter(
        category => category.source.roles.primarykey
      );
      const rows = categories[0].values
        .map((categoryValue, i) => {
          const obj: CharticulatorContainer.Dataset.Row = {
            _id: /*"ID" +*/ i.toString()
          };
          let rowHasHighlightedColumn = false;
          let rowHash = rowIdentity.length
            ? rowIdentity.map(idRow => idRow.values[i]).toString()
            : "";
          for (const column of columns) {
            const valueColumn = columnToValues[column.powerBIName];
            if (!valueColumn) {
              return null;
            }
            let value = valueColumn.values[i];

            if (value == null) {
              if (
                columns.find(col => col.name === column.name).metadata.kind ===
                "numerical"
              ) {
                value = 0;
              } else {
                return null;
              }
            }
            obj[column.powerBIName] = value;
            rowHash += (value || "null").toString();
            // if one value column has highlights
            if (valueColumn.highlights[i]) {
              rowHasHighlightedColumn = true;
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
              highlight: rowHasHighlightedColumn,
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
            name: defaultTable.name,
            columns: columns.map(column => {
              return {
                displayName: this.getUserColumnName(
                  dataView,
                  column.powerBIName
                ),
                name: column.powerBIName,
                type: column.type,
                metadata: column.metadata
              };
            }),
            rows
          },
          chartLinks &&
            powerBILinkColumns && {
              name: linksTable.name,
              columns:
                powerBILinkColumns.length >= 2
                  ? chartLinks.map(column => {
                      return {
                        displayName: this.getUserColumnName(
                          dataView,
                          column.powerBIName
                        ),
                        name: column.powerBIName,
                        type: column.type,
                        metadata: column.metadata
                      };
                    })
                  : null,
              rows: categories[0].values
                .map((source, index) => {
                  const obj: CharticulatorContainer.Dataset.Row = {
                    _id: index.toString()
                  };
                  for (const column of chartLinks) {
                    const valueColumn =
                      columnToValues[column.powerBIName || column.name];
                    if (valueColumn) {
                      const value = valueColumn.values[index];
                      obj[column.powerBIName || column.name] = value;
                    }
                  }
                  return obj;
                })
                .filter(row => row)
            },
          tooltips &&
            tooltipsTableColumns && {
              name: powerBITooltipsTablename,
              columns: tooltipsTable ? tooltipsTable.columns : null,
              rows: categories[0].values
                .map((source, index) => {
                  const obj = {
                    _id: index.toString()
                  };
                  for (const column of tooltips) {
                    const valueColumn =
                      columnToValues[column.powerBIName || column.name];
                    if (valueColumn) {
                      const value = valueColumn.values[index];
                      obj[column.powerBIName || column.name] = value;
                    }
                  }
                  return obj;
                })
                .filter(row => row)
            }
        ].filter(table => table && table.columns)
      };
      return { dataset, rowInfo };
    }

    protected getProperties(options: VisualUpdateOptions) {
      const defaultProperties: { [name: string]: any } = {
        decimalSeparator: defaultDecimalSeparator,
        thousandSeparator: defaultThousandSeparator,
        currency: defaultCurrency,
        group: "3"
      };
      for (const p of this.template.properties as PowerBIProperty[]) {
        defaultProperties[p.powerBIName || p.displayName] = p.default;
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

      const objectKeys = Object.keys(objects);
      for (const key of objectKeys) {
        const object = objects[key];
        for (const p of this.template.properties.filter(
          p => p.objectID === key.slice(3) // remove ID_ prefix
        ) as PowerBIProperty[]) {
          if (object[p.powerBIName] != undefined) {
            if ((object[p.powerBIName] as any).solid) {
              defaultProperties[p.powerBIName] = (object[
                p.powerBIName
              ] as any).solid.color;
            } else {
              defaultProperties[p.powerBIName] = object[p.powerBIName];
            }
          }
        }
      }

      defaultProperties.decimalSeparator = objects.general?.decimalSeparator;
      defaultProperties.thousandSeparator = objects.general?.thousandSeparator;
      defaultProperties.currency = objects.general?.currency;
      defaultProperties.group = objects.general?.group;

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
        const dataView = options.dataViews[0];
        this.resize(options.viewport.width, options.viewport.height);

        const getDatasetResult = this.getDataset(dataView);

        if (getDatasetResult == null) {
          // If dataset is null, show a warning message
          this.divChart.innerHTML = `
                <h2>Dataset incomplete. Please specify all data fields.</h2>
            `;
          this.currentDatasetJSON = null;
          this.handleMouseMove = null;
          this.unmountContainer();
        } else {
          try {
            CharticulatorContainer.ChartContainer.setFormatOptions({
              currency: parseSafe(
                (dataView.metadata.objects?.[SettingsNames.General]
                  .currency as any) || defaultCurrency,
                defaultCurrency
              ),
              decimal:
                (dataView.metadata.objects?.[SettingsNames.General]
                  .decimalSeparator as any) || defaultDecimalSeparator,
              thousands:
                (dataView.metadata.objects?.[SettingsNames.General]
                  .thousandSeparator as any) || defaultThousandSeparator,
              grouping: parseSafe(
                (dataView.metadata.objects?.[SettingsNames.General]
                  .group as any) || defaultGroup,
                defaultGroup
              )
            });
          } catch (ex) {
            console.warn("Loading localization settings failed");
          }
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

            const defaultTable = this.getDefaultTable(this.template);
            const columns = defaultTable.columns as PowerBIColumn[];
            this.chartTemplate.assignTable(
              defaultTable.name,
              defaultTable.name
            );
            for (const column of columns) {
              this.chartTemplate.assignColumn(
                defaultTable.name,
                column.name,
                column.powerBIName || column.name
              );
            }

            // links table
            const linksTable = this.getLinksTable(this.template); // todo save table name in property
            const links = linksTable && (linksTable.columns as PowerBIColumn[]);
            if (links) {
              this.chartTemplate.assignTable(linksTable.name, linksTable.name);
              for (const column of links) {
                this.chartTemplate.assignColumn(
                  linksTable.name,
                  column.name,
                  column.powerBIName || column.name
                );
              }
            }

            // tooltips table
            const tooltipsTable = this.getTooltipsTable(this.template); // todo save table name in property
            const tooltips =
              tooltipsTable && (tooltipsTable.columns as PowerBIColumn[]);
            if (tooltips) {
              this.chartTemplate.assignTable(
                tooltipsTable.name,
                powerBITooltipsTablename
              );
              for (const column of tooltips) {
                this.chartTemplate.assignColumn(
                  tooltipsTable.name,
                  column.name,
                  column.powerBIName || column.name
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
              const targetProperty = property.target.property;
              if (targetProperty) {
                if (
                  typeof targetProperty === "object" &&
                  (targetProperty.property === "xData" ||
                    targetProperty.property === "yData" ||
                    targetProperty.property === "axis") &&
                  targetProperty.field === "categories"
                ) {
                  const direction = this.properties[property.powerBIName];
                  let values = CharticulatorContainer.ChartTemplate.GetChartProperty(
                    chart,
                    property.objectID,
                    {
                      property: targetProperty.property,
                      field: targetProperty.field
                    }
                  );
                  if (values) {
                    values = this.deepClone(values);
                    values = (values as string[]).sort();
                    if (direction === "descending") {
                      values = (values as string[]).reverse();
                    }
                    CharticulatorContainer.ChartTemplate.SetChartProperty(
                      chart,
                      property.objectID,
                      targetProperty,
                      values
                    );
                  }
                } else {
                  if (this.properties[property.powerBIName] != null) {
                    CharticulatorContainer.ChartTemplate.SetChartProperty(
                      chart,
                      property.objectID,
                      targetProperty,
                      this.properties[property.powerBIName]
                    );
                  }
                }
              } else {
                if (this.properties[property.powerBIName] != null) {
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
            }

            // Make selection ids:
            const selectionIDs: visuals.ISelectionId[] = [];
            const selectionID2RowIndex = new WeakMap<ISelectionId, number>();
            const rowIndex2selectionID = new Map<number, ISelectionId>();
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
              rowIndex2selectionID.set(i, selectionID);
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

            if (this.enableDrillDown) {
              const service = this.selectionManager;
              this.chartContainer.addContextMenuListener(
                // TODO change to point object
                (table, rowIndices, options) => {
                  const { clientX, clientY, event } = options;
                  if ((service as any).showContextMenu) {
                    const selection = rowIndex2selectionID.get(rowIndices[0]);
                    (service as any).showContextMenu(selection, {
                      x: clientX,
                      y: clientY
                    });
                    event.preventDefault();
                  }
                }
              );
            }
            const powerBITooltips = dataset.tables.find(
              table => table.name === powerBITooltipsTablename
            );
            const tooltipsTableColumns = [
              ...(options.dataViews[0].categorical.categories
                ? options.dataViews[0].categorical.categories
                : []),
              ...(options.dataViews[0].categorical.values
                ? options.dataViews[0].categorical.values
                : [])
            ];
            const visualHasTooltipData = tooltipsTableColumns.find(
              column => column.source.roles.powerBITooltips
            );
            if (
              this.host.tooltipService.enabled() &&
              powerBITooltips &&
              visualHasTooltipData
            ) {
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
                      const tooltiprow =
                        powerBITooltips && powerBITooltips.rows[idx];
                      const row = dataset.tables[0].rows[idx];
                      return (
                        Object.keys(tooltiprow)
                          // excule _id column
                          .filter(x => x != "_id")
                          .map(key => {
                            const header = getDatasetResult.rowInfo.get(row)
                              .granularity;
                            let value = tooltiprow[key];

                            const column = [
                              ...dataset.tables[0].columns,
                              ...(dataset.tables[1]
                                ? dataset.tables[1].columns
                                : [])
                            ].filter(n => n.name === key)[0];

                            if (
                              value !== undefined &&
                              value !== null &&
                              column
                            ) {
                              if (
                                column.type ===
                                CharticulatorContainer.Specification.DataType
                                  .Number
                              ) {
                                value = parseFloat(value + "").toFixed(2);
                              } else if (
                                column.type ===
                                CharticulatorContainer.Specification.DataType
                                  .Date
                              ) {
                                const numVal = value as number;
                                if (typeof numVal.toFixed === "function") {
                                  const parsed = new Date(numVal);
                                  value = parsed.toDateString();
                                }
                              }
                            }

                            return {
                              displayName: key,
                              header,
                              value
                            };
                          })
                      );
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
              this.chartContainer.addMouseLeaveListener(
                (table, rowIndices) => {}
              );
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
              const targetProperty = property.target.property;
              if (targetProperty) {
                if (
                  typeof targetProperty === "object" &&
                  (targetProperty.property === "xData" ||
                    targetProperty.property === "yData" ||
                    targetProperty.property === "axis") &&
                  targetProperty.field === "categories"
                ) {
                  const direction = this.properties[property.powerBIName];
                  let values = this.chartContainer.getProperty(
                    property.objectID,
                    {
                      property: targetProperty.property,
                      field: targetProperty.field
                    }
                  );
                  if (values) {
                    values = this.deepClone(values);
                    values = (values as string[]).sort();
                    if (direction === "descending") {
                      values = (values as string[]).reverse();
                    }
                    this.chartContainer.setProperty(
                      property.objectID,
                      targetProperty,
                      values
                    );
                  }
                } else {
                  if (this.properties[property.powerBIName] != null) {
                    this.chartContainer.setProperty(
                      property.objectID,
                      targetProperty,
                      this.properties[property.powerBIName]
                    );
                  }
                }
              }
              if (
                property.target.attribute &&
                this.properties[property.powerBIName] != null
              ) {
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
                const defaultTable = this.getDefaultTable(this.template);
                this.chartContainer.setSelection(defaultTable.name, indices);
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
      const objectName = options.objectName as string;
      const objectEnumeration = [];
      const properties: { [name: string]: any } = {};
      const templateProperties = this.template.properties.filter(
        p => p.objectID === objectName.replace(propertyAndObjectNamePrefix, "")
      ) as PowerBIProperty[];

      if (objectName === SettingsNames.General) {
        objectEnumeration.push({
          objectName,
          properties: {
            decimalSeparator:
              this.properties.decimalSeparator || defaultDecimalSeparator,
            thousandSeparator:
              this.properties.thousandSeparator || defaultThousandSeparator,
            currency: this.properties.currency || defaultCurrency,
            group: this.properties.group || defaultGroup
          },
          selector: null
        });
      } else {
        for (const p of templateProperties) {
          if (this.properties[p.powerBIName] !== undefined) {
            if (
              p.displayName.indexOf("xData.categories") > -1 ||
              p.displayName.indexOf("yData.categories") > -1 ||
              p.displayName.indexOf("axis.categories") > -1
            ) {
              const values = this.chartContainer.getProperty(p.objectID, {
                property: (p.target.property as any).property,
                field: (p.target.property as any).field
              });
              if (values) {
                const a = values[0].toString();
                const b = values[(values as any[]).length - 1].toString();
                if (b.localeCompare(a) > -1) {
                  properties[p.powerBIName] = "ascending";
                } else {
                  properties[p.powerBIName] = "descending";
                }
              } else {
                properties[p.powerBIName] = "ascending";
              }
            } else {
              properties[p.powerBIName] = this.properties[p.powerBIName];
            }
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
        objectEnumeration.push({ objectName, properties, selector: null });
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
