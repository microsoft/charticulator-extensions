// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// The main file for the visual

namespace powerbi.extensibility.visual {
  const simpleBarChart = '{"specification":{"_id":"5iippmyucb","classID":"chart.rectangle","properties":{"name":"Chart","backgroundColor":null,"backgroundOpacity":1},"mappings":{"marginTop":{"type":"value","value":80}},"glyphs":[{"_id":"80erjnptnzt","classID":"glyph.rectangle","properties":{"name":"Glyph"},"table":"goals by teams","marks":[{"_id":"q23wew7bl7n","classID":"mark.anchor","properties":{"name":"Anchor"},"mappings":{"x":{"type":"parent","parentAttribute":"icx"},"y":{"type":"parent","parentAttribute":"icy"}}},{"_id":"2i1k05stiel","classID":"mark.rect","properties":{"visible":true,"shape":"rectangle","name":"Shape1"},"mappings":{"fill":{"type":"scale","table":"goals by teams","expression":"first(Nation)","valueType":"string","scale":"tbwtw9f0vl"},"strokeWidth":{"type":"value","value":1},"opacity":{"type":"value","value":1},"visible":{"type":"value","value":true},"x1":{"type":"parent","parentAttribute":"ix1"},"y1":{"type":"parent","parentAttribute":"iy1"},"x2":{"type":"parent","parentAttribute":"ix2"},"y2":{"type":"parent","parentAttribute":"iy2"},"height":{"type":"scale","table":"goals by teams","expression":"avg(Goals)","valueType":"number","scale":"y9cmv0l8nk"}}}],"mappings":{},"constraints":[]}],"elements":[{"_id":"6zmj599u3qp","classID":"plot-segment.cartesian","glyph":"80erjnptnzt","table":"goals by teams","filter":null,"mappings":{"x1":{"type":"parent","parentAttribute":"x1"},"y1":{"type":"parent","parentAttribute":"y1"},"x2":{"type":"parent","parentAttribute":"x2"},"y2":{"type":"parent","parentAttribute":"y2"}},"properties":{"name":"PlotSegment1","visible":true,"marginX1":0,"marginY1":0,"marginX2":0,"marginY2":0,"sublayout":{"type":"dodge-x","order":null,"ratioX":0.1,"ratioY":0.1,"align":{"x":"start","y":"start"},"grid":{"direction":"x","xCount":null,"yCount":null}},"xData":{"type":"categorical","expression":"first(Team)","valueType":"string","gapRatio":0.1,"visible":true,"side":"default","style":{"tickColor":{"r":0,"g":0,"b":0},"lineColor":{"r":0,"g":0,"b":0},"fontFamily":"Arial","fontSize":12,"tickSize":5},"categories":["Barcelona","Juve","Liverpol","Macabi Haifa","Man City","Man United","P.S.G","Real Madrid"]}}},{"_id":"sxl5hm9m3d","classID":"mark.text","properties":{"name":"Title","visible":true,"alignment":{"x":"middle","y":"top","xMargin":0,"yMargin":30},"rotation":0},"mappings":{"x":{"type":"parent","parentAttribute":"cx"},"y":{"type":"parent","parentAttribute":"oy2"},"text":{"type":"value","value":"goals by teams"},"fontSize":{"type":"value","value":24},"color":{"type":"value","value":{"r":0,"g":0,"b":0}}}}],"scales":[{"_id":"y9cmv0l8nk","classID":"scale.linear<number,number>","properties":{"name":"Scale1","domainMin":0,"domainMax":150},"mappings":{"rangeMin":{"type":"value","value":0}},"inputType":"number","outputType":"number"},{"_id":"tbwtw9f0vl","classID":"scale.categorical<string,color>","properties":{"name":"Scale2","mapping":{"Spain":{"r":166,"g":206,"b":227},"Israel":{"r":31,"g":120,"b":180},"England":{"r":178,"g":223,"b":138},"Italy":{"r":51,"g":160,"b":44},"Franch":{"r":251,"g":154,"b":153}}},"mappings":{},"inputType":"string","outputType":"color"}],"constraints":[],"resources":[]},"defaultAttributes":{"q23wew7bl7n":{"x":0,"y":0},"2i1k05stiel":{"x1":-4.500000000000001,"y1":-16.685000016277478,"x2":4.500000000000001,"y2":16.685000016277478,"cx":0,"cy":0,"width":9.000000000000002,"height":33.370000032554955,"strokeWidth":0.1,"opacity":0.1},"sxl5hm9m3d":{"x":0,"y":75,"fontSize":6,"opacity":0.25}},"tables":[{"name":"goals by teams","columns":[{"displayName":"Category","name":"Team","type":"string","metadata":{"kind":"categorical","orderMode":"alphabetically"},"powerBIName":"Team"},{"displayName":"Measure","name":"Goals","type":"number","metadata":{"kind":"numerical"},"powerBIName":"Goals"},{"displayName":"Legend","name":"Nation","type":"string","metadata":{"kind":"categorical","orderMode":"alphabetically"},"powerBIName":"Nation"}]}],"inference":[{"objectID":"6zmj599u3qp","dataSource":{"table":"goals by teams"},"axis":{"expression":"first(Team)","type":"categorical","property":"xData"}},{"objectID":"y9cmv0l8nk","scale":{"classID":"scale.linear<number,number>","expressions":["avg(Goals)"],"properties":{"mapping":"mapping"}},"dataSource":{"table":"goals by teams"}},{"objectID":"tbwtw9f0vl","scale":{"classID":"scale.categorical<string,color>","expressions":["first(Nation)"],"properties":{"mapping":"mapping"}},"dataSource":{"table":"goals by teams"}}],"properties":[{"objectID":"sxl5hm9m3d","target":{"attribute":"text"},"type":"text","default":"goals by teams","displayName":"Title/text","powerBIName":"sxl5hm9m3dTitle_text"},{"objectID":"y9cmv0l8nk","target":{"property":"domainMin"},"type":"number","displayName":"Scale1/domainMin","powerBIName":"y9cmv0l8nkScale1_domainMin"},{"objectID":"y9cmv0l8nk","target":{"property":"domainMax"},"type":"number","displayName":"Scale1/domainMax","powerBIName":"y9cmv0l8nkScale1_domainMax"},{"objectID":"5iippmyucb","target":{"attribute":"marginLeft"},"type":"number","default":50,"displayName":"Chart/marginLeft","powerBIName":"5iippmyucbChart_marginLeft"},{"objectID":"5iippmyucb","target":{"attribute":"marginRight"},"type":"number","default":50,"displayName":"Chart/marginRight","powerBIName":"5iippmyucbChart_marginRight"},{"objectID":"5iippmyucb","target":{"attribute":"marginTop"},"type":"number","default":80,"displayName":"Chart/marginTop","powerBIName":"5iippmyucbChart_marginTop"},{"objectID":"5iippmyucb","target":{"attribute":"marginBottom"},"type":"number","default":50,"displayName":"Chart/marginBottom","powerBIName":"5iippmyucbChart_marginBottom"}]}';
  const simplePieChart = '{"specification":{"_id":"3vcx2sdvqoj","classID":"chart.rectangle","properties":{"name":"Chart","backgroundColor":null,"backgroundOpacity":1},"mappings":{"marginTop":{"type":"value","value":80},"marginRight":{"type":"value","value":100}},"glyphs":[{"_id":"0f9ijim4qhsj","classID":"glyph.rectangle","properties":{"name":"Glyph"},"table":"goals by teams","marks":[{"_id":"rry0dkp9tm","classID":"mark.anchor","properties":{"name":"Anchor"},"mappings":{"x":{"type":"parent","parentAttribute":"icx"},"y":{"type":"parent","parentAttribute":"icy"}}},{"_id":"2u0tv82igtk","classID":"mark.rect","properties":{"visible":true,"shape":"rectangle","name":"Shape1"},"mappings":{"fill":{"type":"scale","table":"goals by teams","expression":"first(Nation)","valueType":"string","scale":"0gsibynypbt"},"strokeWidth":{"type":"value","value":1},"opacity":{"type":"value","value":1},"visible":{"type":"value","value":true},"x1":{"type":"parent","parentAttribute":"ix1"},"y1":{"type":"parent","parentAttribute":"iy1"},"x2":{"type":"parent","parentAttribute":"ix2"},"y2":{"type":"parent","parentAttribute":"iy2"},"height":{"type":"scale","table":"goals by teams","expression":"avg(Goals)","valueType":"number","scale":"cxxix9fnz0q"},"stroke":{"type":"value","value":{"r":0,"g":0,"b":0}}}}],"mappings":{},"constraints":[]}],"elements":[{"_id":"9k1i0oi0xr","classID":"plot-segment.polar","glyph":"0f9ijim4qhsj","table":"goals by teams","filter":null,"mappings":{"x1":{"type":"parent","parentAttribute":"x1"},"x2":{"type":"parent","parentAttribute":"x2"},"y1":{"type":"parent","parentAttribute":"y1"},"y2":{"type":"parent","parentAttribute":"y2"}},"properties":{"name":"PlotSegment1","visible":true,"sublayout":{"type":"dodge-x","order":null,"ratioX":0.1,"ratioY":0.1,"align":{"x":"start","y":"start"},"grid":{"direction":"x","xCount":null,"yCount":null}},"xData":{"type":"categorical","expression":"first(Team)","valueType":"string","gapRatio":0,"visible":true,"side":"default","style":{"tickColor":{"r":0,"g":0,"b":0},"lineColor":{"r":0,"g":0,"b":0},"fontFamily":"Arial","fontSize":12,"tickSize":5},"categories":["Barcelona","Juve","Liverpol","Macabi Haifa","Man City","Man United","P.S.G","Real Madrid"]},"marginX1":0,"marginY1":0,"marginX2":0,"marginY2":0,"startAngle":0,"endAngle":360,"innerRatio":0,"outerRatio":0.9,"curve":[[{"x":-1,"y":0},{"x":-0.25,"y":-0.5},{"x":0.25,"y":0.5},{"x":1,"y":0}]],"normalStart":-0.2,"normalEnd":0.2,"equalizeArea":true}},{"_id":"yvjnlby981","classID":"mark.text","properties":{"name":"Title","visible":true,"alignment":{"x":"middle","y":"top","xMargin":0,"yMargin":30},"rotation":0},"mappings":{"x":{"type":"parent","parentAttribute":"cx"},"y":{"type":"parent","parentAttribute":"oy2"},"text":{"type":"value","value":"goals by teams"},"fontSize":{"type":"value","value":24},"color":{"type":"value","value":{"r":0,"g":0,"b":0}}}},{"_id":"tkqoh0cm6m","classID":"legend.categorical","properties":{"visible":true,"alignX":"end","alignY":"end","fontFamily":"Arial","fontSize":14,"textColor":{"r":0,"g":0,"b":0},"name":"Legend1","scale":"0gsibynypbt"},"mappings":{"x":{"type":"parent","parentAttribute":"x2"},"y":{"type":"parent","parentAttribute":"y2"}}}],"scales":[{"_id":"0gsibynypbt","classID":"scale.categorical<string,color>","properties":{"name":"Scale1","mapping":{"Spain":{"r":166,"g":206,"b":227},"Israel":{"r":31,"g":120,"b":180},"England":{"r":178,"g":223,"b":138},"Italy":{"r":51,"g":160,"b":44},"Franch":{"r":251,"g":154,"b":153}}},"mappings":{},"inputType":"string","outputType":"color"},{"_id":"cxxix9fnz0q","classID":"scale.linear<number,number>","properties":{"name":"Scale2","domainMin":0,"domainMax":150},"mappings":{"rangeMin":{"type":"value","value":0}},"inputType":"number","outputType":"number"}],"constraints":[],"resources":[]},"defaultAttributes":{"rry0dkp9tm":{"x":0,"y":0},"2u0tv82igtk":{"x1":-2.25,"y1":-7.5082500073248655,"x2":2.25,"y2":7.5082500073248655,"cx":0,"cy":0,"width":4.5,"height":15.016500014649731,"strokeWidth":0.1,"opacity":0.1},"yvjnlby981":{"x":-6.25,"y":75,"fontSize":6,"opacity":0.25}},"tables":[{"name":"goals by teams","columns":[{"displayName":"Category","name":"Team","type":"string","metadata":{"kind":"categorical","orderMode":"alphabetically"},"powerBIName":"Team"},{"displayName":"Measure","name":"Goals","type":"number","metadata":{"kind":"numerical"},"powerBIName":"Goals"},{"displayName":"Legend","name":"Nation","type":"string","metadata":{"kind":"categorical","orderMode":"alphabetically"},"powerBIName":"Nation"}]}],"inference":[{"objectID":"9k1i0oi0xr","dataSource":{"table":"goals by teams"},"axis":{"expression":"first(Team)","type":"categorical","property":"xData"}},{"objectID":"0gsibynypbt","scale":{"classID":"scale.categorical<string,color>","expressions":["first(Nation)"],"properties":{"mapping":"mapping"}},"dataSource":{"table":"goals by teams"}},{"objectID":"cxxix9fnz0q","scale":{"classID":"scale.linear<number,number>","expressions":["avg(Goals)"],"properties":{"mapping":"mapping"}},"dataSource":{"table":"goals by teams"}}],"properties":[{"objectID":"yvjnlby981","target":{"attribute":"text"},"type":"text","default":"goals by teams","displayName":"Title/text","powerBIName":"yvjnlby981Title_text"},{"objectID":"cxxix9fnz0q","target":{"property":"domainMin"},"type":"number","displayName":"Scale2/domainMin","powerBIName":"cxxix9fnz0qScale2_domainMin"},{"objectID":"cxxix9fnz0q","target":{"property":"domainMax"},"type":"number","displayName":"Scale2/domainMax","powerBIName":"cxxix9fnz0qScale2_domainMax"},{"objectID":"3vcx2sdvqoj","target":{"attribute":"marginLeft"},"type":"number","default":50,"displayName":"Chart/marginLeft","powerBIName":"3vcx2sdvqojChart_marginLeft"},{"objectID":"3vcx2sdvqoj","target":{"attribute":"marginRight"},"type":"number","default":100,"displayName":"Chart/marginRight","powerBIName":"3vcx2sdvqojChart_marginRight"},{"objectID":"3vcx2sdvqoj","target":{"attribute":"marginTop"},"type":"number","default":80,"displayName":"Chart/marginTop","powerBIName":"3vcx2sdvqojChart_marginTop"},{"objectID":"3vcx2sdvqoj","target":{"attribute":"marginBottom"},"type":"number","default":50,"displayName":"Chart/marginBottom","powerBIName":"3vcx2sdvqojChart_marginBottom"}]}';
  
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
    private highlights: any[];
    private lastOptions: VisualUpdateOptions;
    private selectionIds: ISelectionId[];
    private currentSelections: ISelectionId[];
    private firstUpdate: boolean;
    protected template2: any; // for testing

    constructor(options: VisualConstructorOptions) {
      try {
        this.selectionManager = options.host.createSelectionManager();
        this.selectionManager.registerOnSelectCallback((selectionsIds) => {this.onSelectCallBack(selectionsIds)});
        debugger;
        this.firstUpdate = true;

        this.template2 = "<%= templateData %>" as any;
        this.enableTooltip = "<%= enableTooltip %>" as any;
        this.container = options.element;
        this.divChart = document.createElement("div");
        this.divChart.style.cursor = "default";
        this.divChart.style.pointerEvents = "none";
        this.host = options.host;
        options.element.appendChild(this.divChart);

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

    public updateSelections() {
      debugger;
      // Idan - add interactivity utils
      const highlights = [];
      const selecionsID = this.selectionManager.getSelectionIds();
      if (selecionsID && selecionsID[0]){
        for (let i = 0; i < this.selectionIds.length; i++) {
          highlights.push(null);
          for (let j = 0; j < selecionsID.length; j++) {
            if ((this.selectionIds[i] as any).includes(selecionsID[j])){
              highlights[i] = true
            }
          }
        }       
      }
      this.highlights = highlights;
    }

    public onSelectCallBack(selecionsID:any[]) {
      debugger;
      if (this.lastOptions){
        this.update(this.lastOptions);
      }
      console.log("test")
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
      // debugger;
      const inDataColumns = options.dataViews[0].table.columns;
      const inDataRows = options.dataViews[0].table.rows;

      // Match columns
      const columnToValues: {
        [name: string]: {
          values: CharticulatorContainer.Specification.DataValue[];
          highlights?: boolean[];
        };
      } = {};
      const columns = this.template.tables[0].columns as PowerBIColumn[];
      for (const column of columns) {
        let found = false;
        if (inDataColumns != null) {
          // for (const inDataOneColumn of inDataColumns) {
          for (let i = 0; i < inDataColumns.length; i++) {
            if (inDataColumns[i].roles[column.powerBIName]) {
              const rowI = inDataRows.map(values => values[i]);
              columnToValues[column.powerBIName] = {
                values: CharticulatorContainer.Dataset.convertColumnType(
                  rowI,
                  column.type
                ),
                highlights: this.highlights.map((value, i) => {
                  return this.highlights
                    ? this.highlights[i] != null && value != null
                      ? this.highlights[i].valueOf() == value.valueOf()
                      : false
                    : false;
                })
              };
              found = true;
            }
          }
        }
        if (!found) {
          return null;
        }
      }

      const rowInfo = new Map<
        CharticulatorContainer.Dataset.Row,
        { highlight: boolean; index: number; granularity: string }
      >();
      const dataset: any = {
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
            rows: inDataRows.map((rowi, i) => {
              const chartDataSetRowI: CharticulatorContainer.Dataset.Row = {
                _id: "ID" + i.toString()
              };
              for (const column of columns) {
                chartDataSetRowI[column.name] =
                  rowi[this.getColumnIndex(column.name, inDataColumns)];
              }

              // set the correct type here

              rowInfo.set(chartDataSetRowI, {
                highlight:
                  this.highlights && this.highlights[i] != null ? true : false,
                index: i,
                granularity: null // Idan -  maybe not needed!!!
              });

              return chartDataSetRowI;
            })
          }
        ]
      };
      // debugger;;
      // this.getTeamColors(inDataRows,this.getColumnIndex("Nation",inDataColumns))
      return { dataset, rowInfo };
    }

    private getColumnIndex(
      charticulatorColumnName: string,
      inDataColumns: any
    ): number {
      for (let i = 0; i < inDataColumns.length; i++) {
        if (inDataColumns[i].roles[charticulatorColumnName]) {
          return i;
        }
      }

      return -1;
    }

    private getTeamColors(row: any[], ColumnIndex: number): void {
      const colors = {};
      for (let i = 0; i < row.length; i++) {
        const min = 0;
        const max = 256;
        const r = Math.floor(Math.random() * (+max - +min)) + +min;
        const g = Math.floor(Math.random() * (+max - +min)) + +min;
        const b = Math.floor(Math.random() * (+max - +min)) + +min;
        const color = { r, g, b };
        colors[row[i][ColumnIndex]] = color;
      }

      const specification = (this.chartTemplate as any).template.specification;
      for (let i = 0; i < specification.scales.length; i++) {
        if (specification.scales[i].properties.mapping) {
          specification.scales[i].properties.mapping = colors;
        }
      }
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
      debugger;
      if (this.firstUpdate){
        // let number = Math.floor(Math.random() * 2);
        // if (number == 1){
        //   this.template = JSON.parse(simpleBarChart);
        // } else {
        //   this.template = JSON.parse(simplePieChart);
        // }
        // this.chartTemplate = new CharticulatorContainer.ChartTemplate(
        //   this.template
        // );

        this.template = this.template2;
        
        this.chartTemplate = new CharticulatorContainer.ChartTemplate(
          this.template
        );
        this.properties = this.getProperties(null);
      }
      this.firstUpdate = false;

      this.lastOptions = options;
      this.highlights = (options as any).highlights;
      if(!this.highlights || this.highlights.length ==0){
          this.updateSelections();
      }
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
            // debugger;

            this.chartContainer = new CharticulatorContainer.ChartContainer(
              instance,
              dataset
            );

            // Idan - use new interface!

            // Make selection ids:
            const selectionIDs: visuals.ISelectionId[] = [];
            const selectionID2RowIndex = new WeakMap<ISelectionId, number>();
            dataset.tables[0].rows.forEach((row, i) => {
              const selectionID = (this.host as any)
                .createSelectionIdBuilder()
                .withIdentity(
                  options.dataViews[0].table.columns[0].queryName,
                  options.dataViews[0].table.identity[i]
                )
                .createSelectionId();
              selectionIDs.push(selectionID);
              selectionID2RowIndex.set(selectionID, i);
            });
            this.selectionIds = selectionIDs;
            this.chartContainer.addSelectionListener((table, rowIndices) => {
              debugger;
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

            /* tooltip part! */

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
                          // fix tooltip issue when value is a number
                          value = value ? value.toString() : null;
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
