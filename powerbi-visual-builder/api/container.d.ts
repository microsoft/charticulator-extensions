// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

// TODO: need a tool to generate this automatically

declare module CharticulatorContainer {
  function initialize(options: any): Promise<void>;

  /** Represents a chart template */
  class ChartTemplate {
    /** Create a chart template */
    constructor(template: Specification.Template.ChartTemplate);
    getDatasetSchema(): Specification.Template.Table[];
    /** Reset slot assignments */
    reset(): void;
    /** Assign a table */
    assignTable(tableName: string, table: string): void;
    /** Assign an expression to a data mapping slot */
    assignColumn(tableName: string, columnName: string, column: string): void;
    /** Get variable map for a given table */
    getVariableMap(table: string): {
      [name: string]: string;
    };
    transformExpression(expr: string, table?: string): string;
    transformTextExpression(expr: string, table?: string): string;
    transformGroupBy(groupBy: Specification.Types.GroupBy, table: string): Specification.Types.GroupBy;
    instantiate(dataset: Dataset.Dataset): ChartTemplateInstance;
  }
  class ChartTemplateInstance {
    readonly chart: Specification.Chart;
    readonly dataset: Dataset.Dataset;
    constructor(chart: Specification.Chart, dataset: Dataset.Dataset);
    getProperties(): void;
    getProperty(id: string): void;
    setProperty(id: string, value: any): void;
  }
  class ChartContainer {
    readonly chart: Specification.Chart;
    readonly dataset: Dataset.Dataset;
    constructor(chart: Specification.Chart, dataset: Dataset.Dataset);
    /** Resize the chart */
    resize(width: number, height: number): void;
    /** Listen to selection change */
    addSelectionListener(listener: (table: string, rowIndices: number[]) => void): EventSubscription;
    /** Set data selection and update the chart */
    setSelection(table: string, rowIndices: number[]): void;
    /** Clear data selection and update the chart */
    clearSelection(): void;
    /** Mount the chart to a container element */
    mount(container: string | Element, width?: number, height?: number): void;
    /** Unmounr the chart */
    unmount(): void;
  }

  interface Color {
    r: number;
    g: number;
    b: number;
  }

  interface Point {
    x: number;
    y: number;
  }

  interface EventSubscription {
    remove(): void;
  }

  module Dataset {
    type ValueType = string | number | Date | boolean;
    interface Dataset {
      /** Name of the dataset */
      name: string;
      /** Tables in the dataset */
      tables: Table[];
    }
    interface ColumnMetadata {
      /** Conceptural data type: categorical (including ordinal), numerical, text, boolean */
      kind: string;
      /** The unit of the data type, used in scale inference when mapping multiple columns */
      unit?: string;
      /** Order of categories for categorical type */
      order?: string[];
      orderMode?: "alphabetically" | "occurrence" | "order";
      /** Formatting for other data types */
      format?: string;
    }
    interface Column {
      /** Name of the column, used to address the entry from row */
      name: string;
      /** Data type in memory (number, string, Date, boolean, etc) */
      type: string;
      /** Metadata on this column */
      metadata: ColumnMetadata;
    }
    interface Row {
      /** Internal row ID, automatically assigned to be unique */
      _id: string;
      /** Row attributes */
      [name: string]: ValueType;
    }
    interface Table {
      /** Table name */
      name: string;
      /** Columns in the table */
      columns: Column[];
      /** Rows in the table */
      rows: Row[];
    }
  }

  module Specification {
    /** Objects with an unique ID */
    interface Identifiable {
      /** Unique ID */
      _id: string;
    }
    /** Supported data value types */
    type DataValue = number | string | boolean | Date;
    /** Data row */
    interface DataRow {
      _id: string;
      [name: string]: DataValue;
    }
    type Expression = string;
    /** Attribute value types */
    type AttributeValue = number | string | boolean | Color | Point | AttributeList | AttributeMap;
    /** Attribute value list */
    interface AttributeList extends ArrayLike<AttributeValue> {
    }
    /** Attribute value map */
    interface AttributeMap {
      [name: string]: AttributeValue;
    }
    /** Attribute mappings */
    interface Mappings {
      [name: string]: Mapping;
    }
    /** Attribute mapping */
    interface Mapping {
      /** Mapping type */
      type: string;
    }
    /** Scale mapping: use a scale */
    interface ScaleMapping extends Mapping {
      type: "scale";
      /** The table to draw data from */
      table: string;
      /** The data column */
      expression: Expression;
      /** Value type */
      valueType: string;
      /** The id of the scale to use. If null, use the expression directly */
      scale?: string;
    }
    /** Text mapping: map data to text */
    interface TextMapping extends Mapping {
      type: "text";
      /** The table to draw data from */
      table: string;
      /** The text expression */
      textExpression: string;
    }
    /** Value mapping: a constant value */
    interface ValueMapping extends Mapping {
      type: "value";
      /** The constant value */
      value: AttributeValue;
    }
    /** Parent mapping: use an attribute of the item's parent item */
    interface ParentMapping extends Mapping {
      type: "parent";
      /** The attribute of the parent item */
      parentAttribute: string;
    }
    /** Constraint */
    interface Constraint {
      /** Constraint type */
      type: string;
      attributes: AttributeMap;
    }
    /** Object attributes */
    interface ObjectProperties extends AttributeMap {
      /** The name of the object, used in UI */
      name?: string;
      visible?: boolean;
      emphasisMethod?: EmphasisMethod;
    }
    /** General object */
    interface Object extends Identifiable {
      /** The class ID for the Object */
      classID: string;
      /** Attributes  */
      properties: ObjectProperties;
      /** Scale attribute mappings */
      mappings: Mappings;
    }
    /** Element: a single graphical mark, such as rect, circle, wedge; an element is driven by a single data row */
    interface Element extends Object {
    }
    /** Glyph: a compound of elements, with constraints between them; a glyph is driven by a single data row */
    interface Glyph extends Object {
      /** The data table this mark correspond to */
      table: string;
      /** Elements within the mark */
      marks: Element[];
      /** Layout constraints for this mark */
      constraints: Constraint[];
    }
    /** Scale */
    interface Scale extends Object {
      inputType: string;
      outputType: string;
    }
    /** MarkLayout: the "PlotSegment" */
    interface PlotSegment extends Object {
      /** The mark to use */
      glyph: string;
      /** The data table to get data rows from */
      table: string;
      /** Filter applied to the data table */
      filter?: Types.Filter;
      /** Group the data by a specified categorical column (filter is applied before grouping) */
      groupBy?: Types.GroupBy;
      /** Order the data (filter & groupBy is applied before order */
      order?: Types.SortBy;
    }
    /** Guide */
    interface Guide extends Object {
    }
    /** Guide Coordinator */
    interface GuideCoordinator extends Object {
    }
    /** Links */
    interface Links extends Object {
    }
    /** ChartElement is a PlotSegment or a Guide */
    type ChartElement = PlotSegment | Guide | GuideCoordinator;
    /** Resource item */
    interface Resource {
      /** Resource item ID */
      id: string;
      /** Resource type: image */
      type: string;
      /** Resource data */
      data: any;
    }
    /** A chart is a set of chart elements and constraints between them, with guides and scales */
    interface Chart extends Object {
      /** Marks */
      glyphs: Glyph[];
      /** Scales */
      scales: Scale[];
      /** Chart elements */
      elements: ChartElement[];
      /** Chart-level constraints */
      constraints: Constraint[];
      /** Resources */
      resources: Resource[];
    }
    /** General object state */
    interface ObjectState {
      attributes: AttributeMap;
    }
    /** Element state */
    interface MarkState extends ObjectState {
    }
    /** Scale state */
    interface ScaleState extends ObjectState {
    }
    /** Glyph state */
    interface GlyphState extends ObjectState {
      marks: MarkState[];
      /**
          * Should this specific glyph instance be emphasized
          */
      emphasized?: boolean;
    }
    /** PlotSegment state */
    interface PlotSegmentState extends ObjectState {
      glyphs: GlyphState[];
      dataRowIndices: number[][];
    }
    /** Guide state */
    interface GuideState extends ObjectState {
    }
    /** Chart element state, one of PlotSegmentState or GuideState */
    type ChartElementState = PlotSegmentState | GuideState;
    /** Chart state */
    interface ChartState extends ObjectState {
      /** Mark binding states corresponding to Chart.marks */
      elements: ChartElementState[];
      /** Scale states corresponding to Chart.scales */
      scales: ScaleState[];
    }
    /**
        * Represents the type of method to use when emphasizing an element
        */
    enum EmphasisMethod {
      Saturation = "saturatation",
      Outline = "outline",
    }

    module Template {
      type PropertyField = string | {
        property: string;
        field: any;
      };
      interface ChartTemplate {
        /** The original chart specification */
        specification: Chart;
        /** Data tables */
        tables: Table[];
        /** Infer attribute or property from data */
        inference: Inference[];
        /** Expose property editor */
        properties: Property[];
      }
      interface Column {
        displayName: string;
        name: string;
        type: string;
        metadata: Dataset.ColumnMetadata;
      }
      interface Table {
        name: string;
        columns: Column[];
      }
      interface Property {
        objectID: string;
        displayName?: string;
        target: {
          property?: PropertyField;
          attribute?: string;
        };
        type: string;
        default?: string | number | boolean;
      }
      /** Infer values from data */
      interface Inference {
        objectID: string;
        dataSource?: {
          table: string;
          groupBy?: Types.GroupBy;
        };
        axis?: AxisInference;
        scale?: ScaleInference;
        expression?: ExpressionInference;
      }
      /** Infer axis parameter, set to axis property */
      interface AxisInference {
        /** Data expression for the axis */
        expression: string;
        /** Type */
        type: "default" | "categorical" | "numerical";
        /** Infer axis data and assign to this property */
        property: PropertyField;
      }
      /** Infer scale parameter, set to scale's domain property */
      interface ScaleInference {
        classID: string;
        expressions: string[];
        properties: {
          min?: PropertyField;
          max?: PropertyField;
          mapping?: PropertyField;
        };
      }
      /** Fix expression */
      interface ExpressionInference {
        expression: string;
        property: PropertyField;
      }
    }
    module Types {
      interface AxisDataBinding extends AttributeMap {
        type: "default" | "numerical" | "categorical";
        visible: boolean;
        side: "default" | "opposite";
        /** Data mapping expression */
        expression?: Expression;
        valueType?: string;
        /** Domain for linear/logarithm types */
        numericalMode?: "linear" | "logarithm";
        domainMin?: number;
        domainMax?: number;
        /** Categories for categorical type */
        categories?: string[];
        gapRatio?: number;
        /** Pre/post gap, will override the default with OR operation */
        enablePrePostGap?: boolean;
        tickDataExpression?: Expression;
        style?: AxisRenderingStyle;
      }
      interface AxisRenderingStyle extends AttributeMap {
        lineColor: Color;
        tickColor: Color;
        fontFamily: string;
        fontSize: number;
        tickSize: number;
      }
      interface TextAlignment extends AttributeMap {
        x: "left" | "middle" | "right";
        y: "top" | "middle" | "bottom";
        xMargin: number;
        yMargin: number;
      }
      interface ColorGradient extends AttributeMap {
        colorspace: "hcl" | "lab";
        colors: Color[];
      }
      /** LinkAnchor: specifies an anchor in a link */
      interface LinkAnchorPoint extends AttributeMap {
        /** X attribute reference */
        x: {
          element: string;
          attribute: string;
        };
        /** Y attribute reference */
        y: {
          element: string;
          attribute: string;
        };
        /** Link direction for curves */
        direction: {
          x: number;
          y: number;
        };
      }
      /** Filter specification, specify one of categories or expression */
      interface Filter extends AttributeMap {
        /** Filter by a categorical variable */
        categories?: {
          /** The expression to draw values from */
          expression: string;
          /** The accepted values */
          values: {
            [value: string]: boolean;
          };
        };
        /** Filter by an arbitrary expression */
        expression?: Expression;
      }
      /** GroupBy specification */
      interface GroupBy extends AttributeMap {
        /** Group by a string expression */
        expression?: Expression;
      }
      /** Order expression */
      interface SortBy extends AttributeMap {
        expression?: Expression;
      }
    }
  }
}