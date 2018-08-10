Charticulator Extension: Power BI Visual Builder
====

This extension adds the "Export as Power BI Custom Visual" functionality.

## How it Works

Charticulator provides a `container.bundle.min.js` that allows us to import a chart template and feed data into the template to create visualizations.
This extension works by producing a Power BI compatible `.pbiviz` file that bundles the `container.bundle.min.js` and some bridging code (see `template_code/visual.js`).

## Build


Build the Charticulator source code first, then, run the following:

```bash
# Install dependencies
npm install

# Build the extension
npm run build
```

## Usage

Copy the `dist/powerbi_visual_builder.js` into the `extensions` folder in Charticulator's source code.

Edit Charticulator's `config.yml`:

```yaml
Extensions:
  - script: extensions/powerbi_visual_builder.js
    initialize: |
      let extension = new CharticulatorPowerBIVisualBuilder.PowerBIVisualBuilder();
      application.addExtension(extension);
```
