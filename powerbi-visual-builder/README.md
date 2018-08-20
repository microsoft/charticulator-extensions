Charticulator Extension: Power BI Visual Builder
====

This extension adds the "Export as Power BI Custom Visual" functionality.

## How it Works

Charticulator provides a `container.bundle.min.js` that allows us to import a chart template and feed data into it to create visualizations. This extension works by producing a Power BI compatible `.pbiviz` file that bundles the `container.bundle.min.js` and some bridging code (see `template_code/visual.js`).

## Requirements
* Node 8+: https://nodejs.org
* Yarn 1.5.1+: https://yarnpkg.com/

## Build

```bash
# Install dependencies
yarn

# Build the extension
yarn build
```

## Usage

Copy the `dist/powerbi_visual_builder.js` into the `extensions` folder in Charticulator's source code.

Edit Charticulator's `config.yml`:

```yaml
Extensions:
  - script: extensions/powerbi_visual_builder.js
    initialize: |
      let extension = new CharticulatorPowerBIVisualBuilder.PowerBIVisualBuilder('scripts/container.bundle.js');
      application.addExtension(extension);
```

Enjoy!