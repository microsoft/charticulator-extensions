const fs = require("fs-extra");
const multirun = require("multirun");

let isProd = false;
let sequence = [];
process.argv.slice(2).forEach(arg => {
  const m = arg.match(/^--([0-9a-zA-Z]+)\=(.*)$/);
  if (m) {
    const name = m[1], value = m[2];
    if (name == "mode") {
      isProd = value == "production";
    }
  } else {
    sequence.push(arg);
  }
});

// The default dev sequence
const devSequence = ["cleanup", "typescript", "resources", "webpack"];

let COMMANDS = {
  cleanup: () => fs.remove("dist"),

  generate_schema_dts: [
    "json2ts -i api/v2.1.0/schema.capabilities.json -o api/v2.1.0/schema.capabilities.d.ts",
    "json2ts -i api/v2.1.0/schema.dependencies.json -o api/v2.1.0/schema.dependencies.d.ts",
    "json2ts -i api/v2.1.0/schema.pbiviz.json -o api/v2.1.0/schema.pbiviz.d.ts"
  ],

  typescript: [
    "tsc -p src_extension",
    "tsc -p src_visual"
  ],

  webpack: "webpack --mode=development",

  resources: async () => {
    let resources = {
      icon: "data:image/png;base64," + await fs.readFile("assets/icon.png", "base64"),
      libraries: "",
      visual: await fs.readFile("dist/visual.js", "utf-8")
    };
    // TODO: this was an unsuccessful attempt to let the exported visual support IE11.
    // let babel = require("babel-core");
    // let enableBabel = false;
    // if (enableBabel) {
    //     function makePowerBICompatible(code) {
    //         return babel.transform(code, {
    //             presets: ["es2015-ie"]
    //         }).code;
    //     }
    //     // A mirrored window object is used in Power BI hosting code, so setting window.X doesn't create a global variable X.
    //     // Attempt to fix this by setting global variables directly
    //     let transformObjectsToFromGlobal = ["WeakSet", "WeakMap", "Set", "Map", "Symbol", "Promise", "regeneratorRuntime"].map((name) => {
    //         return `
    //         if(typeof(${name}) == "undefined") {
    //             if(window.${name} != undefined) {
    //                 ${name} = window.${name};
    //             }
    //         }
    //         if(typeof(${name}) != "undefined") {
    //             window.${name} = ${name};
    //         }
    //     `
    //     }).join("\n");
    //     resources.libraries += fs.readFileSync("node_modules/babel-polyfill/dist/polyfill.js", "utf-8") + transformObjectsToFromGlobal;
    //     // Moved to builder
    //     // resources.container = makePowerBICompatible(resources.container);
    //     resources.visual = makePowerBICompatible(resources.visual);
    // }
    await fs.writeFile("./dist/resources.json", JSON.stringify(resources), "utf-8");
  }
};

/** Run the specified commands names in sequence */
async function runCommands(sequence) {
  for (const cmd of sequence) {
    console.log("Build: " + cmd);
    await multirun.run(COMMANDS[cmd]);
  }
}

// Execute the specified commands, with no args, run the default sequence
if (sequence.length == 0) {
  sequence = devSequence;
}
runCommands(sequence).catch((e) => {
  console.log(e.message);
  process.exit(-1);
});