// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

let fs = require("fs");
let babel = require("babel-core");

let enableBabel = false;

let resources = {
    icon: "data:image/png;base64," + fs.readFileSync("assets/icon.png", "base64"),
    libraries: "",
    visual: fs.readFileSync("template_code/visual.js", "utf-8")
};

if (enableBabel) {

    function makePowerBICompatible(code) {
        return babel.transform(code, {
            presets: ["es2015-ie"]
        }).code;
    }

    let transformObjectsToFromGlobal = ["WeakSet", "WeakMap", "Set", "Map", "Symbol", "Promise", "regeneratorRuntime"].map((name) => {
        return `
        if(typeof(${name}) == "undefined") {
            if(window.${name} != undefined) {
                ${name} = window.${name};
            }
        }
        if(typeof(${name}) != "undefined") {
            window.${name} = ${name};
        }
    `
    }).join("\n");

    resources.libraries += fs.readFileSync("node_modules/babel-polyfill/dist/polyfill.js", "utf-8") + transformObjectsToFromGlobal;
    // Moved to builder
    // resources.container = makePowerBICompatible(resources.container);
    resources.visual = makePowerBICompatible(resources.visual);
}

fs.writeFileSync("./src/resources.json", JSON.stringify(resources), "utf-8");