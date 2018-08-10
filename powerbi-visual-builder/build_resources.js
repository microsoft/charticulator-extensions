let fs = require("fs");

let resources = {
    icon: "data:image/png;base64," + fs.readFileSync("assets/icon.png", "base64"),
    libraries: "",
    container: fs.readFileSync("../../charticulator/dist/scripts/container.bundle.min.js", "utf-8"),
    visual: fs.readFileSync("template_code/visual.js", "utf-8")
};

fs.writeFileSync("./src/resources.json", JSON.stringify(resources), "utf-8");
