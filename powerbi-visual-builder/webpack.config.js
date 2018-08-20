module.exports = {
    entry: {
        powerbi_visual_builder: "./src/builder.js"
    },
    output: {
        filename: "[name].js",
        path: __dirname + "/dist",
        // Export the app as a global variable "Charticulator"
        libraryTarget: "var",
        library: "CharticulatorPowerBIVisualBuilder"
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                loader: 'babel-loader',
                query: {
                    presets: ['es2015']
                }
            }
        ]
    },
};