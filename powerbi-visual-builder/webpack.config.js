module.exports = {
    entry: {
        powerbi_visual_builder: "./dist/builder.js"
    },
    output: {
        filename: "[name].js",
        path: __dirname + "/dist",
        // Export the app as a global variable "CharticulatorPowerBIVisualBuilder"
        libraryTarget: "umd",
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