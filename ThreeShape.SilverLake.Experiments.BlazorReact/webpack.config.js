const path = require('path')
const fs = require("fs");

module.exports = {
    entry: () => fs.readdirSync("./React/").filter(f => f.endsWith(".jsx")).map(f => `./React/${f}`),
    // Where files should be sent once they are bundled
    output: {
        filename: "app.js",
        path: path.resolve(__dirname, "./wwwroot/scripts")
    },
    // Rules of how webpack will take our files, complie & bundle them for the browser 
    module: {
        rules: [
            {
                test: /\.(js|jsx)$/,
                exclude: /nodeModules/,
                use: {
                    loader: 'babel-loader'
                }
            },
            {
                test: /\.s?css$/,
                use: [
                    'style-loader',
                    'css-loader',
                    'sass-loader',
                    {
                        loader: 'sass-resources-loader',
                        options: {
                            resources: [path.resolve(__dirname, "./React/variables.scss")]
                        },
                    },
                ]
            }
        ]
    },
    resolve: {
        extensions: ['.jsx', '.js']
    }
}
