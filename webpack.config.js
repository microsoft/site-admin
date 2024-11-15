var path = require("path");
var project = require("./package.json");
const MomentLocalesPlugin = require('moment-locales-webpack-plugin');

// Export the configuration
module.exports = (env, argv) => {
    var isDev = argv.mode === "development";

    // Return the configuration
    return {
        // Main project files
        entry: [
            path.resolve(__dirname, project.main)
        ],

        // Output information
        output: {
            path: path.resolve(__dirname, "dist"),
            filename: project.name + (isDev ? "" : ".min") + ".js",
            publicPath: ""
        },

        // Keep only 'en' locales with Moment.js
        plugins: [
            new MomentLocalesPlugin(),
        ],

        // Resolve the file names
        resolve: {
            extensions: [".js", ".css", ".scss", ".ts"]
        },

        // Loaders
        module: {
            rules: [
                // SASS to JavaScript
                {
                    // Target the sass and css files
                    test: /\.s?css$/,
                    // Define the compiler to use
                    use: [
                        // Create style nodes from the CommonJS code
                        { loader: "style-loader" },
                        // Translate css to CommonJS
                        { loader: "css-loader" },
                        // Compile sass to css
                        {
                            loader: "sass-loader",
                            options: {
                                implementation: require("sass")
                            }
                        }
                    ]
                },
                // Handle Image Files
                {
                    test: /\.(jpe?g|png|gif|svg|eot|woff|ttf)$/,
                    loader: "url-loader"
                },
                // TypeScript to JavaScript
                {
                    // Target TypeScript files
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: [
                        // JavaScript (ES5) -> JavaScript (Current)
                        {
                            loader: "babel-loader",
                            options: { presets: ["@babel/preset-env"] }
                        },
                        // TypeScript -> JavaScript (ES5)
                        { loader: "ts-loader" }
                    ]
                }
            ]
        }
    };
}