const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
require("dotenv").config();

module.exports = async (env, options) => {
    return {
        devtool: "source-map",
        entry: {
            polyfill: "@babel/polyfill",
            taskpane: "./taskpane.ts",
            commands: "./commands.ts",
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"],
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: "babel-loader",
                },
                {
                    test: /\.s[ac]ss$/i,
                    use: [
                        // Creates `style` nodes from JS strings
                        "style-loader",
                        // Translates CSS into CommonJS
                        "css-loader",
                        // Compiles Sass to CSS
                        "sass-loader",
                    ],
                },
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: "ts-loader",
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader",
                },
                {
                    test: /\.(png|jpg|jpeg|gif)$/,
                    loader: "file-loader",
                    options: {
                        name: "[path][name].[ext]",
                    },
                },
            ],
        },
        plugins: [
            new CleanWebpackPlugin(),
            new HtmlWebpackPlugin({
                filename: "commands.html",
                template: "./commands.html",
                chunks: ["polyfill", "commands"],
            }),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./taskpane.html",
                chunks: ["polyfill", "taskpane"],
            }),
            new CopyWebpackPlugin({
                patterns: [
                    {
                        to: "[name][ext]",
                        from: "manifest.xml",
                    },
                    {
                        from: "./assets",
                        to: "assets",
                        force: true,
                    },
                ],
            }),
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*",
            },
            https:
                options.https !== undefined
                    ? options.https
                    : await devCerts.getHttpsServerOptions().then((config) => {
                          // Unsuported key.
                          delete config.ca;
                          return config;
                      }),
            port: process.env.npm_package_config_dev_server_port || 3000,
        },
    };
};
