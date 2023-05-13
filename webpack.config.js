/* eslint-disable no-undef */

const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const nodeExternals = require("webpack-node-externals");

const urlDev = "https://localhost:44316/";
const urlProd = "https://orange-water-0d260d810.3.azurestaticapps.net/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = [
    {
      devtool: "source-map",
      entry: {
        polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
        taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
        mainpane: "./src/taskpane/mainpane.ts",
        commands: "./src/commands/commands.js",
        login: "./src/authentication/login.js",
        fallbackauthdialog: "./src/helpers/fallbackauthdialog.js",
      },
      resolve: {
        extensions: [".html", ".js"],
        fallback: {
          buffer: require.resolve("buffer/"),
          http: require.resolve("stream-http"),
          https: require.resolve("https-browserify"),
          url: require.resolve("url/"),
        },
      },
      module: {
        rules: [
          {
            test: /\.js$/,
            exclude: /node_modules/,
            use: {
              loader: "babel-loader",
              options: {
                presets: ["@babel/preset-env"],
              },
            },
          },
          {
            test: /\.html$/,
            exclude: /node_modules/,
            use: "html-loader",
          },
          {
            test: /\.(png|jpg|jpeg|gif|ico)$/,
            type: "asset/resource",
            generator: {
              filename: "assets/[name][ext][query]",
            },
          },
        ],
      },
      plugins: [
        new HtmlWebpackPlugin({
          filename: "taskpane.html",
          template: "./src/taskpane/taskpane.html",
          chunks: ["polyfill", "taskpane"],
        }),
        new HtmlWebpackPlugin({
          filename: "mainpane.html",
          template: "./src/taskpane/mainpane.html",
          chunks: ["polyfill", "mainpane"],
        }),
        new HtmlWebpackPlugin({
          filename: "commands.html",
          template: "./src/commands/commands.html",
          chunks: ["polyfill", "commands"],
        }),
        new HtmlWebpackPlugin({
          filename: "fallbackauthdialog.html",
          template: "./src/helpers/fallbackauthdialog.html",
          chunks: ["polyfill", "fallbackauthdialog"],
        }),
        new HtmlWebpackPlugin({
          filename: "login.html",
          template: "./src/authentication/login.html",
          chunks: ["polyfill", "login"],
        }),
        new HtmlWebpackPlugin({
          filename: "index.html",
          template: "./src/taskpane/taskpane.html",
          chunks: ["polyfill", "taskpane"],
        }),
        new CopyWebpackPlugin({
          patterns: [
            {
              from: "assets/*",
              to: "assets/[name][ext][query]",
            },
            {
              from: "manifest*.xml",
              to: "[name]" + "[ext]",
              transform(content) {
                if (dev) {
                  return content;
                } else {
                  return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                }
              },
            },
          ],
        }),
      ],
    },
    {
      devtool: "source-map",
      target: "node",
      entry: {
        middletier: "./src/middle-tier/app.js",
      },
      output: {
        clean: true,
      },
      externals: [nodeExternals()],
      resolve: {
        extensions: [".js"],
      },
      module: {
        rules: [
          {
            test: /\.js$/,
            exclude: /node_modules/,
            use: {
              loader: "babel-loader",
              options: {
                presets: ["@babel/preset-env"],
              },
            },
          },
        ],
      },
      plugins: [
        new CopyWebpackPlugin({
          patterns: [
            {
              from: ".env",
              to: ".",
            },
          ],
        }),
      ],
    },
  ];

  return config;
};
