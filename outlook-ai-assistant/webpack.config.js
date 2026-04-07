/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const MiniCssExtractPlugin = require("mini-css-extract-plugin");

const urlDev = "https://localhost:3000/";

module.exports = (env, options) => {
  const devMode = options.mode === "development";

  return {
    devtool: devMode ? "source-map" : false,
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "ts-loader",
          },
        },
        {
          test: /\.css$/,
          use: [MiniCssExtractPlugin.loader, "css-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
        },
      ],
    },
    plugins: [
      new MiniCssExtractPlugin({ filename: "[name].css" }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/**/*", to: "[path][name][ext]" },
          { from: "manifest.xml", to: "manifest.xml" },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: {
          // Self-signed certificate — run `npx office-addin-dev-certs install` first
          // or provide your own cert/key paths here.
        },
      },
      port: 3000,
      hot: true,
    },
    output: {
      clean: true,
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
      publicPath: devMode ? urlDev : "/",
    },
  };
};
