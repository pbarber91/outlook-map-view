/* eslint-disable no-undef */

require("dotenv").config();

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const DEV_BASE_URL = "https://localhost:3000";
const PROD_BASE_URL = (process.env.APP_BASE_URL || "").replace(/\/+$/, "");
const MAPBOX_ACCESS_TOKEN = process.env.MAPBOX_ACCESS_TOKEN || "";
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID || "";
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID || "";

function getBaseUrl(isDev) {
  if (isDev) {
    return DEV_BASE_URL;
  }

  if (PROD_BASE_URL) {
    return PROD_BASE_URL;
  }

  return "https://outlook-map-view.vercel.app";
}

function getValidDomain(baseUrl) {
  try {
    return new URL(baseUrl).host;
  } catch {
    return "localhost";
  }
}

function transformManifest(content, isDev) {
  const baseUrl = getBaseUrl(isDev);
  const validDomain = getValidDomain(baseUrl);

  return content
    .toString()
    .replace(/__APP_BASE_URL__/g, baseUrl)
    .replace(/__APP_VALID_DOMAIN__/g, validDomain);
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return {
    ca: httpsOptions.ca,
    key: httpsOptions.key,
    cert: httpsOptions.cert,
  };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const baseUrl = getBaseUrl(dev);

  return {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
      commands: "./src/commands/commands.ts",
    },
    output: {
      clean: true,
      publicPath: `${baseUrl}/`,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
      fallback: {
        process: require.resolve("process/browser"),
      },
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
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
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "./src/taskpane/popup-complete.html",
            to: "popup-complete.html",
          },
          {
            from: "manifest*.json",
            to: "[name][ext]",
            transform(content) {
              return transformManifest(content, dev);
            },
          },
        ],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
        process: "process/browser",
      }),
      new webpack.DefinePlugin({
        __MAPBOX_ACCESS_TOKEN__: JSON.stringify(MAPBOX_ACCESS_TOKEN),
        __AZURE_CLIENT_ID__: JSON.stringify(AZURE_CLIENT_ID),
        __AZURE_TENANT_ID__: JSON.stringify(AZURE_TENANT_ID),
      }),
    ],
    devServer: {
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };
};