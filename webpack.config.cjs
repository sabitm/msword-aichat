const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = async (_env, argv) => {
  const isProd = argv.mode === "production";

  let https = true;
  try {
    const devCerts = require("office-addin-dev-certs");
    https = await devCerts.getHttpsServerOptions();
    console.log("Using office-addin-dev-certs for webpack-dev-server HTTPS.");
  } catch (error) {
    console.warn("office-addin-dev-certs unavailable:", error.message);
  }

  return {
    mode: isProd ? "production" : "development",
    devtool: isProd ? false : "source-map",
    entry: {
      taskpane: ["./src/polyfills.ts", "./src/taskpane/main.legacy.tsx"],
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: isProd ? "[name].[contenthash:8].bundle.js" : "[name].bundle.js",
      clean: true,
      publicPath: "/",
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: "babel-loader",
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "taskpane.template.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "commands.html", to: "commands.html" },
          { from: "public/assets", to: "assets" },
        ],
      }),
    ],
    devServer: {
      port: 3000,
      host: "localhost",
      server: {
        type: "https",
        options: typeof https === "object" ? https : undefined,
      },
      static: {
        directory: path.join(__dirname, "dist"),
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      hot: true,
      historyApiFallback: false,
    },
    performance: {
      hints: isProd ? "warning" : false,
      maxEntrypointSize: 3_000_000,
      maxAssetSize: 3_000_000,
    },
  };
};