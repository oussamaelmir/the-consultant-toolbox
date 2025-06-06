const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:60367/";
const urlProd = "https://the-consultant-toolbox.azurewebsites.net";

module.exports = {
  mode: "production",
  entry: {
    polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
    taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
    commands: "./src/commands/commands.ts",
    flags: "./src/flags/flags.ts",
  },
  output: {
    filename: "[name].js",
    path: path.resolve(__dirname, "dist"),
    clean: true,
    publicPath: "", // Important for relative paths
  },
  resolve: {
    extensions: [".ts", ".html", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        exclude: /node_modules/,
        use: "babel-loader",
      },
      {
        test: /\.html$/,
        use: "html-loader",
      },
      {
        test: /\.(png|jpg|jpeg|gif|ico)$/,
        type: "asset/resource",
        generator: {
          filename: "assets/[name][ext]",
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
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["polyfill", "commands"],
    }),
    new HtmlWebpackPlugin({
      filename: "flags.html",
      template: "./src/flags/flags.html",
      chunks: ["polyfill", "flags"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        // ✅ Static images and assets
        { from: "assets/*", to: "assets/[name][ext]" },

        // ✅ Standalone pages under folders
        { from: "src/support/index.html", to: "support/index.html" },
        { from: "src/privacy/index.html", to: "privacy/index.html" },

        // ✅ Static web.config (optional on Linux, required on Windows only)
        { from: "web.config", to: "web.config" },

        // ✅ Manifest for Office Add-in, with environment URL replacement
        {
          from: "manifest*.xml",
          to: "[name][ext]",
          transform(content) {
            return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
          },
        },
      ],
    }),
  ],
};
