module.exports = {
  presets: [
    [
      "@babel/preset-env",
      {
        useBuiltIns: "entry",
        corejs: 3,
        targets: { ie: "11" },
      },
    ],
    ["@babel/preset-react", { runtime: "classic" }],
    "@babel/preset-typescript",
  ],
};