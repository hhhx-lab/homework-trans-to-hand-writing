const { defineConfig } = require("@vue/cli-service");
const path = require("path");

module.exports = defineConfig({
  transpileDependencies: true,
  devServer: {
    host: "localhost", // 或者使用 '0.0.0.0'
    port: 8080,
    proxy: {
      "/api": {
        target: "http://127.0.0.1:5005",
        changeOrigin: true,
      },
    },
  },
  pages: {
    index: {
      // entry for the page
      entry: "src/main.js",
      // the source template
      template: "public/index.html",
      // output as dist/index.html
      filename: "index.html",
      // when using title option,
      // template title tag needs to be <title><%= htmlWebpackPlugin.options.title %></title>
      title: "作业文档转手写体工作台",
    },
  },
  filenameHashing: process.env.NODE_ENV === "production",
  configureWebpack: {
    performance: {
      maxAssetSize: 768 * 1024,
      maxEntrypointSize: 768 * 1024,
    },
    optimization: {
      splitChunks: {
        chunks: "all",
        cacheGroups: {
          vue: {
            test: /[\\/]node_modules[\\/](vue|vue-router|vuex|@vueuse)[\\/]/,
            name: "vendor-vue",
            chunks: "all",
            priority: 40,
          },
          bootstrap: {
            test: /[\\/]node_modules[\\/](bootstrap|@popperjs)[\\/]/,
            name: "vendor-bootstrap",
            chunks: "all",
            priority: 30,
          },
          http: {
            test: /[\\/]node_modules[\\/](axios|axios-retry)[\\/]/,
            name: "vendor-http",
            chunks: "all",
            priority: 20,
          },
          sweetalert: {
            test: /[\\/]node_modules[\\/]sweetalert2[\\/]/,
            name: "vendor-sweetalert",
            chunks: "all",
            priority: 20,
          },
        },
      },
    },
  },
  chainWebpack: (config) => {
    config.plugin("copy").tap((args) => {
      const options = args[0] || {};
      const patterns = options.patterns || [];
      patterns.forEach((pattern) => {
        if (pattern.from && path.basename(pattern.from) === "public") {
          pattern.globOptions = pattern.globOptions || {};
          pattern.globOptions.ignore = [
            ...(pattern.globOptions.ignore || []),
            "**/default.png",
            "**/default1.png",
            "**/writing.png",
            "**/favicon.svg",
          ];
        }
      });
      return args;
    });
  },
  pwa: {
    name: "作业文档转手写体工作台",
    themeColor: "#4fc08d",
    msTileColor: "#000000",
    iconPaths: {
      faviconSVG: "icon.svg",
      favicon32: "favicon-96x96.png",
      favicon16: "favicon-96x96.png",
      appleTouchIcon: "apple-touch-icon.png",
      maskIcon: "icon.svg",
      msTileImage: "web-app-manifest-192x192.png",
    },
    workboxPluginMode: "GenerateSW",
    workboxOptions: {
      skipWaiting: true,
      clientsClaim: true,
      exclude: [/default\.png$/, /default1\.png$/, /writing\.png$/, /favicon\.svg$/],
      runtimeCaching: [
        // StaleWhileRevalidate - 返回缓存同时更新（适用于静态资源）
        {
          urlPattern: /\.(js|css|woff2?)$/,
          handler: "StaleWhileRevalidate",
          options: {
            cacheName: "static-resources",
            cacheableResponse: { statuses: [0, 200] },
          },
        },
        // CacheFirst - 缓存优先（适用于图片等不常变更的资源）
        {
          urlPattern: /\.(png|jpg|jpeg|gif|svg|ico|webp)$/,
          handler: "CacheFirst",
          options: {
            cacheName: "image-cache",
            expiration: {
              maxEntries: 100,
              maxAgeSeconds: 30 * 24 * 60 * 60, // 30 天
            },
            cacheableResponse: { statuses: [0, 200] },
          },
        },
      ],
    },
  },
});
