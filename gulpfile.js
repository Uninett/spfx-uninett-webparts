'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');
const merge = require('webpack-merge');
const TerserPlugin = require('terser-webpack-plugin-legacy');
build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.configureWebpack.setConfig({
    additionalConfiguration: function (config) {
      let newConfig = config;
      config.plugins.forEach((plugin, i) => {
        if (plugin.options && plugin.options.mangle) {
          config.plugins.splice(i, 1);
          newConfig = merge(config, {
            plugins: [
              new TerserPlugin()
            ]
          });
        }
      });
  
      return newConfig;
    }
  });

build.initialize(gulp);
