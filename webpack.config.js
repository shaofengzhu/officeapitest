const path = require('path');
const webpack = require('webpack');

module.exports = {
    target: 'node',
    context: path.resolve(__dirname, './src'),
    entry: {
        app: './index.js',
    },
    output: {
        path: path.resolve(__dirname, './dist'),
        filename: '[name].bundle.js'
    }
};
