/**
 * Custom webpack config for InfoCard PCF control.
 *
 * In production builds (PcfBuildMode=production), Terser strips console.log
 * and console.debug calls. console.warn and console.error are preserved.
 *
 * Dev builds (default for pac pcf push) keep all logging.
 */
const TerserPlugin = require("terser-webpack-plugin");

module.exports = {
    optimization: {
        minimizer: [
            new TerserPlugin({
                terserOptions: {
                    compress: {
                        pure_funcs: ["console.log", "console.debug"],
                    },
                },
            }),
        ],
    },
};
