module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  roots: ["<rootDir>/InfoCard/tests"],
  testMatch: ["**/*.test.ts", "**/*.test.tsx"],
  moduleFileExtensions: ["ts", "tsx", "js", "jsx", "json"],
  collectCoverage: true,
  coverageDirectory: "coverage",
  coverageReporters: ["text", "lcov"],
  collectCoverageFrom: [
    "InfoCard/**/*.ts",
    "InfoCard/**/*.tsx",
    "!InfoCard/generated/**",
    "!InfoCard/tests/**",
  ],
  moduleNameMapper: {
    "^./generated/ManifestTypes$": "<rootDir>/InfoCard/generated/ManifestTypes",
  },
  globals: {
    "ts-jest": {
      tsconfig: {
        jsx: "react",
        esModuleInterop: true,
        allowJs: true,
        module: "commonjs",
        target: "es6",
        strict: false,
        typeRoots: ["node_modules/@types"],
      },
    },
  },
};
