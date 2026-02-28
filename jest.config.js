/** @type {import('jest').Config} */
export default {
  preset: 'ts-jest/presets/default-esm',
  testEnvironment: 'node',
  testMatch: [
    "<rootDir>/src/__tests__/**/*.spec.ts"
  ],
  moduleNameMapper: {
    "^(\\.\\.?\\/.+)\\.js$": "$1"
  },
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  transform: {
    '^.+\\.(ts|tsx|js)$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json'
    }]
  },
  collectCoverageFrom: [
    "<rootDir>/src/**/*.ts",
    "!<rootDir>/src/__tests__/**/*.*"
  ],
  coverageReporters: [
    "text",
    "cobertura"
  ]
};