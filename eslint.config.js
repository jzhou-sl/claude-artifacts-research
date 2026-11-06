const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/default');

module.exports = [
  { ignores: ['src/generated/**', 'src/webparts/**'] },
  ...spfxProfile,
  {
    files: ['**/*.ts', '**/*.tsx'],
    languageOptions: {
      parserOptions: {
        tsconfigRootDir: __dirname,
        project: './tsconfig.json'
      }
    }
  }
];
