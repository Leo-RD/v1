// Configuration ESLint (format "flat config") : applique le profil React officiel SPFx
// aux fichiers TypeScript du projet, avec les informations de type issues de tsconfig.json.

const spfxProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');

module.exports = [
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
