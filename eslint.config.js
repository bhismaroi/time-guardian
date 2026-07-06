import js from "@eslint/js";
import globals from "globals";
import reactHooks from "eslint-plugin-react-hooks";
import reactRefresh from "eslint-plugin-react-refresh";
import tseslint from "typescript-eslint";

export default tseslint.config(
  { ignores: ["dist"] },
  {
    extends: [js.configs.recommended, ...tseslint.configs.recommended],
    files: ["**/*.{ts,tsx}"],
    languageOptions: {
      ecmaVersion: 2020,
      globals: globals.browser,
    },
    plugins: {
      "react-hooks": reactHooks,
      "react-refresh": reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      "react-refresh/only-export-components": ["warn", { allowConstantExport: true }],
      // Phase 5.3d: re-enable the lint rule that catches unused
      // variables/functions. The TS compiler's noUnusedLocals /
      // noUnusedParameters flags already catch most of the same
      // issues at compile time (Phase 5.3c). The lint rule is the
      // JS-level analog and picks up cases the TS compiler doesn't
      // (e.g. type-only re-exports, function-expression style). Set
      // to "warn" so it doesn't break the build for the harness
      // class-field issues deferred in Phase 5.3b.
      "@typescript-eslint/no-unused-vars": "warn",
    },
  },
);
