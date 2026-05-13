import { defineConfig } from "vite-plus";

export default defineConfig({
  lint: { options: { typeAware: true, typeCheck: true } },
  fmt: { printWidth: 120, singleQuote: true, sortImports: true, endOfLine: "crlf" },
  test: { globals: true, environment: "node", include: ["test/**/*.test.ts"] },
});
