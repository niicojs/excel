# AGENTS.md

Instructions for AI agents working on this codebase.

## Project Overview

`@niicojs/excel` is a TypeScript library for Excel operations. Zero-config setup using modern, high-performance tooling.

## Tech Stack

- **Runtime**: Node.js >= 20
- **Package Manager**: Bun
- **Language**: TypeScript 5.7+ (strict mode)
- **Bundler**: Bunchee (via tsdx)
- **Testing**: Vitest
- **Linting**: Oxlint (Rust-powered, 50-100x faster than ESLint)
- **Formatting**: Oxfmt (Rust-powered, 35x faster than Prettier)

## Project Structure

```
src/
  index.ts          # Library entry point - ALL public APIs exported here
test/
  *.test.ts         # Test files using Vitest
dist/               # Build output (generated, do not edit)
.github/workflows/  # CI configuration
```

## Commands

### Build & Development

```bash
bun install          # Install dependencies
bun run dev          # Start development mode with watch
bun run build        # Build for production (outputs ESM, CJS, and .d.ts)
```

### Testing

```bash
bun run test                        # Run all tests
bun run test:watch                  # Run tests in watch mode
bun run test -- test/index.test.ts  # Run a specific test file
bun run test -- -t "test name"      # Run tests matching a pattern
```

### Code Quality

```bash
bun run lint         # Lint code with Oxlint
bun run format       # Format code with Oxfmt
bun run format:check # Check formatting without modifying
bun run typecheck    # Run TypeScript type checking
```

### Before Committing

Always run these checks before committing:

```bash
bun run typecheck && bun run lint && bun run test
```

## Code Style Guidelines

### Formatting (enforced by Oxfmt)

- **Line width**: 120 characters max
- **Quotes**: Single quotes for strings
- **Semicolons**: Required
- **Indentation**: 2 spaces

### Imports

- Use ES module imports (`import`/`export`)
- Group imports in order: external packages, then local modules
- Use named imports over namespace imports when possible

```typescript
// Good
import { something } from 'external-package';
import { localThing } from './local';

// Avoid
import * as pkg from 'external-package';
```

### Exports

- Use named exports, avoid default exports
- Export all public APIs from `src/index.ts`
- Keep the public API surface minimal

```typescript
// Good
export const myFunction = () => {};
export type MyType = { ... };

// Avoid
export default myFunction;
```

### Types

- Use TypeScript strict mode (enabled in tsconfig.json)
- Prefer `interface` for object shapes, `type` for unions/intersections
- Always type function parameters and return values for public APIs
- Use JSDoc comments for public API documentation

```typescript
/**
 * Adds two numbers together.
 * @param a - First number
 * @param b - Second number
 * @returns The sum of a and b
 */
export const sum = (a: number, b: number): number => {
  return a + b;
};
```

### Naming Conventions

- **Functions/variables**: camelCase (`myFunction`, `myVariable`)
- **Types/Interfaces**: PascalCase (`MyInterface`, `MyType`)
- **Constants**: camelCase or UPPER_SNAKE_CASE for true constants
- **Files**: kebab-case for multi-word files (`my-module.ts`)
- **Test files**: `*.test.ts` pattern

### Error Handling

- Use typed errors when possible
- Throw descriptive error messages
- Document error conditions in JSDoc comments

```typescript
/**
 * Parses an Excel cell reference.
 * @throws {Error} If the reference format is invalid
 */
export const parseReference = (ref: string): CellRef => {
  if (!isValidRef(ref)) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }
  // ...
};
```

### TypeScript Compiler Rules (from tsconfig.json)

These are enforced by the compiler - do not disable them:

- `noUnusedLocals`: No unused local variables
- `noUnusedParameters`: No unused function parameters
- `noImplicitReturns`: All code paths must return a value
- `noFallthroughCasesInSwitch`: No fallthrough in switch statements
- `strict`: All strict type-checking options enabled

## Testing Guidelines

- Write tests in the `test/` directory
- Use Vitest's `describe`, `it`, and `expect`
- Test file naming: `*.test.ts`
- Aim for comprehensive coverage of public APIs

```typescript
import { describe, it, expect } from 'vitest';
import { myFunction } from '../src';

describe('myFunction', () => {
  it('handles normal input', () => {
    expect(myFunction('input')).toBe('expected');
  });

  it('handles edge cases', () => {
    expect(myFunction('')).toBe('');
  });

  it('throws on invalid input', () => {
    expect(() => myFunction(null)).toThrow();
  });
});
```

## Adding New Features

1. Implement the feature in `src/`
2. Export public APIs from `src/index.ts`
3. Add comprehensive tests in `test/`
4. Add JSDoc comments for public APIs
5. Run `bun run typecheck && bun run lint && bun run test`
6. Run `bun run build` to verify the build works

## Module Outputs

The library outputs dual formats with full TypeScript support:

- `dist/index.js` - ESM (import)
- `dist/index.cjs` - CommonJS (require)
- `dist/index.d.ts` - TypeScript declarations

## CI Pipeline

The CI runs on every push/PR to main/master:

1. Lint (`bun run lint`)
2. Type check (`bun run typecheck`)
3. Test (`bun run test`)
4. Build (`bun run build`)

Tested on Node 20 and 22 across Ubuntu, Windows, and macOS.

## Important Notes

- Always run `bun run build` before publishing
- Keep the public API surface minimal and well-documented
- Maintain backward compatibility when possible
- Do not modify files in `dist/` - they are generated
