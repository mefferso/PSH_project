# Testing strategy

## Layers

### 1. Pure deterministic tests — required on every push

These tests have no network and no Google dependency. They execute production rainfall-core functions against committed Francine source fixtures.

Command:

```bash
npm run test:rainfall
```

### 2. Build and adapter guardrails — required on every push

These confirm the production Apps Script adapter actually delegates rainfall matching/parsing/source precedence to the tested core and that the generated deployable file is current and valid JavaScript.

Command:

```bash
npm run test:smoke
```

### 3. Workbook fixture checks — required on every push

These preserve the exact Francine/reference XLSX files and WeatherSTEM comparisons.

Command:

```bash
npm run fixtures:inspect
```

### 4. Live API integration — scheduled/manual

These catch upstream API schema/data-access changes. They are intentionally separate from deterministic CI so an external outage does not make a correct commit nondeterministically fail.

Command:

```bash
npm run test:live
```

## Regression rule

Every production bug that can be represented without Google runtime state should get a deterministic fixture before or with its fix. Historical reference values that current authoritative sources cannot reproduce should be documented as manual exceptions rather than simulated as automatable. A bug is not considered fixed merely because a manual bound-Sheet run happened to pass.
