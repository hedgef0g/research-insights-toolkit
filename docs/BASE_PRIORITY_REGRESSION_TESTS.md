# Base Priority Regression Tests (Issue #69)

**Purpose:** Manual regression test cases for base subtype priority selection logic.

**Current Status:** No automated test runner exists. These are manual test specifications.

**Base Priority Order:** Effective (0) > Unweighted (1) > plain Base (2) > Weighted (3)

---

## Manual Test Cases

Each test case describes the input table structure, expected behavior, and how to verify it in Excel.

### Test 1: Effective Base + Weighted Base

Input rows:
- Row %: proportion type
- Effective Base: base type, baseSubtype=effective
- Weighted Base: base type, baseSubtype=weighted

Expected: Effective Base is selected for significance testing.

Verify in Excel: Select all three rows. Run add-in. Check that the Effective Base row (not Weighted Base) is used in calculations.

Code location: src/core/metric-detector.js, selectBestFromConsecutiveBases() at line 337

---

### Test 2: Unweighted Base + Weighted Base

Input rows:
- Row %: proportion type
- Unweighted Base: base type, baseSubtype=unweighted
- Weighted Base: base type, baseSubtype=weighted

Expected: Unweighted Base is selected.

Verify in Excel: Select all three rows. Run add-in. Unweighted Base should be the test base.

Code location: src/core/metric-detector.js, selectBestFromConsecutiveBases() at line 337

---

### Test 3: plain Base + Weighted Base

Input rows:
- Row %: proportion type
- Base: base type, baseSubtype=undefined (plain)
- Weighted Base: base type, baseSubtype=weighted

Expected: plain Base is selected.

Verify in Excel: Select all three rows. Run add-in. The plain Base row (not Weighted) should be used.

Code location: src/core/metric-detector.js, getBasePriorityValue() at line 320 (plain base returns priority 2)

---

### Test 4: Weighted Base only

Input rows:
- Row %: proportion type
- Weighted Base: base type, baseSubtype=weighted

Expected: Weighted Base is used as fallback when no other base subtype exists.

Verify in Excel: Select both rows. Run add-in. Weighted Base should be used for calculations.

Code location: src/core/metric-detector.js, findBestBaseRowIndex() at line 357

---

### Test 5: plain Base only

Input rows:
- Row %: proportion type
- Base: base type, baseSubtype=undefined (plain)

Expected: Existing plain Base behavior is unchanged. Block has no baseSubtype attached.

Verify in Excel: This is the legacy single-base case. Run add-in. Behavior should match pre-issue-69 behavior.

Code location: src/core/metric-detector.js, attachBaseSubtype() at line 256 (preserves undefined baseSubtype)

---

### Test 6: NPS-first with consecutive subtype base rows

Input rows:
- Row 1: proportion type
- NPS: nps type
- Promoters: promoters type
- Detractors: detractors type
- Weighted Base: base type, baseSubtype=weighted
- Effective Base: base type, baseSubtype=effective

Expected: NPS-first block uses the same base priority logic. Effective Base (highest priority) is selected for the NPS block.

Verify in Excel: Select all rows. Run add-in. Check that NPS significance markers use Effective Base, not Weighted Base.

Code location: src/core/metric-detector.js, line 441 (NPS-first uses selectBestFromConsecutiveBases)

---

## Code Implementation References

The base priority logic is implemented in src/core/metric-detector.js:

| Function | Lines | Purpose |
|----------|-------|---------|
| getBasePriorityValue() | 320-326 | Assigns numeric priority to each base subtype |
| selectBestFromConsecutiveBases() | 337-351 | Selects base with lowest (best) priority value |
| findBestBaseRowIndex() | 357-361 | Finds first base and selects best from consecutive |
| attachBaseSubtype() | 256-262 | Adds baseSubtype to block if base row has one |
| buildCalculationBlocks() | 370-559 | Main entry point; applies priority selection |

Key behavior:
- Priority values: Effective=0, Unweighted=1, plain=2, Weighted=3
- When multiple consecutive bases exist, the lowest priority value is selected
- If no baseSubtype is present, priority is 2 (plain Base)

---

## Future Automation Approach

To implement automated regression testing:

1. Create a test runner that imports buildCalculationBlocks() from metric-detector.js
2. For each test case, create mock row diagnostics (array of objects with rowType, baseSubtype)
3. Call buildCalculationBlocks() with the mock data
4. Assert that block.baseRowIndex and block.baseSubtype match expected values
5. Add to CI/CD pipeline to run on each PR

Example pattern:
```
Test 1 input:
  [proportion, {base, effective}, {base, weighted}]
  
Call:
  blocks = buildCalculationBlocks(rowDiagnostics)
  
Assert:
  blocks[0].baseRowIndex === 1
  blocks[0].baseSubtype === "effective"
```

This approach requires selecting a test framework (Jest, Mocha, etc.) and setting up test infrastructure. See package.json for current dev dependencies.

---

## Validation Status

- [x] Six required test cases documented
- [x] Priority order verified in code (lines 320-326)
- [x] Code paths identified for each case
- [x] Manual validation steps provided
- [x] Future automation approach sketched
- [x] No runtime code modified
- [x] No forbidden files touched
