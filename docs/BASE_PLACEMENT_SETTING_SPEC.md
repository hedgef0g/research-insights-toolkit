### Proposed Base Placement Setting Specification (docs/BASE_PLACEMENT_SETTING.md)

#### Purpose

Define how the RIT add‑in detects and uses the **Base** row relative to metric rows.  Real‑world tables may place the base either *below* or *above* the metrics; currently RIT assumes the base is below.  This setting introduces three modes—**Auto**, **Base below**, and **Base above**—and describes their impact on each metric type.

#### User-facing setting

Add a dropdown labelled **“Расположение базы”** (“Base placement”) to the Task Pane, alongside confidence-level and test-tail settings (not in a separate “Advanced settings” block).  Options:

| Option (RU)     | Option (EN)    | Behaviour & default                                                                                                                                                                                                                              |
| --------------- | -------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| **Авто**        | **Auto**       | **Default.**  Attempts to find a Base row **below** the metrics.  If not found, looks **above**.  If both exist, uses the one closest to the metrics (preferring below).  Missing base produces a warning and skips significance for that block. |
| **База снизу**  | **Base below** | Forces detection of the Base **below** metric rows.  If the Base appears above, a warning is shown and the block is not processed.                                                                                                               |
| **База сверху** | **Base above** | Forces detection of the Base **above** metric rows.  If the Base appears below, a warning is shown and the block is not processed.                                                                                                               |

#### Behaviour by metric type

1. **Proportion rows** (`%`, “Percent”, `Доля` etc.):

   * *Base below:* Current behaviour: each proportion row uses the base row immediately below.
   * *Base above:* Each proportion row uses the base row immediately above.
   * *Auto:* Searches below first; if none, searches above.

2. **Mean rows + SD/Variance**:

   * *Base below:* Block structure is `Mean` → `SD`/`Variance` → `Base`.
   * *Base above:* Block structure is `Base` → `Mean` → `SD`/`Variance`.
   * Auto chooses as for proportions.

3. **NPS-first / NPS-first with Neutral** (`NPS` / `Promoters` / `Detractors` / `Base`):

   * *Base below:* Default RIT behaviour (base after NPS rows).
   * *Base above:* Block structure becomes `Base` → `NPS` → `Promoters` → `Detractors` (and `Neutral` if present).
   * Auto: same search logic.

4. **Extended NPS** (scale rows → `NPS` → `Base`):

   * Base below: current detection.
   * Base above: structure becomes `Base` → scale rows → `NPS`.
   * Auto: same search logic.

5. **Shared base across blocks**:  When a Base row is missing for a given block, RIT searches for the first available Base (closest in the chosen direction) and uses it as a **shared base** for preceding metric rows.  For example, if a proportion block lacks its own Base but the following mean block has a Base row, the mean block’s Base will be used for the proportions (in Base‑below mode).  The same logic applies in Base‑above mode (searching upward).  This extends existing “Shared Base” behaviour but honours the placement mode.

6. **Multiple metrics, mixed base positions**:  If multiple blocks exist in a selection and some require Base‑above while others require Base‑below, RIT warns the user and does not attempt to reconcile them; manual selection or separate runs are advised.

#### UI / UX notes

* The dropdown should be grouped with other statistical settings, not hidden deeply.  A tooltip should explain each option.
* Warnings will appear in the status message area of the Task Pane when a block fails to match the chosen mode or a base cannot be found.

#### Risks and edge cases

* **Mixed base positions** within a single selection can confuse Auto; user should select consistent blocks.
* **Blocks with no Base** will still be skipped, even in Auto, to preserve statistical validity.
* **Performance impact** is minimal; detection just adds a one‑row search above the block.

#### Follow‑up implementation tasks

1. **Update detection logic** in `metric-detector.js` and `block-detector.js` to honour the new setting.
2. **Update the Task Pane UI** to add the base‑placement dropdown with default Auto.
3. **Add new GST test cases** documenting base‑above and shared‑base scenarios.
4. **Update the Table Structure Matrix**: once implemented, promote base‑placement rows from “FUTURE” to “SUPPORTED” or “PARTIAL”.
5. **Ensure existing selection logic** remains backwards‑compatible under “Base below”.
