# Assign OA by Province (VBA Macro)

This VBA macro automatically assigns OA (Operation Assistants) to customer records based on their province, avoiding assignment duplication from the past 4 months and following specified percentage distribution rules per province.

## 📌 Purpose

- Ensure fair and balanced OA distribution per province.
- Prevent re-assigning the same OA within the last 4 months.
- Use percentage weights from `OA_Master` to guide assignment ratios.

## 📂 Sheets Used

- **Sheet1**:
  - `Column R`: OA to be assigned (output)
  - `Column AI`: Province
  - `Columns S, T, U, V`: OA assignment history for May, Apr, Mar, Feb

- **OA_Master**:
  - `Column A`: Province
  - `Column B`: OA name
  - `Column C`: Assigned percentage

## ⚙️ How It Works

1. Loops through each row in `Sheet1`.
2. Looks up the province from column `AI`.
3. Gathers all OAs for that province from `OA_Master` along with their percentage weights.
4. Filters out OAs assigned in the last 4 months.
5. Randomly assigns one OA from the eligible list, based on percentage distribution.
6. If all OAs are recently used, it falls back to the oldest one from past 4 months.
7. Writes result to column `R`.

## 💻 How to Use

1. Open your Excel file.
2. Press `Alt + F11` to open the VBA Editor.
3. Insert a new Module and paste the macro.
4. Run `AssignOA_ByPercentage_AvoidDuplicates`.
5. Result will appear in column `R` in `Sheet1`.

## 📝 Notes

- Make sure both `Sheet1` and `OA_Master` are named exactly as above.
- Percentages in `OA_Master` must total 100% per province (but partial totals are okay too).
- This macro does not use formulas — final OA results are static values.

---

✅ This tool helps automate OA assignment fairly and reproducibly across provinces.
