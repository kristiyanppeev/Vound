import pathlib
from agent import run

TABLES_DIR = pathlib.Path("tables")


def find_tables(prefix: str) -> list[pathlib.Path]:
    prefix = prefix.strip().lower()
    return sorted(
        p for p in TABLES_DIR.glob("*.xlsx")
        if p.name.lower().startswith(prefix)
    )


def pick_table() -> pathlib.Path:
    while True:
        query = input("Enter the first few letters of the table filename: ").strip()
        if not query:
            print("  Please enter at least one character.\n")
            continue

        matches = find_tables(query)

        if not matches:
            print(f"  No table found starting with '{query}'. Try again.\n")
            continue

        if len(matches) == 1:
            print(f"  Found: {matches[0].name}")
            return matches[0]

        print(f"  Found {len(matches)} matches:")
        for i, p in enumerate(matches, 1):
            print(f"    [{i}] {p.name}")
        while True:
            choice = input("  Enter number to select: ").strip()
            if choice.isdigit() and 1 <= int(choice) <= len(matches):
                return matches[int(choice) - 1]
            print(f"  Invalid choice, enter a number between 1 and {len(matches)}.")


def main():
    table_path = pick_table()
    print()
    result, validation = run(str(table_path))

    print("\n-- Extracted cells --")
    for sheet_name, cells in result.items():
        print(f"\n[{sheet_name}]")
        for c in cells:
            print(f"  {c['cell']:6s}  {c['type']:10s}  {str(c.get('current_value', '')):15s}  {c['description']}")

    print("\n-- Validation results --")
    for sheet_name, v in validation.items():
        if v.accepted:
            print(f"  [{sheet_name}] accepted")
        else:
            print(f"  [{sheet_name}] REJECTED — {len(v.cells_to_fix)} cells flagged")
            for fix in v.cells_to_fix:
                print(f"    {fix.cell}: {fix.reason}")


if __name__ == "__main__":
    main()
