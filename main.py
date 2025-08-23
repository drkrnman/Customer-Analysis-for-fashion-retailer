from utils import (
    read_summary,
    ltv_factors,
    ltv_cohort,
    revenue_structure,
    statistical_tests
)

# Legacy CLI kept for reference. The app now launches the GUI from gui_app.py

def print_menu():
    print("\n\t\t Global Fashion Retail Sales:")
    print("\nMain menu")
    print()
    print("1. Read executive summary")
    print("2. LTV factors")
    print("3. LTV cohort")
    print("4. Revenue structure")
    print("5. Statistical test")
    print("0. Exit")


def main():
    from gui_app import run_gui
    run_gui()


if __name__ == "__main__":
    main()
