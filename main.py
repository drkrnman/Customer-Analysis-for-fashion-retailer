from utils import (
    read_summary,
    ltv_factors,
    ltv_cohort,
    revenue_structure,
    statistical_tests
)

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
    while True:
        try:
            print_menu()
            choice = input("\nEnter the number in the list (or 'exit'): ").strip()

            if choice.lower() == 'exit' or choice == '0':
                print("Finished, bye!")
                break

            choice = int(choice)

            match choice:
                case 1:
                    read_summary()
                case 2:
                    ltv_factors()
                case 3:
                    ltv_cohort()
                case 4:
                    revenue_structure()
                case 5:
                    statistical_tests()
                case _:
                    print("Not an option!")
        except ValueError:
            print("Please enter a number or 'exit'.")
        except Exception as e:
            print(f"Unexpected error: {e}")

if __name__ == "__main__":
    main()
