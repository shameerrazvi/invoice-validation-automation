"""Regenerate the sample batch (CSV) and the Excel validation tool that holds it."""
from src import generate_batch, workbook


def main():
    print("Generating sample batch")
    data = generate_batch.main()
    path = workbook.build_tool_workbook(data)
    print(f"Excel tool written to {path}")


if __name__ == "__main__":
    main()
