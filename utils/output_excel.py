import os

import pandas as pd


def output_as(data, parameters):
    """
    Outputs the data to an Excel file with the specified parameters.

    :param data: Data to be written to the Excel file.
    :param paramaters: Dictionary containing parameters for output.
    """

    file_name = parameters.get("file_name", "Aggregated_data.xlsx")
    output_path = parameters.get(
        "output_path", parameters["folder_path"] + "/" + "Output"
    )

    # Ensure the directory exists
    os.makedirs(output_path, exist_ok=True)
    print(f"Data has been written to {output_path}")
    # Write the data to an Excel file
    with pd.ExcelWriter(output_path + "/" + file_name, engine="openpyxl") as writer:
        # data.to_excel(writer, index=False)

        for sheet, sheet_data in data.items():
            safe_sheet = sheet.replace(" ", "_")
            print(f"Writing sheet: {safe_sheet}")
            if isinstance(sheet_data, pd.DataFrame):
                sheet_data.to_excel(writer, sheet_name=safe_sheet, index=False)
            else:
                # If it's not a DataFrame, convert to DataFrame first
                pd.DataFrame(sheet_data).to_excel(
                    writer, sheet_name=safe_sheet, index=False
                )

        # df.to_csv(outdir / f"{safe_sheet}.csv", index=False, encoding="utf-8-sig")
