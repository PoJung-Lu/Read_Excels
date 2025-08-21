import os
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as


parameters = {
    "read_all_sheets": True,  # Set to True if you want to read all sheets in each Excel file
    "sheet_names": "公共危險物品運作資料",  # Specify the sheet name if needed
    "path_data": "../Data/科技廠救災能量",  # Specify the path to your Excel files
    # "path_data": "../Test",  # For testing purposes, you can change this to a test folder
    "pattern": "top_ten_operating_chemicals",  # You can specify a pattern if needed, ['default', '苗栗縣', top_ten_operating_chemicals]
    "path_output": "/Output",  # Specify the output folder
    "output_path": "Output",  # Specify the output path
    "file_name": "Aggregated_data.xlsx",  # Specify the output file name
}


def main():
    # Create an instance of the read_data class
    data_reader = read_data.read_data(parameters=parameters)

    # Access the combined DataFrame
    path = data_reader.get_path()
    combined_data = data_reader.read_excel_files()
    output_as(combined_data, data_reader.parameters)

    # for file, df in combined_data.items():
    #     print(f"Data from {file}:")
    #     print(df.head(5))  # Display the first few rows of each sheet
    #     print(df.keys())


def high_tech_industry_main():
    """Function to handle high-tech industry data processing."""
    data_reader = read_data.read_data(parameters=parameters)
    folder_path = data_reader.get_path()
    data_reader.parameters["output_path"] = (
        folder_path + data_reader.parameters["path_output"]
    )
    # print(f"Folder path: {folder_path}")
    folders = os.listdir(folder_path)
    if "Output" in folders:
        folders.remove("Output")  # Remove the Output folder if it exists
    # for folder in folders:
    #     file_path = os.path.join(folder_path, folder)
    #     data_reader.parameters["folder_path"] = file_path
    #     print(f"Processing folder: {folder}")
    #     data_reader.parameters["file_name"] = folder + ".xlsx"
    #     # Read the Excel file
    #     if data_reader.parameters["pattern"] == "default":
    #         print("Pattern is default")
    #         combined_data = data_reader.read_excel_files()

    #     else:
    #         print(f"Reading with pattern: {data_reader.parameters['pattern']}")
    #         combined_data = data_reader.read_with_pattern(
    #             {}, {}, data_reader.parameters["pattern"]
    #         )
    #     output_as(
    #         combined_data, data_reader.parameters
    #     )  # Output the combined data to an Excel file
    # print("All folders processed successfully.")

    ############    sort required data

    # data_reader.parameters["path_data"] = (
    #     data_reader.parameters["path_data"] + data_reader.parameters["path_output"]
    # )
    # data_reader.parameters["folder_path"] = data_reader.get_path()
    # data_reader.parameters["pattern"] = "sort_by_location"
    # data_reader.parameters["file_name"] = "Sorted_data.xlsx"
    # sorted_data = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}
    # sorted_factory = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}

    # for f, (r, v) in data_reader.read_excel_files():
    #     print(f"Processing file: {f}", r)
    #     sorted_data[r].extend(v)
    #     sorted_factory[r].append(f)
    #     print(sum(i.shape[0] for i in sorted_data[r]))

    # print("Data sorted by location:", sorted_data.keys())
    # print("Factory sorted by location:", sorted_factory)

    # sorted_data = {
    #     k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
    #     for k, v in sorted_data.items()
    # }
    # output_as(sorted_data, data_reader.parameters)

    ############## Analyze data

    data_reader.parameters["path_data"] = (
        (data_reader.parameters["path_data"] + data_reader.parameters["path_output"])
        if not data_reader.parameters["path_output"]
        in data_reader.parameters["path_data"]
        else data_reader.parameters["path_data"]
    )

    data_reader.parameters["folder_path"] = data_reader.get_path()
    data_reader.parameters["read_all_sheets"] = True
    keys, values = data_reader.read_one_excel(
        data_reader.parameters["folder_path"] + "/Sorted_data.xlsx"
    )

    sort_column = ["化學物質名稱", "容器材質", "物質儲存型態"]
    sort_value = ["廠內最大儲存量(公斤)", "廠內最大儲存量(公升)"]
    sorted_data = [{}, {}, {}]  # [sort_by_hazmat, sort_by_container, sort_by_state]
    file_name = ["sort_by_hazmat.xlsx", "sort_by_container.xlsx", "sort_by_state.xlsx"]

    for k, v in zip(keys, values):
        if k == "其他":
            break

        v["化學物質名稱"] = v["化學物質名稱"].astype(str)
        # Convert "廠內最大儲存量(公斤)" to float, using pandas string method, str.extract().
        v["廠內最大儲存量(公斤)"] = (
            v["廠內最大儲存量(公斤)"]
            .astype(str)
            # .str.replace("M3", "000")
            .str.extract(r"(\d+\.?\d*)")
            .astype(float)
        )
        v["廠內最大儲存量(公升)"] = (
            v["廠內最大儲存量(公升)"]
            .astype(str)
            # .str.replace("M3", "000")
            .str.extract(r"(\d+\.?\d*)")
            .astype(float)
        )

        for c, d, n in zip(sort_column, sorted_data, file_name):
            d[k] = (v.groupby([c])[sort_value].sum().reset_index()).sort_values(
                by=sort_value[::-1], ascending=[False, False]
            )

            data_reader.parameters["file_name"] = n
            output_as(d, data_reader.parameters)


if __name__ == "__main__":

    # main()
    high_tech_industry_main()
