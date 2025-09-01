import os
import pandas as pd
import utils.read_data as read_data
from utils.output_excel import output_as
from collections import defaultdict

parameters = {
    "read_all_sheets": True,  # Set to True if you want to read all sheets in each Excel file
    "sheet_names": "公共危險物品運作資料",  # Specify the sheet name if needed
    "path_data": "../Data/科技廠救災能量",  # Specify the path to your Excel files
    # "path_data": "../Test",  # For testing purposes, you can change this to a test folder
    "pattern": "top_ten_operating_chemicals",  # You can specify a pattern if needed, ['default', '苗栗縣', top_ten_operating_chemicals, sort_by_location, industry_rescue_equipment]
    "path_output": "/Output",  # Specify the output folder
    # "output_path": "Output",  # Specify the output path
    "file_name": "Aggregated_data.xlsx",  # Specify the output file name
}


def main():
    # Create an instance of the read_data class
    data_reader = read_data.read_data(parameters)

    # Access the combined DataFrame
    path = data_reader.get_path()
    combined_data = data_reader.read_excel_files()
    output_as(combined_data, data_reader.parameters)

    # for file, df in combined_data.items():
    #     print(f"Data from {file}:")
    #     print(df.head(5))  # Display the first few rows of each sheet
    #     print(df.keys())


def high_tech_industry_chems_main():
    """Function to handle high-tech industry data processing."""
    data_reader = read_data.read_data(parameters)
    folder_path = data_reader.get_path()
    data_reader.parameters["output_path"] = (
        folder_path + data_reader.parameters["path_output"]
    )
    # print(f"Folder path: {folder_path}")

    folders = os.listdir(folder_path)
    if "Output" in folders:
        folders.remove("Output")  # Remove the Output folder if it exists

    for folder in folders:
        file_path = os.path.join(folder_path, folder)
        data_reader.parameters["folder_path"] = file_path
        print(f"Processing folder: {folder}")
        data_reader.parameters["file_name"] = folder + ".xlsx"
        # Read the Excel file
        print(f"Reading with pattern: {data_reader.parameters['pattern']}")

        combined_data = defaultdict(list)
        for f, (k, v) in data_reader.read_excel_files():
            if len(k) == 0:
                print(f"No data found in {f}. Skipping.")
            else:
                combined_data[k[0]].append(v)
        combined_data = {
            k: pd.concat(v, ignore_index=True) for k, v in combined_data.items()
        }
        output_as(
            combined_data, data_reader.parameters
        )  # Output the combined data to an Excel file
    print("All folders processed successfully.")
    ###########    sort required data

    data_reader.parameters["path_data"] = (
        data_reader.parameters["path_data"] + data_reader.parameters["path_output"]
    )
    data_reader.parameters["folder_path"] = data_reader.get_path()
    data_reader.parameters["pattern"] = "sort_by_location"
    data_reader.parameters["file_name"] = "Sorted_data.xlsx"
    sorted_data = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}
    sorted_factory = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}

    for f, (r, v) in data_reader.read_excel_files():
        print(f"Processing file: {f}", r)
        sorted_data[r].extend(v)
        sorted_factory[r].append(f)
        print(sum(i.shape[0] for i in sorted_data[r]))

    print("Data sorted by location:", sorted_data.keys())
    print("Factory sorted by location:", sorted_factory)

    sorted_data = {
        k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
        for k, v in sorted_data.items()
    }
    output_as(sorted_data, data_reader.parameters)

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
    sorted_data = [{}, {}, {}]
    file_name = ["sort_by_hazmat.xlsx", "sort_by_container.xlsx", "sort_by_state.xlsx"]

    for k, v in zip(keys, values):
        if k == "其他":
            continue
        else:
            v["化學物質名稱"] = v["化學物質名稱"].astype(str)
            v["容器材質"] = v["容器材質"].astype(str)
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


def high_tech_industry_rescue_equipment_main():
    """
    Function to handle high-tech industry rescue equipment data processing.
    """
    data_reader = read_data.read_data(parameters)

    folder_path = data_reader.get_path()
    print(f"Folder path: {folder_path}")
    data_reader.parameters["output_path"] = (
        folder_path + data_reader.parameters["path_output"] + "/Rescue_equipment"
    )
    data_reader.parameters["pattern"] = "industry_rescue_equipment"

    # print(f"Folder path: {folder_path}")
    folders = os.listdir(folder_path)
    if "Output" in folders:
        folders.remove("Output")  # Remove the Output folder if it exists

    for folder in folders:
        file_path = os.path.join(folder_path, folder)
        data_reader.parameters["folder_path"] = file_path
        print(f"Processing folder: {folder}")
        data_reader.parameters["file_name"] = folder + ".xlsx"
        # Read the Excel file
        print(f"Reading with pattern: {data_reader.parameters['pattern']}")

        combined_data = defaultdict(list)
        for f, (k, v) in data_reader.read_excel_files():
            if len(k) == 0:
                print(f"No data found in {f}. Skipping.")
            else:
                combined_data[k[0]].append(v)
        combined_data = {
            k: pd.concat(v, ignore_index=True) for k, v in combined_data.items()
        }
        output_as(
            combined_data, data_reader.parameters
        )  # Output the combined data to an Excel file
    print("All folders processed successfully.")
    ###################
    """Sort data"""
    data_reader.parameters["path_data"] = (
        data_reader.parameters["path_data"]
        + data_reader.parameters["path_output"]
        + "/Rescue_equipment"
    )
    data_reader.parameters["folder_path"] = data_reader.get_path()
    data_reader.parameters["pattern"] = "sort_by_location"
    data_reader.parameters["file_name"] = "Sorted_data.xlsx"
    sorted_data = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}
    sorted_factory = {"北部園區": [], "中部園區": [], "南部園區": [], "其他": []}

    for f, (r, v) in data_reader.read_excel_files():
        print(f"Processing file: {f}", r)
        sorted_data[r].extend(v)
        sorted_factory[r].append(f)
        print(sum(i.shape[0] for i in sorted_data[r]))

    print("Data sorted by location:", sorted_data.keys())
    print("Factory sorted by location:", sorted_factory)

    sorted_data = {
        k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
        for k, v in sorted_data.items()
    }
    output_as(sorted_data, data_reader.parameters)

    data_reader.parameters["path_data"] = (
        (data_reader.parameters["path_data"] + data_reader.parameters["path_output"])
        if not (
            data_reader.parameters["path_output"] in data_reader.parameters["path_data"]
        )
        else data_reader.parameters["path_data"]
    )
    ################# Analyze data

    data_reader.parameters["path_data"] = (
        (
            data_reader.parameters["path_data"]
            + data_reader.parameters["path_output"]
            + "/Rescue_equipment"
        )
        if not data_reader.parameters["path_output"]
        in data_reader.parameters["path_data"]
        else data_reader.parameters["path_data"]
    )

    data_reader.parameters["folder_path"] = data_reader.get_path()
    data_reader.parameters["read_all_sheets"] = True
    keys, values = data_reader.read_one_excel(
        data_reader.parameters["folder_path"] + "/Sorted_data.xlsx"
    )

    sort_column = ["證照", "演練", "應變設備"]
    sort_value = [["證照數量"], ["演練數量"], ["應變設備數量", "應變設備可支援數量"]]
    sorted_data = [{}, {}, {}]
    file_name = [
        "sort_by_certificate.xlsx",
        "sort_by_training.xlsx",
        "sort_by_equipment.xlsx",
    ]

    for k, v in zip(keys, values):
        if k == "其他":
            continue
        else:
            v["證照"] = v["證照"].astype(str)
            v["演練"] = v["演練"].astype(str)
            v["應變設備"] = v["應變設備"].astype(str)

            # Convert "廠內最大儲存量(公斤)" to float, using pandas string method, str.extract().
            v["證照數量"] = (
                v["證照數量"].astype(str).str.extract(r"(\d+)").astype(float)
            )
            v["演練數量"] = (
                v["演練數量"].astype(str).str.extract(r"(\d+)").astype(float)
            )
            v["應變設備數量"] = (
                v["應變設備數量"].astype(str).str.extract(r"(\d+)").astype(float)
            )
            v["應變設備可支援數量"] = (
                v["應變設備可支援數量"].astype(str).str.extract(r"(\d+)").astype(float)
            )

            for c, s, d, n in zip(sort_column, sort_value, sorted_data, file_name):
                print(k, s)
                d[k] = v.groupby([c])[s].sum().reset_index()

                data_reader.parameters["file_name"] = n
                output_as(d, data_reader.parameters)


if __name__ == "__main__":

    # main()
    high_tech_industry_chems_main()
    high_tech_industry_rescue_equipment_main()
