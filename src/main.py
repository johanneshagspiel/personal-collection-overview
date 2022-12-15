from pathlib import Path
from src.Entities.Collection_Overview_Creator import Collection_Overview_Creator

if __name__ == '__main__':

    root_path = Path(__file__).parents[1]
    file_1_path = root_path.joinpath("resources", "KUNSTW_AbisK_01.2022.xlsx")
    file_2_path = root_path.joinpath("resources", "KUNSTW_LbisZ_01.2022.xlsx")
    file_path_list = [file_1_path, file_2_path]

    Collection_Overview_Creator.create_collection_overview(file_path_list)
