from pathlib import Path
from pdfkit import from_string, from_file
from src.Entities.Collection_Overview_Creator import Collection_Overview_Creator

if __name__ == '__main__':

    root_path = Path(__file__).parents[1]
    # # file_1_path = root_path.joinpath("resources", "KUNSTW_AbisG_1v3_03.2023.xlsx")
    # # file_2_path = root_path.joinpath("resources", "KUNSTW_HbisO_2v3_03.2023.xlsx")
    # # file_3_path = root_path.joinpath("resources", "KUNSTW_PbisZ__3v3_03.2023.xlsx")
    # # file_path_list = [file_1_path, file_2_path, file_3_path]

    file_path = root_path.joinpath("resources", "Test.xlsx")
    file_path_list = [file_path]

    Collection_Overview_Creator.create_collection_overview(file_path_list)

    # from_file(input="C:/Users/Johannes/Desktop/programs/personal_projects/personal-collection-overview/resources/print/individual_htmls/1.html",
    #             output_path="test.pdf",
    #             options={"enable-local-file-access": "",
    #                      'page-size':'A4',
    #                      'encoding':'utf-8',
    #                      'margin-top':'0cm',
    #                      'margin-bottom':'0cm',
    #                      'margin-left':'0cm',
    #                      'margin-right':'0cm',
    #                      'dpi': 400})

