from pathlib import Path
from PyPDF2 import PdfMerger
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
from pdfkit import from_string
from src.Entities.Collection_Entry import Collection_Entry
from src.Entities.template_factory import Template_Factory

class Collection_Overview_Creator:

    @staticmethod
    def create_collection_overview(file_path_list):

        root_path = Path(__file__).parents[2]
        pdf_folder_path = root_path.joinpath("resources", "individual_pdfs")
        html_folder_path = root_path.joinpath("resources", "individual_htmls")
        final_pdf_file_path = root_path.joinpath("resources", "Sammlungs_Katalog_2022.pdf")
        image_folder_path = root_path.joinpath("resources", "images")

        art_count = 1
        collectionEntryList = []

        for file_index, file_path in enumerate(file_path_list):

            wb = load_workbook(file_path, data_only=True)
            sheet = wb['KUNSTW_20_B']

            active_rows = wb.active
            max_rows = active_rows.max_row

            # Put your sheet in the loader
            image_loader = SheetImageLoader(sheet)

            row_count = 0
            keep_going = True

            for row in sheet.iter_rows():

                print_string = f"File {file_index + 1} / {len(file_path_list)} - Row {row_count + 1} / {max_rows}"
                print(print_string)

                if row_count > 0 and keep_going:
                    cell_count = 0

                    hochformat = 'O' + str(row_count + 1)
                    querformat = 'P' + str(row_count + 1)
                    quadratisch = 'Q' + str(row_count + 1)

                    image_file_path = str(image_folder_path) + "\\" + str(art_count) + ".jpg"

                    try:
                        hoch_image = image_loader.get(hochformat)
                        hoch_image.save(image_file_path)
                        image_format = "hoch"
                        image_path = image_file_path
                    except Exception as e:
                        try:
                            quer_image = image_loader.get(querformat)
                            quer_image.save(image_file_path)
                            image_format = "quer"
                            image_path = image_file_path
                        except Exception as e:
                            try:
                                quadratisch_image = image_loader.get(quadratisch)
                                quadratisch_image.save(image_file_path)
                                image_format = "quad"
                                image_path = image_file_path
                            except Exception as e:
                                image_format = "no"
                                image_path = "no"

                    for cell in row:

                        if keep_going:

                            if cell_count == 0:
                                if cell.value == None:
                                    keep_going = False
                                else:
                                    id = int(cell.value)

                            elif cell_count == 1:
                                location = cell.value
                            elif cell_count == 2:
                                ken_1 = cell.value
                            elif cell_count == 3:
                                ken_2 = cell.value
                            elif cell_count == 4:
                                count = cell.value
                            elif cell_count == 5:
                                series = cell.value
                            elif cell_count == 6:
                                artist = cell.value
                            elif cell_count == 7:
                                country = cell.value
                            elif cell_count == 8:
                                title = cell.value
                            elif cell_count == 9:
                                technique_count = cell.value
                            elif cell_count == 10:
                                material = cell.value
                            elif cell_count == 11:
                                production = cell.value
                            elif cell_count == 12:
                                frame = cell.value
                            elif cell_count == 13:
                                size = cell.value

                            elif cell_count == 19:
                                edition = cell.value
                            elif cell_count == 20:
                                edition_pp = cell.value
                            elif cell_count == 21:
                                edition_total = cell.value
                            elif cell_count == 22:
                                number = cell.value
                            elif cell_count == 23:
                                signature = cell.value
                            elif cell_count == 24:
                                certificate = cell.value
                            elif cell_count == 25:
                                price = cell.value
                            elif cell_count == 27:
                                seller = cell.value
                            elif cell_count == 28:
                                sale_month = cell.value
                            elif cell_count == 29:
                                sale_year = cell.value
                            elif cell_count == 31:
                                miscellaneous = cell.value

                            cell_count += 1

                    if keep_going:
                        newCollectionEntry = Collection_Entry(id, location, ken_1, ken_2, count, series, artist,
                                                              country, title, technique_count, material, production,
                                                              frame, size, image_format, image_path, edition,
                                                              edition_pp, edition_total, number, signature, certificate,
                                                              price,
                                                              seller, sale_month, sale_year, miscellaneous)
                        collectionEntryList.append(newCollectionEntry)
                        art_count += 1

                row_count += 1

        stop_end = 3
        merger = PdfMerger()
        for count, collectionEntry in enumerate(collectionEntryList):
            print_string = str(count + 1) + "/" + str(len(collectionEntryList))
            print(print_string)

            pdf_file_name = str(count + 1) + ".pdf"
            pdf_file_path = str(pdf_folder_path) + "\\" + pdf_file_name

            html_file_name = str(count + 1) + ".html"
            html_file_path = str(html_folder_path) + "\\" + html_file_name

            html_string = Template_Factory.create_template(collectionEntry, count + 1, len(collectionEntryList))

            with open(html_file_path, "w", encoding="utf-8") as html_file:
                html_file.write(html_string)
            html_file.close()

            from_string(html_string, output_path=pdf_file_path, options={"enable-local-file-access": ""})

            merger.append(pdf_file_path)

            # if count == stop_end:
            #     break

        merger.write(final_pdf_file_path)
        merger.close()

