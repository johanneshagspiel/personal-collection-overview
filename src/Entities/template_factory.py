from datetime import datetime
from pathlib import Path
from PIL import Image
from jinja2 import Environment, FileSystemLoader

class Template_Factory:
    root_path = Path(__file__).parents[2]
    file_path = root_path.joinpath("resources", "templates")

    file_loader = FileSystemLoader(file_path)
    env = Environment(loader=file_loader)

    MAX_HEIGHT = 400
    MAX_WIDTH = 900

    def __init__(self):
        pass

    @classmethod
    def create_print_full_page(cls, collectionEntry):

        css_path_str = "file:///" + str(cls.file_path) + "\\print\\fullPage.css"
        template = cls.env.get_template('fullPage.html')

        image_format_str = collectionEntry.image_format
        image_path = "file:///" + str(collectionEntry.image_path)

        width = 0
        height = 0

        if image_format_str != "no":
            im = Image.open(str(collectionEntry.image_path))
            width, height = im.size

            height_coefficient = cls.MAX_HEIGHT / height
            height = cls.MAX_HEIGHT
            width = int(height_coefficient * width)

            if width > cls.MAX_WIDTH:
                width_coefficient = cls.MAX_WIDTH / width
                width = cls.MAX_WIDTH
                height = int(width_coefficient * height)

        artist_str = collectionEntry.artist
        if not artist_str:
            artist_str = "Unbekannt"

        title_str = collectionEntry.title
        if not title_str:
            title_str = "Unbekannt"

        edition_str = collectionEntry.edition
        if not edition_str:
            edition_str = "0"
        edition_pp_str = collectionEntry.edition_pp
        if not edition_pp_str:
            edition_pp_str = "0"
        edition_total_str = collectionEntry.edition_total
        if not edition_total_str:
            edition_total_str = "0"
        total_edition_str = f"{edition_str} / {edition_pp_str} / {edition_total_str}"

        number_str = collectionEntry.number
        if not number_str:
            number_str = "Unbekannt"

        technique_str = collectionEntry.technique_count
        if not technique_str:
            technique_str = "Unbekannt"

        seller_str = collectionEntry.seller
        if not seller_str:
            seller_str = "Unbekannt"

        material_str = collectionEntry.material
        if not material_str:
            material_str = "Unbekannt"

        sale_month_str = collectionEntry.sale_month
        if sale_month_str:
            if isinstance(sale_month_str, str):
                print(sale_month_str)
            else:
                sale_month_str = str(sale_month_str.day) + "-" +  str(sale_month_str.year) + "-" + str(sale_month_str.month)
        else:
            sale_month_str = "Unbekannt"

        production_str = collectionEntry.production
        if not production_str:
            production_str = "Unbekannt"

        price_str = collectionEntry.price
        if not price_str:
            price_str = "Unbekannt"

        size_str = collectionEntry.size
        if not size_str:
            size_str = "Unbekannt"

        ken_1_str = collectionEntry.ken_1
        ken_2_str = collectionEntry.ken_2
        inventory_number_str = str(ken_1_str) + " " + str(ken_2_str)
        if not ken_1_str or not ken_2_str:
            inventory_number_str = "Unbekannt"

        count_str = collectionEntry.count
        if not count_str:
            count_str = "Unbekannt"

        location_str = collectionEntry.location
        if not location_str:
            location_str = "Unbekannt"

        miscellaneous_str = collectionEntry.miscellaneous
        if not miscellaneous_str:
            miscellaneous_str = "Unbekannt"

        id_str = collectionEntry.id
        if not id_str:
            id_str = "Unbekannt"

        print_date_str = datetime.now().strftime("%d.%m.%Y")


        html_output = template.render(css_path=css_path_str,

                                      artist=artist_str,
                                      title=title_str,

                                      image_path=image_path,
                                      height=height,
                                      width=width,

                                      edition=total_edition_str,
                                      number=number_str,
                                      technique=technique_str,
                                      seller=seller_str,
                                      material=material_str,
                                      sale_month=sale_month_str,
                                      production=production_str,
                                      price=price_str,
                                      size=size_str,
                                      inventory_number=inventory_number_str,
                                      count=count_str,
                                      location=location_str,

                                      miscellaneous=miscellaneous_str,

                                      print_date=print_date_str,
                                      id=id_str,
                                      )

        return html_output

    @classmethod
    def create_print_image_page(cls, collectionEntry):

        css_path_str = "file:///" + str(cls.file_path) + "\\print\\imagePage.css"
        template = cls.env.get_template('imagePage.html')

        image_format_str = collectionEntry.image_format
        image_path = "file:///" + str(collectionEntry.image_path)

        maxHeight = 770
        maxWidth = 550

        width = 0
        height = 0

        if image_format_str != "no":
            im = Image.open(str(collectionEntry.image_path))
            width, height = im.size

            height_coefficient = maxHeight / height
            height = maxHeight
            width = int(height_coefficient * width)

            if width > maxWidth:
                width_coefficient = maxWidth / width
                width = maxWidth
                height = int(width_coefficient * height)

        artist_str = collectionEntry.artist
        if not artist_str:
            artist_str = "Unbekannt"

        title_str = collectionEntry.title
        if not title_str:
            title_str = "Unbekannt"

        ken_1_str = collectionEntry.ken_1
        ken_2_str = collectionEntry.ken_2
        inventory_number_str = str(ken_1_str) + " " + str(ken_2_str)
        if not ken_1_str or not ken_2_str:
            inventory_number_str = "Unbekannt"

        html_output = template.render(css_path=css_path_str,

                                      image_path=image_path,
                                      height=height,
                                      width=width,

                                      artist=artist_str,
                                      title=title_str,
                                      inventory_number=inventory_number_str)

        return html_output
    @classmethod
    def create_individual_template(cls, mode, collectionEntry, page_count, page_total):

        if mode == "print":
            css_path_str = "file:///" + str(cls.file_path) + "\\print\\fullPage.css"
            template = cls.env.get_template('print_base.html')
        elif mode == "website":
            css_path_str = "file:///" + str(cls.file_path) + "\\base.css"
            template = cls.env.get_template('base.html')

        id_str = collectionEntry.id
        if not id_str:
            id_str = "Unbekannt"

        location_str = collectionEntry.location
        if not location_str:
            location_str = "Unbekannt"

        ken_1_str = collectionEntry.ken_1
        if not ken_1_str:
            ken_1_str = "Unbekannt"

        ken_2_str = collectionEntry.ken_2
        if not ken_2_str:
            ken_2_str = "Unbekannt"

        count_str = collectionEntry.count
        if not count_str:
            count_str = "Unbekannt"

        series_str = collectionEntry.series
        if not series_str:
            series_str = "Unbekannt"

        artist_str = collectionEntry.artist
        if not artist_str:
            artist_str = "Unbekannt"

        title_str = collectionEntry.title
        if not title_str:
            title_str = "Unbekannt"

        country_str = collectionEntry.country
        if not country_str:
            country_str = "Unbekannt"

        technique_count_str = collectionEntry.technique_count
        if not technique_count_str:
            technique_count_str = "Unbekannt"

        material_str = collectionEntry.material
        if not material_str:
            material_str = "Unbekannt"

        production_str = collectionEntry.production
        if not production_str:
            production_str = "Unbekannt"

        frame_str = collectionEntry.frame
        if not frame_str:
            frame_str = "Unbekannt"

        size_str = collectionEntry.size
        if not size_str:
            size_str = "Unbekannt"

        edition_str = collectionEntry.edition
        if not edition_str:
            edition_str = "Unbekannt"

        edition_pp_str = collectionEntry.edition_pp
        if not edition_pp_str:
            edition_pp_str = "Unbekannt"

        edition_total_str = collectionEntry.edition_total
        if not edition_total_str:
            edition_total_str = "Unbekannt"

        number_str = collectionEntry.number
        if not number_str:
            number_str = "Unbekannt"

        signature_str = collectionEntry.signature
        if not signature_str:
            signature_str = "Unbekannt"

        certificate_str = collectionEntry.certificate
        if not certificate_str:
            certificate_str = "Unbekannt"

        price_str = collectionEntry.price
        if not price_str:
            price_str = "Unbekannt"

        seller_str = collectionEntry.seller
        if not seller_str:
            seller_str = "Unbekannt"

        sale_month_str = collectionEntry.sale_month
        if sale_month_str:
            if isinstance(sale_month_str, str):
                print(sale_month_str)
            else:
                sale_month_str = str(sale_month_str.day) + "-" +  str(sale_month_str.year) + "-" + str(sale_month_str.month)
        else:
            sale_month_str = "Unbekannt"

        miscellaneous_str = collectionEntry.miscellaneous
        if not miscellaneous_str:
            miscellaneous_str = "Unbekannt"

        image_format_str = collectionEntry.image_format
        image_path = "file:///" + str(collectionEntry.image_path)

        width = 0
        height = 0

        if image_format_str != "no":
            im = Image.open(str(collectionEntry.image_path))
            width, height = im.size

            height_coefficient = cls.MAX_HEIGHT / height
            height = cls.MAX_HEIGHT
            width = int(height_coefficient * width)

            if width > cls.MAX_WIDTH:
                width_coefficient = cls.MAX_WIDTH / width
                width = cls.MAX_WIDTH
                height = int(width_coefficient * height)


        html_output = template.render(css_path=css_path_str,
                                      id=id_str,
                                      location=location_str,
                                      ken_1=ken_1_str,
                                      ken_2=ken_2_str,
                                      count=count_str,
                                      series=series_str,
                                      artist=artist_str,
                                      title=title_str,
                                      country=country_str,
                                      technique_count=technique_count_str,
                                      material=material_str,
                                      production=production_str,
                                      frame=frame_str,
                                      size=size_str,
                                      edition=edition_str,
                                      edition_pp=edition_pp_str,
                                      edition_total=edition_total_str,
                                      number=number_str,
                                      signature=signature_str,
                                      certificate=certificate_str,
                                      price=price_str,
                                      seller=seller_str,
                                      sale_month=sale_month_str,
                                      miscellaneous=miscellaneous_str,
                                      image_format_str=image_format_str,
                                      image_path=image_path,
                                      width=width,
                                      height=height,
                                      page_count=page_count,
                                      page_total=page_total)

        return html_output
