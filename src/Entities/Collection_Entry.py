
class Collection_Entry:

    """Lfd. N.	Ab.N.	Ken-1	K-2	Anz.	Serie	Künstler	Land	Titel	Technik / Anzahl Einzelstücke	"Materialart /Papierart/-qualität"	"Produktion:
Hersteller / Drucker"	Rahmen	Größe: Br/Hö	hochformat	querformat	quadratisch	q/h	Rand: lXu/o	Auf-lage	Auf-lage AP / PP	Auf-lage Total	Num.	Signatur / Nummer.	Zertif.	Preis: €	Preis: Insge.	Verkäufer	Kaufmon.	Kaufjahr.	Jahr
"""

    def __init__(self, id, location, ken_1, ken_2, count, series, artist, country, title, technique_count, material, production,
                 frame, size, image_format, image_path, edition, edition_pp, edition_total, number, signature, certificate, price,
                 seller, sale_month, sale_year, miscellaneous):
        self.id = id
        self.location = location
        self.ken_1 = ken_1
        self.ken_2 = ken_2
        self.count = count
        self.series = series
        self.artist = artist
        self.country = country
        self.title = title
        self.technique_count = technique_count
        self.material = material
        self.production = production
        self.frame = frame
        self.size = size

        self.image_format = image_format
        self.image_path = image_path

        self.edition = edition
        self.edition_pp = edition_pp
        self.edition_total = edition_total
        self.number = number
        self.signature = signature
        self.certificate = certificate
        self.price = price
        self.seller = seller
        self.sale_month = sale_month
        self.sale_year = sale_year
        self.miscellaneous = miscellaneous
