import sqlite3
import xlwt

PATH_EXTAKE = 'c:/Intake/{}{}.xls'
LAMP = 'h8'
DICT_BREND = {
    'AUDI': 'AUD',
    'BYD': 'BYD',
    'BMW': 'BMW',
    'CHANGAN': 'CHG',
    'CHERY': 'CHY',
    'CHEVROLET': 'CHT',
    'CHRYSLER': 'CHR',
    'CITROEN': 'CIT',
    'CADILLAC': 'CDL',
    'DAEWOO': 'DAW',
    'DODGE': 'DDG',
    'DAIHATSU': 'DHS',
    'FIAT': 'FAT',
    'FORD': 'FRD',
    'GREAT WALL': 'GRW',
    'HONDA': 'HND',
    'HYUNDAI': 'HYN',
    'JAGUAR': 'JAG',
    'JEEP': 'JEP',
    'KIA': 'KIA',
    'LADA': 'LAD',
    'LAND ROVER': 'LNR',
    'LEXUS': 'LEX',
    'LINCOLN': 'LLN',
    'LIFAN': 'LFN',
    'MAZDA': 'MZD',
    'MERCEDES-BENZ': 'MRB',
    'MINI': 'MIN',
    'MITSUBISHI': 'MMC',
    'NISSAN': 'NSN',
    'OPEL': 'OPL',
    'PEUGEOT': 'PGT',
    'INFINITI': 'INF',
    'ISUZU': 'ISZ',
    'PORSCHE': 'PRS',
    'RENAULT': 'RNT',
    'ROVER': 'RVR,',
    'SAAB': 'SAB',
    'SEAT': 'SET',
    'SKODA': 'SKD',
    'SMART': 'SMT',
    'SSANGYONG': 'SSY',
    'SUBARU': 'SBR',
    'SUZUKI': 'SUZ',
    'TOYOTA': 'TOY',
    'VOLKSWAGEN': 'WAG',
    'VOLVO': 'VLV',
    'Volkswagen': 'WAG',
}


def create_brandlist():
    lst = []
    con = sqlite3.connect('db.sqlite')
    cur = con.cursor()
    sql = 'SELECT brand FROM lamps'
    cur.execute(sql)
    for result in cur:
        lst.append(result[0])
    con.commit()
    con.close()
    print(len(set(lst)))
    return set(lst)


def list_gen(list_brand):
    listing_brand = []
    con = sqlite3.connect('db.sqlite')
    cur = con.cursor()
    index = 0
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Список заказов")
    for brand in list_brand:
        sql = 'SELECT model, eyars FROM lamps WHERE brand=? AND lamp=?'
        cur.execute(sql, (brand, LAMP))
        for result in cur:
            listing_brand.append(result[0] + result[1])
        obj = str(set(listing_brand))
        print(brand)
        print(obj)
        index = index + 1
        brand_cut = DICT_BREND[brand]
        print(brand_cut)
        sheet1.write(index, 1, brand_cut)
        sheet1.write(index, 2, brand)
        sheet1.write(index, 3, obj)
        book.save(PATH_EXTAKE.format('Лампы -', LAMP))
        listing_brand = []
    con.commit()
    con.close()


def main():
    list_brand = create_brandlist()
    print(list_brand)
    list_gen(list_brand)


main()
