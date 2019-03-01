import xlrd
import json

# Data comes from https://github.com/earthchie/jquery.Thailand.js/blob/master/jquery.Thailand.js/database/raw_database
workbook = xlrd.open_workbook('original_database_from_thaipost.xls')
worksheet = workbook.sheet_by_index(0)

# Province
province_id_column = 0
province_name_column = 1

district_id_column = 2
district_name_column = 3

city_id_column = 5
city_name_column = 6
zip_code_column = 7

province_data = []
province_map = {}
district_data = []
district_map = {}
city_data = []
city_map = {}

province_count = 1
district_count = 1
city_count = 1

for i in range(1, worksheet.nrows):
  # Get province data
  province_id = worksheet.cell(i, province_id_column).value
  province_name = worksheet.cell(i, province_name_column).value
  if province_id not in province_map:
    province_map[province_id] = province_count
    province_data.append({
      "model": "common.province",
      "pk": province_count,
      "fields": {
        "name": province_name,
        "active": True
      }
    })
    province_count += 1

  # Get district data
  district_id = worksheet.cell(i, district_id_column).value
  district_name = worksheet.cell(i, district_name_column).value
  if district_id not in district_map:
    district_map[district_id] = district_count
    district_data.append({
      "model": "common.district",
      "pk": district_count,
      "fields": {
        "name": district_name,
        "province": province_map[province_id],
        "active": True
      }
    })
    district_count += 1

  # Get city data
  city_id = worksheet.cell(i, city_id_column).value
  city_name = worksheet.cell(i, city_name_column).value
  zip_code = worksheet.cell(i, zip_code_column).value
  if city_id not in city_map:
    city_map[city_id] = city_count
    city_data.append({
      "model": "common.city",
      "pk": city_count,
      "fields": {
        "name": city_name,
        "district": district_map[district_id],
        "zip_code": zip_code,
        "active": True
      }
    })
    city_count += 1
  

with open('province.json', 'w') as outfile:
    json.dump(province_data, outfile, ensure_ascii=False, indent=4)

with open('district.json', 'w') as outfile:
    json.dump(district_data, outfile, ensure_ascii=False, indent=4)

with open('city.json', 'w') as outfile:
    json.dump(city_data, outfile, ensure_ascii=False, indent=4)