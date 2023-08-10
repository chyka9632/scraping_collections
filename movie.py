from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movies"
sheet.append(['Movie Rank', 'Movie Name', 'Year of Release', 'IMDB Rating'])
print(excel.sheetnames)

try:
    url = "https://www.imdb.com/chart/top/"
    header = {
        "User-Agent": "Mozilla/5.0 (Windows NT 6.0; Win64; x64) AppleWebKit/601.34 (KHTML, like Gecko) "
                      "Chrome/48.0.2702.358 Safari/601 "
    }
    response = requests.get(url, headers=header).text

    soup = BeautifulSoup(response, "html.parser")

    movies = soup.find('ul', class_='ipc-metadata-list ipc-metadata-list--dividers-between sc-3a353071-0 wTPeg '
                                    'compact-list-view ipc-metadata-list--base').find_all(
        "li", class_="ipc-metadata-list-summary-item sc-bca49391-0 eypSaE cli-parent")

    for movie in movies:
        rank = movie.find('h3', class_='ipc-title__text').text.split(".")[0] + "."  # formatted_rank = f"{rank}."
        name = movie.find('h3', class_='ipc-title__text').text.split(".")[1]
        year = movie.find('div', class_="sc-14dd939d-5 cPiUKY cli-title-metadata").span.text
        rating = movie.find('div', class_="sc-951b09b2-0 hDQwjv sc-14dd939d-2 fKPTOp cli-ratings-container").span.text

        sheet.append([rank, name, year, rating])
        print(rank, name, year, rating)

except Exception as e:
    print(e)

excel.save('IMDB Top 250 Movies.xlsx')
