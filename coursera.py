import requests
import openpyxl
import xml.etree.ElementTree as xml
from bs4 import BeautifulSoup as web
from openpyxl.styles import Font


def get_title(content):
    class1 = ".title"
    if content.select(class1):
        return content.select(class1)[0].get_text()
    elif content.select(class2):
        return content.select(class2)[0].get_text()
    else:
        return "Noname"


def get_start_date(content):
    if content.find(id="start-date-string"):
        return content.find(id="start-date-string").span.get_text()
    return "Without Date"


def get_week_count(content):
    if content.select(".week"):
        return len(content.select(".week"))
    else:
        return 0


def get_rating(content):
    temp = content.select(".ratings-text")
    if temp:
        return temp[0].span.get_text()
    else:
        return "Without Rating"


def fetch_courses_info():
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    feed = requests.get(url).text
    return xml.fromstring(feed)


def get_links(parent):
    return [child[0].text for child in parent]


def handle_data(feed, links_to_handle=20):
    root = feed
    links = get_links(root[:links_to_handle])
    courses_data = get_courses_data(links)
    return courses_data


def get_courses_data(links):
    courses = []
    for link in links:
        course = {}
        page = fetch_page(link)
        page_content = web(page, 'html.parser')
        course["title"] = get_title(page_content)
        course["start_date"] = get_start_date(page_content)
        course["week_count"] = get_week_count(page_content)
        course["avg_rating"] = get_rating(page_content)
        courses.append(course)
    return courses


def fetch_page(link):
    headers = {"Accept-Language": "ru"}
    res = requests.get(link, headers=headers)
    res.encoding = "utf-8"
    return res.text


def save_in_excel(courses, file_name):
    wb = openpyxl.Workbook()
    sheet = wb.active
    row = sheet.row_dimensions[1]
    row.font = Font(size=12, bold=True)
    sheet.title = "Список курсов"
    sheet["A1"] = "Название курса"
    sheet["B1"] = "Начало курса"
    sheet["C1"] = "Кол.Недель"
    sheet["D1"] = "Рейтинг"
    for index in range(len(courses)):
        sheet.cell(column=1, row=index+2, value=courses[index]["title"])
        sheet.cell(column=2, row=index+2, value=courses[index]["start_date"])
        sheet.cell(column=3, row=index+2, value=courses[index]["week_count"])
        sheet.cell(column=4, row=index+2, value=courses[index]["avg_rating"])
    wb.save(file_name)


def main():
    feed = fetch_courses_info()
    data = handle_data(feed)
    save_in_excel(data, "courses.xlsx")


if __name__ == "__main__":
    main()
