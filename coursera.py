import requests
import openpyxl
import argparse
import os
import xml.etree.ElementTree as xml
from bs4 import BeautifulSoup as web
from openpyxl.styles import Font


def get_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", required=True, dest="filename")
    args = parser.parse_args()
    return args


def get_title(content):
    try:
        return content.select(".title")[0].get_text()
    except IndexError:
        return None


def get_language(content):
    try:
        return content.select(".rc-Language")[0].get_text()
    except IndexError:
        return None


def get_start_date(content):
    try:
        return content.find(id="start-date-string").span.get_text()
    except AttributeError:
        return None


def get_week_count(content):
    return len(content.select(".week"))


def get_rating(content):
    try:
        return content.select(".ratings-text")[0].span.get_text()
    except IndexError:
        return None


def fetch_courses_feed():
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    feed = requests.get(url).text
    return xml.fromstring(feed)


def get_parsed_course(page):
    course = {}
    page_content = web(page, "html.parser")
    course["title"] = get_title(page_content)
    course["start_date"] = get_start_date(page_content)
    course["week_count"] = get_week_count(page_content)
    course["avg_rating"] = get_rating(page_content)
    course["language"] = get_language(page_content)
    return course


def fetch_page(url):
    headers = {"Accept-Language": "ru"}
    res = requests.get(url, headers=headers)
    res.encoding = "utf-8"
    print("{} fetched".format(url))
    return res.text


def get_filled_workbook(courses):
    wb = openpyxl.Workbook()
    sheet = wb.active
    row = sheet.row_dimensions[1]
    row.font = Font(size=12, bold=True)
    sheet.title = "Список курсов"
    sheet.append((
        "Название курса",
        "Начало курса",
        "Кол.Недель",
        "Рейтинг",
        "Язык",
    ))
    for course in courses:
        sheet.append(list(course.values()))
    return wb


def save_in_excel(workbook, filename):
    out_path = "{}{}".format(filename, ".xlsx")
    if os.path.exists(out_path):
        return False
    workbook.save(out_path)
    return True


def main():
    feed = fetch_courses_feed()
    urls_to_handle = 20
    urls = [child[0].text for child in feed[:urls_to_handle]]
    pages = (fetch_page(url) for url in urls)
    courses = (get_parsed_course(page) for page in pages)
    workbook = get_filled_workbook(courses)
    filename = get_arguments().filename
    if save_in_excel(workbook, filename):
        print("Saved in {}.xlsx!".format(filename))
    else:
        exit("File exists")


if __name__ == "__main__":
    main()
