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
    except:
        return None


def get_language(content):
    try:
        return content.select(".rc-Language")[0].get_text()
    except:
        return None


def get_start_date(content):
    try:
        return content.find(id="start-date-string").span.get_text()
    except:
        return None


def get_week_count(content):
    try:
        return len(content.select(".week"))
    except:
        return None


def get_rating(content):
    try:
        return content.select(".ratings-text")[0].span.get_text()
    except:
        return None


def fetch_courses_feed():
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    feed = requests.get(url).text
    return xml.fromstring(feed)


def get_courses(feed, urls_to_handle=3):
    urls = [child[0].text for child in feed[:urls_to_handle]]
    courses = []
    for url in urls:
        course = {}
        page = fetch_page(url)
        print("{} fetched".format(url))
        page_content = web(page, "html.parser")
        course["title"] = get_title(page_content)
        course["start_date"] = get_start_date(page_content)
        course["week_count"] = get_week_count(page_content)
        course["avg_rating"] = get_rating(page_content)
        course["language"] = get_language(page_content)
        print("{} parsed".format(url))
        courses.append(course)
    return courses


def fetch_page(link):
    headers = {"Accept-Language": "ru"}
    res = requests.get(link, headers=headers)
    res.encoding = "utf-8"
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
        exit("File exists")
    workbook.save(out_path)
    print("{} saved".format(out_path))


def main():
    feed = fetch_courses_feed()
    courses = get_courses(feed)
    workbook = get_filled_workbook(courses)
    filename = get_arguments().filename
    save_in_excel(workbook, filename)


if __name__ == "__main__":
    main()
