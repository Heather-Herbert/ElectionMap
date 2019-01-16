from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import csv


def simple_get(url):
    """
    Attempts to get the content at `url` by making an HTTP GET request.
    If the content-type of response is some kind of HTML/XML, return the
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    """
    Returns True if the response seems to be HTML, False otherwise.
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200 
            and content_type is not None 
            and content_type.find('html') > -1)


def log_error(e):
    """
    It is always a good idea to log errors. 
    This function just prints them, but you can
    make it do anything.
    """
    print(e)

baseURL = 'https://www.streetcheck.co.uk/postcode/'
with open('postcodes.csv', 'r') as csvFile:
    reader = csv.reader(csvFile)
    items = []
    for row in reader:
        response = simple_get(baseURL + row[0])
        csvWritterFile = open('Demographic.csv', 'a')
        csvWritterFile.write(row[0].encode('utf-8').strip())
        if response is not None:
            html = BeautifulSoup(response, 'html.parser')
            names = set()
            for li in html.select('td'):
                for child in li.descendants:
                    try:
                        csvWritterFile = open('Demographic.csv', 'a')
                        csvWritterFile.write(child.encode('utf-8').strip())
                        csvWritterFile.write(',')
                    except TypeError as e:
                        print e
                    finally:
                        csvWritterFile.close()
        csvWritterFile = open('Demographic.csv', 'a')
        csvWritterFile.write("\n")
        csvWritterFile.close()
